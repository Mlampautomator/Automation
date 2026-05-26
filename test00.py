#!/usr/bin/env python3
"""
Bot do zbierania cen promocji/wyprzedaży — wersja async (bez Selenium).
polskielampy.pl serwuje statyczny HTML, więc Chrome jest zbędny.

Wymagania: pip install aiohttp beautifulsoup4 openpyxl pandas lxml
"""

import re
import time
import asyncio
import aiohttp
import pandas as pd
from bs4 import BeautifulSoup
from openpyxl import Workbook
from datetime import datetime

# ── Konfiguracja ───────────────────────────────────────────────────────────────
PLIK_LINKI  = "links.xlsx"
PLIK_WYNIKI = "wyniki_bot.xlsx"
MAX_STRONY  = 10
CONCURRENCY = 60        # zwiększono z 40 do 60 dla szybszego pobierania
TIMEOUT     = 5         # zmniejszono z 8 do 5 sekund na request
OPOZNIENIE  = 0.0       # brak sztucznego opóźnienia
MAX_RETRY   = 2         # zmniejszono z 3 do 2 prób przy błędzie

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9",
    "Accept-Encoding": "gzip, deflate, br",
}

# ── Pobieranie HTML ────────────────────────────────────────────────────────────

async def fetch(session: aiohttp.ClientSession, url: str):
    for attempt in range(1, MAX_RETRY + 1):
        try:
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=TIMEOUT)) as r:
                if r.status == 429:
                    wait = 2 ** attempt
                    print(f"  429 rate-limit, czekam {wait}s... ({url})")
                    await asyncio.sleep(wait)
                    continue
                r.raise_for_status()
                html = await r.text()
                return BeautifulSoup(html, "lxml")
        except asyncio.TimeoutError:
            print(f"  timeout (proba {attempt}/{MAX_RETRY}): {url}")
            await asyncio.sleep(attempt)
        except Exception as e:
            print(f"  x (proba {attempt}/{MAX_RETRY}) {url}: {e}")
            await asyncio.sleep(attempt)
    return None

# ── Parsowanie strony produktu ─────────────────────────────────────────────────

def parsuj_produkt(soup, url, nazwa_lista, cena_lista):
    """
    polskielampy.pl umieszcza kluczowe dane w meta tagach Open Graph.
    """
    def meta(prop):
        tag = soup.find("meta", property=prop) or soup.find("meta", attrs={"property": prop})
        return (tag or {}).get("content", "").strip()

    def wyciagnij_cene(tekst):
        m = re.search(r"(\d[\d\s]*[,.]\d{2})", tekst)
        return m.group(1).replace(" ", "") if m else ""

    # Nazwa
    h1 = soup.find("h1")
    nazwa = h1.get_text(strip=True) if h1 else meta("og:title")

    # Cena promocyjna — meta product:price:amount (statyczne, zawsze obecne)
    cena = meta("product:price:amount")
    if not cena:
        for sel in [".CenaAktualna", ".CenaPromocyjna", ".price-promo", ".price-sale", ".price"]:
            el = soup.select_one(sel)
            if el:
                cena = wyciagnij_cene(el.get_text())
                if cena:
                    break

    # Cena regularna (przed promocją)
    cena_regularna = ""
    # 1. meta tag product:original_price:amount (niektóre sklepy go używają)
    cena_regularna = meta("product:original_price:amount")
    if not cena_regularna:
        # 2. Typowe selektory CSS dla ceny przekreślonej / oryginalnej
        for sel in [
            "#CenaPoprzednia strong", "#CenaPoprzednia", ".CenaStara", ".CenaPrzekr", ".cena-stara", ".cena-przekr",
            ".price-old", ".price-before", ".price-regular", ".PriceOld",
            "del .price", "s .price", ".old-price", ".regular-price",
            "[class*='OldPrice']", "[class*='old_price']", "[class*='StrikePrice']",
        ]:
            el = soup.select_one(sel)
            if el:
                cena_regularna = wyciagnij_cene(el.get_text())
                if cena_regularna:
                    break
    if not cena_regularna:
        # 2b. Szukaj bezpośrednio w sekcji CenaPoprzednia
        blok = soup.select_one("#CenaPoprzednia")
        if blok:
            cena_regularna = wyciagnij_cene(blok.get_text())
    if not cena_regularna:
        # 2c. Szukaj tekstowo wzorca Cena katalogowa / Cena przed
        text = soup.get_text(" ", strip=True)
        match = re.search(r"(?:Cena katalogowa|Cena przed promocją|Cena przed promocją|Cena przed|Cena przed promocja|Cena katalogowa:)[:\s]*(\d[\d\s]*[,.]\d{2})", text, re.I)
        if match:
            cena_regularna = match.group(1).replace(" ", "")
    if not cena_regularna:
        # 3. Znaczniki semantyczne HTML: <del>, <s> z liczbą
        for tag_name in ["del", "s"]:
            for el in soup.find_all(tag_name):
                val = wyciagnij_cene(el.get_text())
                if val:
                    cena_regularna = val
                    break
            if cena_regularna:
                break
    if not cena_regularna:
        # 4. JSON-LD schema.org (Product -> offers -> priceSpecification)
        import json
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                offers = data.get("offers", {})
                if isinstance(offers, list):
                    offers = offers[0]
                specs = offers.get("priceSpecification", [])
                if isinstance(specs, list):
                    for spec in specs:
                        if "regular" in spec.get("@type", "").lower() or spec.get("priceType", "") == "regular":
                            cena_regularna = str(spec.get("price", ""))
                            break
                if not cena_regularna:
                    # Szukaj pola highPrice (cena max = regularna przy promocji)
                    val = offers.get("highPrice", "")
                    if val:
                        cena_regularna = str(val)
            except Exception:
                pass
            if cena_regularna:
                break

    # EAN — polskielampy wyświetla "Kod EAN: XXXXX" w specyfikacji produktu
    ean = ""
    for tag in soup.find_all(string=re.compile(r"Kod EAN", re.IGNORECASE)):
        m = re.search(r"(\d{8,13})", str(tag))
        if m:
            ean = m.group(1)
            break
    if not ean:
        # szukaj w rodzicu tagu "Kod EAN"
        for tag in soup.find_all(string=re.compile(r"Kod EAN", re.IGNORECASE)):
            parent = tag.parent
            if parent:
                sibling_text = parent.find_next_sibling()
                if sibling_text:
                    m = re.search(r"(\d{8,13})", sibling_text.get_text())
                    if m:
                        ean = m.group(1)
                        break
    if not ean:
        for sel in ["#KodEan strong[itemprop='gtin13']", "#KodEan strong", "#KodEan", "div.TbPoz strong[itemprop='gtin13']", "[itemprop='gtin13']", "[data-ean]"]:
            el = soup.select_one(sel)
            if el:
                m = re.search(r"(\d{8,13})", el.get("content", el.get_text()))
                if m:
                    ean = m.group(1)
                    break
    if not ean:
        element = soup.find(id="KodEan") or soup.find(class_=re.compile(r"KodEan|TbPoz", re.I))
        if element:
            m = re.search(r"(\d{8,13})", element.get_text())
            if m:
                ean = m.group(1)
    if not ean:
        # 2b. Szukaj tekstowo wzorca Kod EAN / gtin13 / GTIN
        text = soup.get_text(" ", strip=True)
        match = re.search(r"(?:Kod\s*EAN|EAN|GTIN13|GTIN)[:\s]*([0-9]{8,13})", text, re.I)
        if match:
            ean = match.group(1)
    if not ean:
        # JSON-LD gtin / gtin13 / gtin8
        import json
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                def find_gtin(obj):
                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            if k.lower() in {"gtin13", "gtin8", "gtin", "mpn"} and v:
                                return str(v)
                            found = find_gtin(v)
                            if found:
                                return found
                    elif isinstance(obj, list):
                        for item in obj:
                            found = find_gtin(item)
                            if found:
                                return found
                    return None
                val = find_gtin(data)
                if val and re.match(r"^\d{8,13}$", str(val)):
                    ean = str(val)
                    break
            except Exception:
                pass
            if ean:
                break

    return {
        "Nazwa":            nazwa or nazwa_lista,
        "EAN":              ean,
        "Cena_regularna":   cena_regularna,
        "Cena_promocyjna":  cena or cena_lista,
        "URL_produktu":     url,
    }

# ── Scraping listy produktów ───────────────────────────────────────────────────

async def scrape_lista(session, url_start, sklep):
    produkty = []
    url = url_start

    for nr in range(1, MAX_STRONY + 1):
        print(f"  -> strona {nr}: {url}")
        soup = await fetch(session, url)
        if soup is None:
            break

        karty = soup.select("div.ProdCena")
        print(f"     {len(karty)} produktow")

        for card in karty:
            a = card.select_one("h3 a")
            if not a:
                continue
            href = a.get("href", "")
            if href and not href.startswith("http"):
                base = url.rsplit("/", 1)[0]
                href = base + "/" + href.lstrip("/")
            cena_el = (card.select_one(".ProduktCena .CenaAktualna")
                       or card.select_one(".CenaAktualna"))
            cena = cena_el.get_text(strip=True) if cena_el else ""
            if a.get_text(strip=True) and href:
                produkty.append({
                    "Sklep":        sklep,
                    "Nazwa_lista":  a.get_text(strip=True),
                    "Cena_lista":   cena,
                    "URL_produktu": href,
                })

        next_tag = soup.find("link", rel="next") or soup.find("a", rel="next")
        if not next_tag:
            break
        next_href = next_tag.get("href", "")
        if not next_href:
            break
        if not next_href.startswith("http"):
            next_href = url.rsplit("/", 1)[0] + "/" + next_href.lstrip("/")
        url = next_href
        # await asyncio.sleep(OPOZNIENIE)  # usunięto dla maksymalnej prędkości

    return produkty

# ── Równoległe pobieranie szczegółów ──────────────────────────────────────────

async def scrape_szczegoly(session, produkty_lista, sklep):
    semafor = asyncio.Semaphore(CONCURRENCY)
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    t_start = time.time()

    async def fetch_one(prod):
        async with semafor:
            soup = await fetch(session, prod["URL_produktu"])
        if soup is None:
            dane = {
                "Nazwa": prod["Nazwa_lista"], "EAN": "",
                "Cena_regularna": "",
                "Cena_promocyjna": prod["Cena_lista"],
                "URL_produktu": prod["URL_produktu"],
            }
        else:
            dane = parsuj_produkt(soup, prod["URL_produktu"], prod["Nazwa_lista"], prod["Cena_lista"])
        dane["Sklep"] = sklep
        dane["Data"]  = now
        return dane

    tasks = [asyncio.create_task(fetch_one(p)) for p in produkty_lista]
    wyniki = []
    for i, coro in enumerate(asyncio.as_completed(tasks), 1):
        wyniki.append(await coro)
        if i % 50 == 0 or i == len(tasks):
            elapsed = time.time() - t_start
            rps = i / elapsed if elapsed > 0 else 0
            eta = (len(tasks) - i) / rps if rps > 0 else 0
            print(f"    [{i}/{len(tasks)}] {elapsed:.1f}s | {rps:.1f} prod/s | ETA: {eta:.0f}s")
    return wyniki

# ── Zapis ──────────────────────────────────────────────────────────────────────

def save_to_excel(data, filename):
    wb = Workbook()
    ws = wb.active
    ws.title = "Ceny Bot"
    ws.append(["Sklep", "Nazwa", "EAN", "Cena_regularna", "Cena_promocyjna", "URL_produktu", "Data"])
    for row in data:
        ws.append([row["Sklep"], row["Nazwa"], row["EAN"],
                   row.get("Cena_regularna", ""), row["Cena_promocyjna"],
                   row["URL_produktu"], row["Data"]])
    wb.save(filename)
    print(f"Zapisano {len(data)} produktow -> {filename}")

# ── Konfiguracja ───────────────────────────────────────────────────────────────

def wczytaj_konfiguracje(plik):
    try:
        df = pd.read_excel(plik, dtype=str)
        df.columns = df.columns.str.strip()
        aktywne = df[df["Aktywny"].str.upper() == "TAK"]
        return aktywne[["Sklep", "URL_wyprzedazy"]].to_dict("records")
    except Exception as e:
        print(f"Blad czytania {plik}: {e}")
        return []

# ── Main ───────────────────────────────────────────────────────────────────────

async def main_async():
    print("Bot zbierania cen — wersja async (bez Chrome)")
    sklepy = wczytaj_konfiguracje(PLIK_LINKI)
    if not sklepy:
        print("Brak aktywnych sklepow w links.xlsx")
        return

    connector = aiohttp.TCPConnector(limit=CONCURRENCY, ttl_dns_cache=300)
    async with aiohttp.ClientSession(headers=HEADERS, connector=connector) as session:
        all_produkty = []
        for sklep_data in sklepy:
            sklep = sklep_data["Sklep"]
            url   = sklep_data["URL_wyprzedazy"]
            print(f"\n{sklep}")

            t0 = time.time()
            lista = await scrape_lista(session, url, sklep)
            print(f"  lista: {len(lista)} produktow ({time.time()-t0:.1f}s)")

            t1 = time.time()
            wyniki = await scrape_szczegoly(session, lista, sklep)
            print(f"  szczegoly: {len(wyniki)} produktow ({time.time()-t1:.1f}s)")

            all_produkty.extend(wyniki)

    if all_produkty:
        save_to_excel(all_produkty, PLIK_WYNIKI)
        print("Gotowe!")
    else:
        print("Nie zebrano zadnych danych.")

def main():
    asyncio.run(main_async())

if __name__ == "__main__":
    main()