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
CONCURRENCY = 8         # obniżono z 60 — serwer blokuje przy wysokim concurrency
TIMEOUT     = 15        # zwiększono z 5 — więcej czasu na odpowiedź serwera
OPOZNIENIE  = 0.5       # dodano 0.5s pauzy między requestami — zmniejsza blokady
MAX_RETRY   = 3         # przywrócono 3 próby dla stabilności

# ── Tryb diagnostyczny ────────────────────────────────────────────────────────
# Ustaw DEBUG_URL na URL produktu, aby zapisać jego HTML do debug_produkt.html
# i sprawdzić co strona faktycznie zwraca (pomocne gdy EAN/cena są puste).
# Przykład:
# DEBUG_URL = "https://polskielampy.pl/lampa-wpuszczana-mirrola-sq-p-333671.html"
DEBUG_URL = ""

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

# Ustaw na True, aby zapisać HTML pierwszego produktu do pliku debug_produkt.html
DEBUG_HTML = False
_debug_saved = False

def parsuj_produkt(soup, url, nazwa_lista, cena_lista):
    """
    polskielampy.pl umieszcza kluczowe dane w meta tagach Open Graph.
    """
    import json

    global _debug_saved
    if DEBUG_URL and url == DEBUG_URL and not _debug_saved:
        with open("debug_produkt.html", "w", encoding="utf-8") as f:
            f.write(str(soup))
        print(f"  [DEBUG] Zapisano HTML do debug_produkt.html ({url})")
        _debug_saved = True
    elif DEBUG_HTML and not _debug_saved:
        with open("debug_produkt.html", "w", encoding="utf-8") as f:
            f.write(str(soup))
        print("  [DEBUG] Zapisano HTML do debug_produkt.html")
        _debug_saved = True

    def meta(prop):
        tag = soup.find("meta", property=prop) or soup.find("meta", attrs={"property": prop})
        return (tag or {}).get("content", "").strip()

    def wyciagnij_cene(tekst):
        # Obsługuje formaty: 89,00 / 89.00 / 1 234,56 / 1234.56
        m = re.search(r"(\d[\d\s]*[,.]\d{2})", tekst)
        return m.group(1).replace("\xa0", "").replace(" ", "").strip() if m else ""

    def normalizuj_cene(val):
        # Zamień separator dziesiętny z kropki na przecinek (format Polski)
        if val and re.match(r"^\d+\.\d{2}$", val.strip()):
            return val.strip().replace(".", ",")
        return wyciagnij_cene(val) if val else ""

    # ── Nazwa ──────────────────────────────────────────────────────────────────
    h1 = soup.find("h1")
    nazwa = h1.get_text(strip=True) if h1 else meta("og:title")

    # ── Cena promocyjna ────────────────────────────────────────────────────────
    cena = normalizuj_cene(meta("product:price:amount"))
    if not cena:
        for sel in [
            ".CenaAktualna", ".CenaPromocyjna", ".price-promo",
            ".price-sale", ".price--promo", ".price",
        ]:
            el = soup.select_one(sel)
            if el:
                cena = wyciagnij_cene(el.get_text())
                if cena:
                    break

    # ── Cena regularna (przed zniżką) ──────────────────────────────────────────
    # polskielampy.pl: <p id="CenaPoprzednia">Cena katalogowa: <strong content="89.00">89,00 zł</strong></p>
    cena_regularna = normalizuj_cene(meta("product:original_price:amount"))

    if not cena_regularna:
        # Próbuj najpierw content-atrybut elementu strong (polskielampy.pl)
        el = soup.select_one("#CenaPoprzednia strong")
        if el:
            val = normalizuj_cene(el.get("content", "")) or wyciagnij_cene(el.get_text())
            if val:
                cena_regularna = val

    if not cena_regularna:
        for sel in [
            "#CenaPoprzednia",
            ".CenaStara", ".CenaPrzekr", ".cena-stara", ".cena-przekr",
            ".price-old", ".price-before", ".price-regular", ".PriceOld",
            "del .price", "s .price", ".old-price", ".regular-price",
            "[class*='OldPrice']", "[class*='old_price']", "[class*='StrikePrice']",
            ".ProduktStaraCena", ".StaraCena", ".price--regular",
            "span.CenaStara", "span.CenaPrzekr",
            "del", "s",
        ]:
            el = soup.select_one(sel)
            if el:
                val = wyciagnij_cene(el.get_text())
                if val:
                    cena_regularna = val
                    break

    if not cena_regularna:
        # Szukaj atrybutu data-price-old / data-old-price na dowolnym elemencie
        for attr in ["data-price-old", "data-old-price", "data-regular-price", "data-base-price"]:
            el = soup.find(attrs={attr: True})
            if el:
                val = wyciagnij_cene(el[attr])
                if val:
                    cena_regularna = val
                    break

    if not cena_regularna:
        # Szukaj tekstowo wzorca polskiego sklepu
        text = soup.get_text(" ", strip=True)
        match = re.search(
            r"(?:Cena katalogowa|Cena przed promocją|Cena przed|Cena regularna|Cena normalna)"
            r"[:\s]*(\d[\d\s\xa0]*[,.]\d{2})",
            text, re.I | re.UNICODE,
        )
        if match:
            cena_regularna = match.group(1).replace("\xa0", "").replace(" ", "")

    if not cena_regularna:
        # Fallback: cena omnibus z dyrektywy Omnibus (najniższa z 30 dni przed promocją)
        # polskielampy.pl: <p id="HistoriaCenProduktu"><span class="Informacja">Najniższa cena z 30 dni to X,XX zł</span></p>
        el = soup.select_one("#HistoriaCenProduktu")
        if el:
            val = wyciagnij_cene(el.get_text())
            if val:
                cena_regularna = val

    if not cena_regularna:
        # JSON-LD schema.org Product -> offers
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                offers = data.get("offers", {})
                if isinstance(offers, list):
                    offers = offers[0]
                # highPrice jest ceną regularną gdy produkt jest na promocji
                val = offers.get("highPrice", "") or offers.get("regularPrice", "")
                if val:
                    cena_regularna = str(val)
                    break
                # priceSpecification
                specs = offers.get("priceSpecification", [])
                if isinstance(specs, list):
                    for spec in specs:
                        t = spec.get("@type", "").lower()
                        pt = spec.get("priceType", "").lower()
                        if "regular" in t or "list" in t or "regular" in pt:
                            cena_regularna = str(spec.get("price", ""))
                            break
                if cena_regularna:
                    break
            except Exception:
                pass

    # ── EAN ───────────────────────────────────────────────────────────────────
    ean = ""

    # 1. Meta tagi produktowe
    for prop in ["product:ean", "og:ean"]:
        val = meta(prop)
        if val and re.match(r"^\d{8,13}$", val):
            ean = val
            break

    if not ean:
        # 2. itemprop gtin13 / gtin8 / gtin / ean
        for attr_val in ["gtin13", "gtin8", "gtin", "ean"]:
            el = soup.find(itemprop=attr_val)
            if el:
                raw = el.get("content", el.get_text(strip=True))
                m = re.search(r"(\d{8,13})", raw)
                if m:
                    ean = m.group(1)
                    break

    if not ean:
        # 3. data-ean / data-gtin / data-barcode
        for attr in ["data-ean", "data-gtin", "data-gtin13", "data-barcode", "data-product-ean"]:
            el = soup.find(attrs={attr: True})
            if el:
                m = re.search(r"(\d{8,13})", el[attr])
                if m:
                    ean = m.group(1)
                    break

    if not ean:
        # 4. Selektory CSS specyficzne dla polskielampy.pl i ogólne
        for sel in [
            "#KodEan strong[itemprop='gtin13']", "#KodEan strong", "#KodEan",
            "div.TbPoz strong[itemprop='gtin13']", "[itemprop='gtin13']",
            ".KodEan", ".ean", ".gtin", ".barcode",
            "td.TbPoz", "span.TbPoz",
        ]:
            el = soup.select_one(sel)
            if el:
                m = re.search(r"(\d{8,13})", el.get("content", el.get_text()))
                if m:
                    ean = m.group(1)
                    break

    if not ean:
        # 5. Szukaj tekstu "Kod EAN" i bierz liczbę z rodzeństwa lub rodzica
        for tag in soup.find_all(string=re.compile(r"Kod\s*EAN|EAN\s*:", re.I)):
            # sprawdź sam tekst
            m = re.search(r"(\d{8,13})", str(tag))
            if m:
                ean = m.group(1)
                break
            # sprawdź następne rodzeństwo
            parent = tag.parent
            if parent:
                nxt = parent.find_next_sibling()
                if nxt:
                    m = re.search(r"(\d{8,13})", nxt.get_text())
                    if m:
                        ean = m.group(1)
                        break
                # sprawdź cały wiersz tabeli (td/tr)
                row = parent.find_parent(["tr", "li", "div"])
                if row:
                    m = re.search(r"(\d{8,13})", row.get_text())
                    if m:
                        ean = m.group(1)
                        break

    if not ean:
        # 6. Tekstowe przeszukiwanie całej strony (ostatnia deska ratunku)
        text = soup.get_text(" ", strip=True)
        match = re.search(
            r"(?:Kod\s*EAN|EAN|GTIN-?13|GTIN-?8|GTIN|Barcode)[:\s#]*([0-9]{8,13})(?!\d)",
            text, re.I,
        )
        if match:
            ean = match.group(1)

    if not ean:
        # 7. JSON-LD gtin / gtin13 / gtin8
        for script in soup.find_all("script", type="application/ld+json"):
            try:
                data = json.loads(script.string or "")
                def find_gtin(obj):
                    if isinstance(obj, dict):
                        for k, v in obj.items():
                            if k.lower() in {"gtin13", "gtin8", "gtin", "ean"} and v:
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
        "Cena_promocyjna":  cena or wyciagnij_cene(cena_lista),
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
        await asyncio.sleep(OPOZNIENIE)  # pauza między stronami listy

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