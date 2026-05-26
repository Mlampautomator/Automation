#!/usr/bin/env python3
"""
Skrobacz Danych Firm - Prosta wersja dla użytkowników Notatnika
Po prostu wklej cały kod do edytora tekstu i uruchom go!
"""

# Auto-zainstaluj wymagane pakiety
import subprocess
import sys

def install_packages():
    packages = ['selenium', 'openpyxl', 'webdriver-manager']
    for package in packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Instalowanie {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-q"])
    print("✓ Wszystkie pakiety są gotowe!\n")

try:
    install_packages()
except Exception as e:
    print("Błąd instalacji pakietów:")
    import traceback
    traceback.print_exc()
    print("Upewnij się, że masz działające połączenie internetowe i zainstalowany Python.")
    input("Naciśnij Enter, aby zakończyć...")
    sys.exit(1)

# Teraz importuj pakiety
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
import time
import re
import os
from urllib.parse import urljoin, urlparse


def normalize_url(base: str, href: str) -> str:
    if not href:
        return ''
    return urljoin(base, href.strip())


def get_host(url: str) -> str:
    parsed = urlparse(url)
    host = parsed.netloc.lower()
    return host[4:] if host.startswith('www.') else host


def get_path_segments(url: str) -> list[str]:
    path = urlparse(url).path
    return [segment for segment in path.split('/') if segment]


def accept_cookies(driver):
    try:
        # Czekaj chwilę na załadowanie strony
        time.sleep(2)

        # Różne selektory dla przycisków akceptacji ciasteczek
        cookie_selectors = [
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'akceptuj')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'zgadzam')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'akceptuję')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'zgoda')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'ok')]",
            "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'akceptuj')]",
            "//a[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
            "//div[contains(@class, 'cookie') and contains(@class, 'accept')]",
            "//div[contains(@id, 'cookie') and contains(@class, 'accept')]",
            "//button[contains(@class, 'accept') or contains(@class, 'agree')]",
            "//button[contains(@id, 'accept') or contains(@id, 'agree')]"
        ]

        clicked = False
        for xpath in cookie_selectors:
            try:
                cookie_buttons = driver.find_elements(By.XPATH, xpath)
                for btn in cookie_buttons:
                    try:
                        if btn.is_displayed() and btn.is_enabled():
                            # Przewiń do przycisku jeśli trzeba
                            driver.execute_script("arguments[0].scrollIntoView(true);", btn)
                            time.sleep(0.5)
                            btn.click()
                            print(f"✓ Zaakceptowano ciasteczka (xpath: {xpath})")
                            clicked = True
                            break
                    except Exception as e:
                        print(f"Debug: Nie udało się kliknąć przycisku: {e}")
                        continue
                if clicked:
                    break
            except Exception:
                continue

        if clicked:
            # Czekaj na zniknięcie baneru ciasteczek
            time.sleep(3)
            print("✓ Baner ciasteczek powinien zniknąć")
        else:
            print("✓ Nie znaleziono przycisku akceptacji ciasteczek (lub już zaakceptowane)")

        # Dodatkowe czekanie na załadowanie strony po akceptacji
        time.sleep(2)

    except Exception as e:
        print(f"✓ Problem z akceptacją ciasteczek: {e}")


def scroll_to_bottom(driver, max_scrolls=40, delay=1.0):
    last_height = driver.execute_script("return document.body.scrollHeight")
    for _ in range(max_scrolls):
        driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        time.sleep(delay)
        new_height = driver.execute_script("return document.body.scrollHeight")
        if new_height == last_height:
            break
        last_height = new_height


def load_more_companies(driver, max_clicks=20):
    """Klikaj przyciski 'Załaduj więcej', 'Pokaż więcej' itp., aby załadować wszystkie firmy."""
    for _ in range(max_clicks):
        try:
            # Szukaj przycisków z tekstem zawierającym 'więcej', 'załaduj', 'pokaż' itp.
            load_buttons = driver.find_elements(By.XPATH, "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'więcej') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'załaduj') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'pokaż') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'load') or contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'show')]")
            clicked = False
            for btn in load_buttons:
                if btn.is_displayed() and btn.is_enabled():
                    btn.click()
                    time.sleep(2)  # Czekaj na załadowanie
                    clicked = True
                    break
            if not clicked:
                break  # Brak więcej przycisków
        except Exception:
            break
    # Na koniec przewiń jeszcze raz
    scroll_to_bottom(driver, max_scrolls=10, delay=0.5)



def normalize_url_for_deduplication(url: str) -> str:
    """Normalizuje URL do deduplikacji - usuwa fragmenty, parametry paginacji, trailing slashes"""
    try:
        parsed = urlparse(url)
        # Usuń fragment (#anchor)
        url = url.split('#')[0]
        # Usuń niektóre parametry paginacji jeśli nie są częścią ścieżki
        query_params = parsed.query
        if query_params:
            # Zachowaj tylko ważne parametry, usuń parametry paginacji
            params = []
            for param in query_params.split('&'):
                if not param.startswith(('page=', 'p=', 'offset=', 'start=')):
                    params.append(param)
            if params:
                url = url.split('?')[0] + '?' + '&'.join(params)
            else:
                url = url.split('?')[0]
        # Normalizuj trailing slash
        url = url.rstrip('/')
        return url.lower()
    except:
        return url.lower().rstrip('/')


def find_company_links(driver, base_url: str) -> list[str]:
    base_host = get_host(base_url)
    current_segments = get_path_segments(base_url)
    anchors = driver.find_elements(By.TAG_NAME, 'a')
    results = []
    seen = set()
    seen_paths = set()  # Dodatkowa deduplikacja na podstawie ścieżek dla otofirmy.pl

    print(f"DEBUG: Szukam linków na {base_host}, aktualne segmenty: {current_segments}")
    print(f"DEBUG: Znaleziono {len(anchors)} linków <a>")

    for link in anchors:
        href = normalize_url(base_url, link.get_attribute('href'))
        if not href or get_host(href) != base_host:
            continue

        # Normalizuj do deduplikacji
        normalized_href = normalize_url_for_deduplication(href)

        path_lower = urlparse(href).path.lower()
        href_segments = get_path_segments(href)

        print(f"DEBUG: Sprawdzam link: {href}, segmenty: {href_segments}")

        if base_host == 'aleo.com':
            if '/firma/' in path_lower and normalized_href not in seen:
                seen.add(normalized_href)
                results.append(href)
                print(f"DEBUG: Dodano (aleo.com): {href}")
        elif base_host == 'baza-firm.com.pl':
            if href.endswith('.html') and len(href_segments) > len(current_segments):
                if normalized_href not in seen and href != base_url:
                    seen.add(normalized_href)
                    results.append(href)
                    print(f"DEBUG: Dodano (.html): {href}")
            elif '/firma/' in path_lower or 'firma' in path_lower or 'biura-projekt' in path_lower:
                if normalized_href not in seen and href != base_url:
                    if len(href_segments) > len(current_segments) or '/firma/' in path_lower:
                        seen.add(normalized_href)
                        results.append(href)
                        print(f"DEBUG: Dodano (firma): {href}")
        elif base_host == 'otofirmy.pl':
            # Na otofirmy.pl bądź bardziej restrykcyjny - szukaj tylko konkretnych wzorców firm
            path_key = urlparse(href).path.lower().strip('/')

            # Sprawdź czy to wygląda na stronę konkretnej firmy (ma ID lub nazwę w URL)
            is_company_link = False
            if '/firma/' in path_lower and (len(href_segments) > 3 or any(seg.isdigit() for seg in href_segments)):
                is_company_link = True
            elif any(seg.isdigit() for seg in href_segments) and len(href_segments) >= 3:
                is_company_link = True
            elif len(path_key.split('/')) >= 2 and any(len(seg) > 10 for seg in href_segments):
                is_company_link = True

            if is_company_link and normalized_href not in seen and href != base_url and path_key not in seen_paths:
                seen.add(normalized_href)
                seen_paths.add(path_key)
                results.append(href)
                print(f"DEBUG: Dodano firmę (otofirmy.pl): {href}")
            elif is_company_link:
                print(f"DEBUG: Pominięto duplikat firmy (otofirmy.pl): {href} [ścieżka: {path_key}]")
        else:
            # Fallback: szukaj wszystkich linków, które wyglądają na strony firm
            if normalized_href not in seen and href != base_url:
                # Sprawdź czy link ma sensowny format (nie jest obrazkiem, css, js itp.)
                parsed = urlparse(href)
                path = parsed.path.lower()
                if not any(ext in path for ext in ['.jpg', '.png', '.gif', '.css', '.js', '.pdf', '.zip']):
                    # Sprawdź czy ścieżka wygląda na stronę firmy (ma przynajmniej 2 segmenty lub zawiera cyfry)
                    segments = get_path_segments(href)
                    if len(segments) >= 2 or any(seg.isdigit() for seg in segments) or len(path) > 10:
                        seen.add(normalized_href)
                        results.append(href)
                        print(f"DEBUG: Dodano (fallback): {href}")

    print(f"DEBUG: Łącznie znaleziono {len(results)} linków do firm")
    return sorted(results)


def find_baza_firm_category_pages(driver, base_url: str) -> list[str]:
    if get_host(base_url) != 'baza-firm.com.pl':
        return []

    # Generuj linki stron paginacji od 2 do 19
    provinces = []
    for page_num in range(2, 20):  # strona 2 do 19
        page_url = base_url.rstrip('/') + f'/strona-{page_num}/'
        provinces.append(page_url)

    return provinces


def find_otofirmy_category_pages(driver, base_url: str) -> list[str]:
    if get_host(base_url) != 'otofirmy.pl':
        return []

    provinces = []
    try:
        # Szukaj linków paginacji
        pagination_selectors = [
            'a[href*="page="]',
            'a[href*="/page/"]',
            'a[href*="?p="]',
            '.pagination a',
            '.paging a',
            'nav a'
        ]

        seen_pages = set()
        for selector in pagination_selectors:
            try:
                links = driver.find_elements(By.CSS_SELECTOR, selector)
                for link in links:
                    href = link.get_attribute('href')
                    if href and href != base_url:
                        # Sprawdź różne wzorce paginacji
                        page_match = re.search(r'[?&](?:page|p)=(\d+)', href) or re.search(r'/page/(\d+)', href)
                        if page_match:
                            page_num = int(page_match.group(1))
                            if page_num > 1 and page_num not in seen_pages and page_num <= 20:
                                seen_pages.add(page_num)
                                provinces.append(href)
            except:
                continue

    except Exception:
        pass

    # Jeśli nie znaleziono paginacji, spróbuj standardowych wzorców
    if not provinces:
        base_parts = base_url.rstrip('/').split('/')
        for page_num in range(2, 11):
            # Spróbuj ?page=N
            page_url = base_url + ('&' if '?' in base_url else '?') + f'page={page_num}'
            provinces.append(page_url)

    return sorted(list(set(provinces)))[:10]  # Ogranicz do 10 unikalnych stron


def extract_company_data(driver):
    page_text = driver.find_element(By.TAG_NAME, 'body').text
    firm_name = 'N/A'
    address = 'N/A'
    city = 'N/A'
    phone = 'N/A'
    email = 'N/A'
    contact_person = 'N/A'

    try:
        firm_name = driver.find_element(By.TAG_NAME, 'h1').text.strip()
    except Exception:
        pass

    if firm_name == 'N/A':
        firm_name = (driver.title or '').strip().replace('\n', ' ')

    address_match = re.search(r'(?i)(?:Adres|Adres\s*:|Lokalizacja|Lokalizacja\s*:|Siedziba|Siedziba\s*:|Adres firmy|Adres siedziby|Kontakt|Kontakt\s*:)\s*([^\n]{10,200})', page_text)
    if address_match:
        address = address_match.group(1).strip()
        print(f"DEBUG: Znaleziono adres z etykietą: {address}")
        # Usuń różne warianty tekstu o mapie
        address = re.sub(r'\s*(?:poka[żz]|zobacz|wyświetl|otwórz|mapa|google|maps)\s+(?:na\s+)?(?:mapie|map[ycie]|lokalizacji|map)', '', address, flags=re.IGNORECASE).strip()
        # Usuń inne niepotrzebne teksty
        address = re.sub(r'\s*(?:i dane|dane|kontakt|telefon|email|www|http|tel|fax|strona|mobile|kom|tel\.|fax\.|e-mail|www\.|http://|https://|mailto:).*', '', address, flags=re.IGNORECASE | re.DOTALL).strip()
        # Usuń wielokrotne spacje
        address = re.sub(r'\s+', ' ', address).strip()
        print(f"DEBUG: Adres po czyszczeniu: {address}")
    else:
        # Szukaj linii zawierającej kod pocztowy (rozszerzone wzorce)
        address_patterns = [
            r'([^\n]*\d{2}-\d{3}[^\n]*)',  # Kod pocztowy
            r'([^\n]*\b\d{1,5}\s+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż][^\n]*)',  # Ulica numer
            r'([^\n]*[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+\s+\d{1,5}[^\n]*)',  # Numer ulica
        ]
        for pattern in address_patterns:
            match = re.search(pattern, page_text, re.IGNORECASE)
            if match:
                candidate = match.group(1).strip()
                print(f"DEBUG: Kandydat adresu z wzorca: {candidate}")
                # Sprawdź czy nie zawiera zbyt wielu słów kluczowych
                keyword_count = sum(1 for word in ['telefon', 'email', 'www', 'http', 'tel.', 'fax', 'kontakt', 'i dane', 'dane'] if word in candidate.lower())
                if keyword_count <= 1:  # Pozwól na maksymalnie 1 słowo kluczowe
                    address = candidate
                    # Dodatkowe czyszczenie kandydata
                    address = re.sub(r'\s*(?:i dane|dane|kontakt|telefon|email|www|http|tel|fax|strona).*', '', address, flags=re.IGNORECASE).strip()
                    address = re.sub(r'\s+', ' ', address).strip()
                    print(f"DEBUG: Wybrano adres: {address}")
                    break

    if address != 'N/A':
        parts = address.split(',')
        if len(parts) > 1:
            city = parts[-1].strip()
        else:
            city_match = re.search(r'(\d{2}-\d{3})\s+(.+?)($|\n)', address)
            if city_match:
                city = city_match.group(2).strip()

    phone_match = re.search(r'(?i)(?:Telefon|Tel|tel|Telefon\s*:|Tel\s*:|tel\s*:)\s*([+]?[48]?\s*[\d\s\-()]{7,})', page_text)
    if phone_match:
        phone_candidate = phone_match.group(1).strip()
        # Wyczyść numer telefonu - usuń spacje, myślniki, nawiasy
        clean_phone = re.sub(r'[^\d+]', '', phone_candidate)
        # Sprawdź czy to wygląda na polski numer telefonu
        if (clean_phone.startswith('+48') and len(clean_phone) == 12) or \
           (clean_phone.startswith('48') and len(clean_phone) == 11) or \
           (clean_phone.startswith('0') and len(clean_phone) == 10) or \
           (len(clean_phone) == 9 and not clean_phone.startswith('+')):
            phone = phone_candidate
    else:
        # Szukaj ogólnych wzorców polskich numerów telefonów
        phone_patterns = [
            r'\+48\s*[\d\s\-()]{9}',  # +48 XXX XXX XXX
            r'48\s*[\d\s\-()]{9}',    # 48 XXX XXX XXX
            r'0[\d\s\-()]{9}',        # 0 XXX XXX XXX
            r'\d{3}[\s\-]\d{3}[\s\-]\d{3}',  # XXX-XXX-XXX lub XXX XXX XXX
            r'\(\d{3}\)\s*\d{3}[\s\-]\d{3}', # (XXX) XXX-XXX
        ]
        for pattern in phone_patterns:
            match = re.search(pattern, page_text)
            if match:
                phone_candidate = match.group(0)
                clean_phone = re.sub(r'[^\d+]', '', phone_candidate)
                if len(clean_phone) >= 9:
                    phone = phone_candidate
                    break

    email_match = re.search(r'[\w\.-]+@[\w\.-]+\.\w+', page_text)
    if email_match:
        email = email_match.group(0)

    contact_match = re.search(r'(?i)(?:Kontakt|Osoba|Przedstawiciel)[:\s]+([A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+\s+[A-ZĄĆĘŁŃÓŚŹŻ][a-ząćęłńóśźż]+)', page_text)
    if contact_match:
        contact_person = contact_match.group(1).strip()
        # Sprawdź czy znaleziony tekst nie zawiera słów wskazujących na fałszywy match
        invalid_words = ['pomoc', 'regulamin', 'informacje', 'dane', 'kontakt', 'telefon', 'email', 'adres', 'firma', 'spółka', 'zakład']
        if any(word in contact_person.lower() for word in invalid_words):
            contact_person = 'N/A'

    return {
        'Nazwa Firmy': firm_name,
        'Adres': address,
        'Miasto': city,
        'Imię i Nazwisko': contact_person,
        'Telefon': phone,
        'Email': email,
    }


# Pobierz URL od użytkownika
print("=" * 60)
print("SKROBACZ DANYCH FIRM")
print("=" * 60)
print("Uwaga: uruchom ten plik project01_PL.py, a nie output.xlsx.")
print("Gdy zobaczysz pytanie, wklej adres strony listy firm i naciśnij Enter.")
url = input("\nWklej URL listy firm: ").strip()

if not url.startswith('http'):
    url = 'https://' + url

print(f"\nURL: {url}")
print("Uruchamianie przeglądarki...")

# Skonfiguruj Selenium WebDriver
chrome_options = Options()
driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=chrome_options)

# Załaduj stronę
driver.get(url)

accept_cookies(driver)

print("Czekam na załadowanie dynamicznej zawartości...")
time.sleep(5)

# Dodatkowe czekanie na zniknięcie ewentualnych overlay
try:
    WebDriverWait(driver, 10).until(
        lambda d: d.execute_script("return document.readyState") == "complete"
    )
    print("✓ Strona w pełni załadowana")
except:
    print("✓ Kontynuuję mimo problemów z ładowaniem")

scroll_to_bottom(driver)
if get_host(url) == 'baza-firm.com.pl':
    load_more_companies(driver)
elif get_host(url) == 'otofirmy.pl':
    load_more_companies(driver)

company_urls = find_company_links(driver, url)
print(f"✓ Z pierwszej strony: {len(company_urls)} linków")

if get_host(url) == 'baza-firm.com.pl':
    provinces = find_baza_firm_category_pages(driver, url)
    if provinces:
        print(f"✓ Znaleziono {len(provinces)} podstron kategorii na baza-firm.com.pl")
        for i, province_url in enumerate(provinces, 1):
            print(f"Przetwarzam podstronę {i}/{len(provinces)}: {province_url}")
            driver.get(province_url)
            time.sleep(3)
            scroll_to_bottom(driver)
            load_more_companies(driver)
            page_urls = find_company_links(driver, province_url)
            print(f"  Z tej strony: {len(page_urls)} linków")
            company_urls.extend(page_urls)
elif get_host(url) == 'otofirmy.pl':
    provinces = find_otofirmy_category_pages(driver, url)
    if provinces:
        print(f"✓ Znaleziono {len(provinces)} podstron kategorii na otofirmy.pl")
        for i, province_url in enumerate(provinces, 1):
            print(f"Przetwarzam podstronę {i}/{len(provinces)}: {province_url}")
            driver.get(province_url)
            time.sleep(3)
            scroll_to_bottom(driver)
            load_more_companies(driver)
            page_urls = find_company_links(driver, province_url)
            print(f"  Z tej strony: {len(page_urls)} linków")
            company_urls.extend(page_urls)
    else:
        print("✓ Nie znaleziono podstron paginacji na otofirmy.pl")

print(f"✓ Łącznie zebrano {len(company_urls)} linków przed deduplikacją")

company_urls = list(dict.fromkeys(company_urls))

# Dodatkowa deduplikacja z normalizacją URL-i
normalized_urls = {}
for url in company_urls:
    normalized = normalize_url_for_deduplication(url)
    if normalized not in normalized_urls:
        normalized_urls[normalized] = url

company_urls = list(normalized_urls.values())
print(f"✓ Po deduplikacji: {len(company_urls)} unikalnych linków")

print(f"✓ Znaleziono {len(company_urls)} unikalnych linków do firm")
if len(company_urls) == 0:
    if '/firma/' in url or 'firma' in url.lower():
        print("⚠️ Nie znaleziono linków do firm na stronie. Spróbuję przetworzyć podany URL jako stronę pojedynczej firmy.")
        company_urls = [url]
    else:
        print("⚠️ Nie znaleziono linków do firm. Upewnij się, że podałeś stronę listy firm lub bezpośredni link do pojedynczej firmy.")
        driver.quit()
        input("Naciśnij Enter, aby zakończyć...")
        sys.exit(1)

# Zapytaj ile pobrać (lub przetwórz wszystkie)
try:
    num_input = input(f"\nPobrać wszystkie {len(company_urls)} firm? (tak/nie) [domyślnie: tak]: ").strip().lower()
    if num_input in ['n', 'nie']:
        num_to_scrape = int(input(f"Ile firm pobrać? (1-{len(company_urls)}): ").strip())
        if num_to_scrape < 1 or num_to_scrape > len(company_urls):
            num_to_scrape = len(company_urls)
    else:
        num_to_scrape = len(company_urls)
except:
    num_to_scrape = len(company_urls)

print(f"Pobieranie {num_to_scrape} firm...\n")

rows = []
for idx, company_url in enumerate(company_urls[:num_to_scrape]):
    try:
        print(f"[{idx + 1}/{num_to_scrape}] Przetwarzanie...", end=' ', flush=True)
        driver.get(company_url)
        time.sleep(2)
        scroll_to_bottom(driver, max_scrolls=20, delay=0.5)

        data = extract_company_data(driver)
        print(f"✓ {data['Nazwa Firmy'][:40]}")
        rows.append(data)
    except Exception:
        print("✗ Błąd")

# Zamknij przeglądarkę
driver.quit()

print(f"\n{'=' * 60}")
print(f"✓ UKOŃCZONE: Wyodrębniono {len(rows)} firm")
print(f"{'=' * 60}\n")

# Zapisz do Excela
wb = Workbook()
ws = wb.active
ws.title = "Firmy"

# Dodaj nagłówki
ws.append(["Nazwa Firmy", "Adres", "Miasto", "Imię i Nazwisko", "Telefon", "Email"])

# Dodaj dane
for row in rows:
    ws.append([row["Nazwa Firmy"], row["Adres"], row["Miasto"], row["Imię i Nazwisko"], row["Telefon"], row["Email"]])

# Automatyczne dopasowanie szerokości kolumn
for column in ws.columns:
    max_length = 0
    column_letter = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    ws.column_dimensions[column_letter].width = min(max_length + 2, 50)

# Pobierz bieżący katalog
output_path = os.path.join(os.getcwd(), "output.xlsx")
wb.save(output_path)
print(f"✓ Dane zapisane w: {output_path}\n")

# Otwórz plik wynikowy automatycznie
try:
    os.startfile(output_path)
    print("✓ Otwieranie pliku output.xlsx...")
except Exception as e:
    print(f"Nie udało się automatycznie otworzyć pliku: {e}")

input("Naciśnij Enter, aby wyjść...")
