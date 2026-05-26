#!/usr/bin/env python3
"""
Program do spisywania cen promocji/wyprzedaży konkurenta wraz z kodem EAN i nazwą produktu.
Można wklejać linki do terminala lub czytać z pliku Excel.

Wymagania: pip install requests beautifulsoup4 openpyxl selenium webdriver-manager
"""

import re
import time
import requests
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from openpyxl import Workbook
from datetime import datetime
import sys

# Konfiguracja
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pl-PL,pl;q=0.9",
}

OPOZNIENIE = 2  # sekundy między requestami

def setup_driver():
    """Konfiguracja Selenium WebDriver."""
    options = Options()
    options.add_argument("--headless")  # Bez interfejsu graficznego
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    service = Service(ChromeDriverManager().install())
    driver = webdriver.Chrome(service=service, options=options)
    return driver

def accept_cookies(driver):
    """Akceptacja ciasteczek na stronie."""
    try:
        time.sleep(2)
        cookie_selectors = [
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'akceptuj')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'zgadzam')]",
            "//button[contains(translate(., 'ABCDEFGHIJKLMNOPQRSTUVWXYZ', 'abcdefghijklmnopqrstuvwxyz'), 'accept')]",
            "//button[contains(@class, 'accept') or contains(@class, 'agree')]",
        ]
        for xpath in cookie_selectors:
            try:
                btn = WebDriverWait(driver, 5).until(EC.element_to_be_clickable((By.XPATH, xpath)))
                btn.click()
                print("✓ Zaakceptowano ciasteczka")
                break
            except:
                continue
    except:
        pass

def scrape_product(url):
    """Scrapuje pojedynczy produkt: nazwa, EAN, cena promocyjna."""
    driver = setup_driver()
    try:
        driver.get(url)
        accept_cookies(driver)
        time.sleep(3)  # Czekaj na załadowanie

        # Pobierz HTML
        html = driver.page_source
        soup = BeautifulSoup(html, 'html.parser')

        # Wyciągnij nazwę produktu
        nazwa_selectors = [
            "h1", ".product-name", ".product-title", "[class*='name']",
            ".title", ".nazwa-produktu"
        ]
        nazwa = ""
        for sel in nazwa_selectors:
            el = soup.select_one(sel)
            if el:
                nazwa = el.get_text(strip=True)
                break

        # Wyciągnij EAN
        ean_selectors = [
            "[data-ean]", ".ean", ".kod-ean", "[class*='ean']",
            "span:contains('EAN')", "div:contains('EAN')"
        ]
        ean = ""
        for sel in ean_selectors:
            if ":" in sel:
                # Tekst zawierający
                elements = soup.find_all(text=re.compile(r"EAN[:\s]*(\d+)", re.IGNORECASE))
                for el in elements:
                    match = re.search(r"EAN[:\s]*(\d+)", el, re.IGNORECASE)
                    if match:
                        ean = match.group(1)
                        break
            else:
                el = soup.select_one(sel)
                if el:
                    text = el.get_text(strip=True)
                    match = re.search(r"(\d{8,13})", text)
                    if match:
                        ean = match.group(1)
                        break
            if ean:
                break

        # Wyciągnij cenę promocyjną
        cena_selectors = [
            ".price-promo", ".cena-promocyjna", ".promo-price",
            ".price-sale", ".sale-price", ".discount-price",
            ".price", ".cena"
        ]
        cena_promocyjna = ""
        for sel in cena_selectors:
            el = soup.select_one(sel)
            if el:
                text = el.get_text(strip=True)
                # Wyciągnij liczbę z tekstu ceny
                match = re.search(r"(\d+[,.]\d{2})", text.replace(" ", "").replace(",", "."))
                if match:
                    cena_promocyjna = match.group(1)
                    break

        return {
            "Nazwa": nazwa,
            "EAN": ean,
            "Cena_promocyjna": cena_promocyjna,
            "Data": datetime.now().strftime("%Y-%m-%d %H:%M")
        }

    except Exception as e:
        print(f"Błąd scrapowania {url}: {e}")
        return None
    finally:
        driver.quit()

def get_links_from_terminal():
    """Pobiera linki z terminala."""
    print("Wklej linki do produktów (jeden na linię). Zakończ pustą linią:")
    links = []
    while True:
        line = input().strip()
        if not line:
            break
        if line.startswith("http"):
            links.append(line)
    return links

def get_links_from_excel(filepath):
    """Pobiera linki z pliku Excel."""
    import pandas as pd
    try:
        df = pd.read_excel(filepath, dtype=str)
        df.columns = df.columns.str.strip()
        # Szukaj kolumny z linkami
        link_columns = [col for col in df.columns if 'link' in col.lower() or 'url' in col.lower()]
        if not link_columns:
            link_columns = df.columns[0]  # Pierwsza kolumna
        links = df[link_columns[0]].dropna().tolist()
        return [link for link in links if link.startswith("http")]
    except Exception as e:
        print(f"Błąd czytania Excela: {e}")
        return []

def save_to_excel(data, filename="output.xlsx"):
    """Zapisuje dane do Excela."""
    wb = Workbook()
    ws = wb.active
    ws.title = "Ceny Promocyjne"
    headers = ["Link", "Nazwa", "EAN", "Cena_promocyjna", "Data"]
    ws.append(headers)
    for row in data:
        ws.append([row["Link"], row["Nazwa"], row["EAN"], row["Cena_promocyjna"], row["Data"]])
    wb.save(filename)
    print(f"✓ Zapisano wyniki do {filename}")

def main():
    print("Program do spisywania cen promocji konkurenta")
    print("Wybierz źródło linków:")
    print("1. Terminal (wklej linki)")
    print("2. Plik Excel")
    choice = input("Wybór (1/2): ").strip()

    if choice == "1":
        links = get_links_from_terminal()
    elif choice == "2":
        filepath = input("Podaj ścieżkę do pliku Excel: ").strip()
        links = get_links_from_excel(filepath)
    else:
        print("Nieprawidłowy wybór.")
        return

    if not links:
        print("Brak linków do przetworzenia.")
        return

    print(f"Znaleziono {len(links)} linków. Rozpoczynam scrapowanie...")

    results = []
    for i, link in enumerate(links, 1):
        print(f"[{i}/{len(links)}] Scrapowanie: {link}")
        data = scrape_product(link)
        if data:
            data["Link"] = link
            results.append(data)
        time.sleep(OPOZNIENIE)

    if results:
        save_to_excel(results)
        print("✓ Zakończono!")
    else:
        print("✗ Nie udało się pobrać żadnych danych.")

if __name__ == "__main__":
    main()