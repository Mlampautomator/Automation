#!/usr/bin/env python3
"""Sprawdź strukturę strony mlamp.pl"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

def check_page():
    url = "https://mlamp.pl/pl/products/zewnetrzna-lampa-scienna-palin-pir-75442-czujnik-ruchu-tuba-ip44-biala-106665"

    print(f"Sprawdzam: {url}")

    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        response = requests.get(url, headers=headers, timeout=30)
        print(f"Status: {response.status_code}")
        print(f"Długość HTML: {len(response.content)} znaków")

        soup = BeautifulSoup(response.content, 'html.parser')

        # Znajdź tytuł
        title = soup.find('title')
        print(f"Tytuł: {title.text if title else 'Brak'}")

        # Znajdź wszystkie obrazki
        images = soup.find_all('img')
        print(f"\nWszystkie obrazki ({len(images)}):")

        product_images = []
        for i, img in enumerate(images):
            src = img.get('src') or img.get('data-src') or img.get('data-original')
            alt = img.get('alt', '')
            class_attr = ' '.join(img.get('class', []))

            print(f"{i+1}. src: {src}")
            print(f"   alt: '{alt}'")
            print(f"   class: '{class_attr}'")

            if src and 'http' in src:
                full_url = urljoin(url, src)
                product_images.append(full_url)
                print(f"   ✅ DODANO: {full_url}")
            print()

        print(f"\nPodsumowanie: {len(product_images)} obrazów z http")

        # Zapisz do pliku
        with open('debug_images.txt', 'w', encoding='utf-8') as f:
            f.write(f"URL: {url}\n\n")
            for img_url in product_images:
                f.write(f"{img_url}\n")

        print("Zapisano do debug_images.txt")

    except Exception as e:
        print(f"Błąd: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    check_page()