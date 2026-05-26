#!/usr/bin/env python3
"""Prosty test pobierania obrazów z mlamp.pl"""

import requests
from bs4 import BeautifulSoup
from urllib.parse import urljoin

def test_single_url():
    # Przykładowy URL
    url = "https://mlamp.pl/pl/products/zewnetrzna-lampa-scienna-palin-pir-75442-czujnik-ruchu-tuba-ip44-biala-106665"

    print(f"Test URL: {url}")

    try:
        headers = {
            'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/91.0.4472.124 Safari/537.36'
        }

        response = requests.get(url, headers=headers, timeout=30)
        response.raise_for_status()

        soup = BeautifulSoup(response.content, 'html.parser')

        print(f"Status: {response.status_code}")
        print(f"Długość HTML: {len(response.content)} znaków")

        # Znajdź wszystkie obrazki
        images = soup.find_all('img')
        print(f"\nZnaleziono {len(images)} obrazków:")

        product_images = []
        for i, img in enumerate(images):
            src = img.get('src') or img.get('data-src') or img.get('data-original')
            alt = img.get('alt', '')
            class_attr = img.get('class', [])

            if src:
                print(f"{i+1}. {src}")
                print(f"   alt: {alt}")
                print(f"   class: {class_attr}")

                # Sprawdź czy to obraz produktu
                if any(keyword in src.lower() for keyword in ['product', 'upload', 'lamp', 'image', '75442']):
                    full_url = urljoin(url, src)
                    product_images.append(full_url)
                    print("   ✅ OBRAZ PRODUKTU!")
                print()

        print(f"\n🎯 Obrazy produktów ({len(product_images)}):")
        for img_url in product_images:
            print(f"   {img_url}")

        # Wyciągnij ID produktu
        import re
        match = re.search(r'-(\d{5,6})(?:-|$)', url)
        if match:
            product_id = match.group(1)
            print(f"\n🆔 ID produktu: {product_id}")

            # Pobierz pierwszy obraz jako test
            if product_images:
                test_url = product_images[0]
                print(f"\n📥 Test pobierania: {test_url}")

                img_response = requests.get(test_url, headers=headers, timeout=30)
                if img_response.status_code == 200:
                    filename = f"{product_id}-1.jpg"
                    with open(filename, 'wb') as f:
                        f.write(img_response.content)
                    print(f"✓ Pobrano jako: {filename}")
                else:
                    print(f"❌ Błąd pobierania: {img_response.status_code}")

    except Exception as e:
        print(f"❌ Błąd: {e}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    test_single_url()