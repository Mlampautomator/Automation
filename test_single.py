#!/usr/bin/env python3
"""Test pojedynczej strony mlamp.pl"""

import subprocess
import sys

# Auto-zainstaluj wymagane pakiety
def install_packages():
    packages = ['selenium', 'webdriver-manager', 'openpyxl']
    for package in packages:
        try:
            __import__(package)
        except ImportError:
            print(f"Instalowanie {package}...")
            subprocess.check_call([sys.executable, "-m", "pip", "install", package, "-q"])
    print("✓ Pakiety gotowe\n")

try:
    install_packages()
except Exception as e:
    print(f"Błąd instalacji: {e}")
    sys.exit(1)

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import time
import openpyxl

def get_first_url():
    """Pobierz pierwszy URL z urls.xlsx"""
    try:
        wb = openpyxl.load_workbook('urls.xlsx')
        ws = wb.active
        for row in ws.iter_rows(values_only=True):
            if row[0] and isinstance(row[0], str) and row[0].startswith('http'):
                return row[0]
    except Exception as e:
        print(f"Błąd czytania Excel: {e}")
    return None

def main():
    print("🔍 Test pojedynczej strony mlamp.pl\n")

    url = get_first_url()
    if not url:
        print("❌ Nie znaleziono URL w urls.xlsx")
        return

    print(f"📍 Test URL: {url}")

    # Setup driver
    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument('user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        print("🌐 Otwieranie strony...")
        driver.get(url)
        time.sleep(3)

        print("🍪 Akceptacja cookies...")
        # Akceptuj cookies
        cookie_selectors = [
            "//button[contains(text(), 'Akceptuję')]",
            "//button[contains(text(), 'Accept')]",
            "//a[contains(text(), 'Zaakceptuj')]",
        ]
        for selector in cookie_selectors:
            try:
                btn = driver.find_element(By.XPATH, selector)
                if btn.is_displayed():
                    btn.click()
                    print(f"✓ Zaakceptowano cookies przez: {selector}")
                    break
            except:
                pass

        time.sleep(2)

        print("🖼️  Szukanie obrazów...")
        all_imgs = driver.find_elements(By.TAG_NAME, "img")
        print(f"📊 Znaleziono {len(all_imgs)} obrazków na stronie\n")

        product_images = []
        for i, img in enumerate(all_imgs):
            src = img.get_attribute('src') or img.get_attribute('data-src') or img.get_attribute('data-original')
            alt = img.get_attribute('alt') or ''
            class_attr = img.get_attribute('class') or ''

            if src:
                print(f"🖼️  {i+1}: {src[:100]}...")
                print(f"   alt='{alt}' class='{class_attr}'")

                # Sprawdź czy to obraz produktu
                if any(keyword in src.lower() for keyword in ['product', 'upload', 'image', 'lamp']):
                    product_images.append(src)
                    print("   ✅ To może być obraz produktu!")
                print()

        print(f"🎯 Znaleziono {len(product_images)} potencjalnych obrazów produktów:")
        for i, img_url in enumerate(product_images, 1):
            print(f"   {i}. {img_url}")

    except Exception as e:
        print(f"❌ Błąd: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    main()