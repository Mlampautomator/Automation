#!/usr/bin/env python3
"""Screenshot strony mlamp.pl"""

from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from selenium.webdriver.common.by import By
import time

def screenshot_page():
    url = "https://mlamp.pl/pl/products/zewnetrzna-lampa-scienna-palin-pir-75442-czujnik-ruchu-tuba-ip44-biala-106665"

    print(f"Robię screenshot: {url}")

    options = Options()
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')

    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)

    try:
        driver.get(url)
        time.sleep(5)

        # Akceptuj cookies
        try:
            cookie_btn = driver.find_element(By.XPATH, "//button[contains(text(), 'Akceptuję')]")
            cookie_btn.click()
            print("Zaakceptowano cookies")
            time.sleep(2)
        except:
            print("Brak przycisku cookies")

        # Zrób screenshot
        driver.save_screenshot('screenshot.png')
        print("Screenshot zapisany jako screenshot.png")

        # Znajdź obrazki
        images = driver.find_elements(By.TAG_NAME, "img")
        print(f"Znaleziono {len(images)} obrazków")

        for i, img in enumerate(images[:5]):
            src = img.get_attribute('src')
            print(f"{i+1}: {src}")

    except Exception as e:
        print(f"Błąd: {e}")
    finally:
        driver.quit()

if __name__ == "__main__":
    screenshot_page()