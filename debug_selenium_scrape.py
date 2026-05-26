from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager
from bs4 import BeautifulSoup

options = Options()
options.add_argument('--headless')
options.add_argument('--no-sandbox')
options.add_argument('--disable-dev-shm-usage')
service = Service(ChromeDriverManager().install())
driver = webdriver.Chrome(service=service, options=options)
try:
    url = 'https://polskielampy.pl/wyprzedaz-lamp-c-766.html'
    print('GET', url)
    driver.get(url)
    driver.implicitly_wait(5)
    html = driver.page_source
    print('HTML len', len(html))
    with open('debug_selenium.html', 'w', encoding='utf-8') as f:
        f.write(html)
    soup = BeautifulSoup(html, 'html.parser')
    print('Links /p-', len(soup.select('a[href*="/p-"]')))
    print('Links p-', len(soup.select('a[href*="p-"]')))
    print('Links a[href*="p-"]', len(soup.select('a[href*="p-"]')))
    print('Links a[class*="product"]', len(soup.select('a[class*="product"]')))
    print('Links a[href*="lamp"]', len(soup.select('a[href*="lamp"]')))
    
    anchors = soup.find_all('a')
    for i,a in enumerate(anchors[:60]):
        href=a.get('href')
        text=a.get_text(strip=True)
        if href and '/p-' in href:
            print('MATCH', href, repr(text), a.parent.name, a.parent.get('class'))
    
    cards = soup.select('[class*="product"]')
    print('product cards', len(cards))
    for c in cards[:10]:
        print('CARD', c.name, c.get('class'), repr(c.get_text(strip=True)[:120]))
finally:
    driver.quit()
