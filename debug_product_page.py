import requests
from bs4 import BeautifulSoup

url = 'https://polskielampy.pl/lampa-wpuszczana-mirrola-sq-kierunkowa-r10278-redlux-outlet-p-333671.html'
print('URL:', url)
r = requests.get(url)
print('status', r.status_code)
soup = BeautifulSoup(r.text, 'html.parser')
print('title:', soup.title.string if soup.title else None)
print('h1 count:', len(soup.select('h1')))
print('EAN strings:')
for t in soup.find_all(string=lambda s: s and 'EAN' in s):
    print(repr(t))
print('Price candidates:')
for sel in ['.CenaAktualna', '.CenaPromocyjna', '.product-price', '.price', '.promo-price', '.price-sale', '.sale-price']:
    for el in soup.select(sel)[:10]:
        print(sel, repr(el.get_text(strip=True)))
print('Search for digits:')
for text in [t.get_text(strip=True) for t in soup.find_all(['span','div','li','td'])][:50]:
    if 'zł' in text or 'EAN' in text:
        print(repr(text))
