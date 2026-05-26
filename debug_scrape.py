import requests
from bs4 import BeautifulSoup

url = 'https://polskielampy.pl/wyprzedaz-lamp-c-766.html'
print('GET', url)
r = requests.get(url)
print('status', r.status_code)

soup = BeautifulSoup(r.text, 'html.parser')
anchors = soup.select('a[href*="/p-"]')
print('anchors', len(anchors))
for a in anchors[:40]:
    href = a.get('href')
    text = a.get_text(strip=True)
    parent = a.parent
    cls = parent.get('class') if parent is not None else None
    print(href, repr(text), parent.name if parent is not None else None, cls)
