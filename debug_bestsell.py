"""Quick smoke-test: parse 1 product from each site."""
import asyncio, aiohttp
from bs4 import BeautifulSoup
import sys
sys.path.insert(0, ".")
from bestsellery_bot import parsuj_shoper_a, parsuj_bajkowelampy, parsuj_tomix

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept-Language": "pl-PL,pl;q=0.9",
}

TESTS = [
    ("PolskieLampy", "https://polskielampy.pl/kula-zewnetrzna-ball-mc30e-zuma-line-p-345173.html", "shoper_a"),
    ("Kinkiecik",    "https://kinkiecik.pl/marton-led-ip44-4000k-wh-az6855-azzardo-p-119354.html", "shoper_a"),
    ("BajkoweLampy", "https://bajkowelampy.pl/product-zul-1000161299-Lampa-Wiszaca-Ledowa-Masiero-Honice-S150-V99-Onyx.html", "bajkowelampy"),
    ("Tomix",        "https://www.tomix.pl/product-pol-39348-Lampa-wiszaca-SFERA-BLACK-MLP5739-Milagro.html", "tomix"),
]

async def main():
    async with aiohttp.ClientSession(headers=HEADERS) as session:
        for name, url, platform in TESTS:
            async with session.get(url, timeout=aiohttp.ClientTimeout(total=15)) as r:
                html = await r.text()
            soup = BeautifulSoup(html, "lxml")

            if platform == "shoper_a":
                symbol, cena_promo, cena_kat = parsuj_shoper_a(soup)
            elif platform == "bajkowelampy":
                symbol, cena_promo, cena_kat = parsuj_bajkowelampy(soup)
            else:
                symbol, cena_promo, cena_kat = parsuj_tomix(soup)

            print(f"\n[{name}]")
            print(f"  Symbol:    {repr(symbol)}")
            print(f"  Promo:     {repr(cena_promo)}")
            print(f"  Katalog:   {repr(cena_kat)}")

asyncio.run(main())
