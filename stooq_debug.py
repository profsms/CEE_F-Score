"""
stooq_debug.py — Diagnose Stooq API connection
Run this first to see exactly what Stooq is returning before collect_returns.py
"""
import requests
from pathlib import Path

APIKEY_FILE = Path('./stooq_apikey.txt')

def load_apikey():
    key = APIKEY_FILE.read_text().strip()
    # Remove any BOM or hidden chars
    key = key.encode('ascii', errors='ignore').decode('ascii').strip()
    return key

apikey = load_apikey()
print(f"Key loaded ({len(apikey)} chars): '{apikey[:8]}...{apikey[-4:]}'")
print()

# Test with known working tickers in multiple formats
test_cases = [
    ('kgh',    'WSE KGHM without suffix'),
    ('kgh.wa', 'WSE KGHM with .wa suffix'),
    ('cez.pr', 'PSE CEZ with .pr suffix'),
    ('cez',    'PSE CEZ without suffix'),
    ('richter.bu', 'BSE Richter with .bu suffix'),
    ('richter',    'BSE Richter without suffix'),
]

for ticker, label in test_cases:
    url = (f"https://stooq.com/q/d/l/"
           f"?s={ticker}&d1=20200101&d2=20211231&i=d&apikey={apikey}")
    print(f"Testing: {label}")
    print(f"  URL: {url[:80]}...")
    try:
        resp = requests.get(url,
                            headers={'User-Agent': 'Mozilla/5.0'},
                            timeout=15)
        text = resp.text.strip()
        print(f"  Status: {resp.status_code}")
        print(f"  Response (first 200 chars):")
        print(f"  {repr(text[:200])}")
        print()
    except Exception as e:
        print(f"  ERROR: {e}")
        print()

# Extra BSE format tests
print("="*55)
print("Testing BSE / Budapest ticker formats")
print("="*55)
bse_attempts = [
    'richter.hu', 'richtergedeon', 'gedeon', 'richt',
    'otp', 'otp.hu', 'otp.bu',          # OTP Bank - very well known
    'mol', 'mol.hu', 'mol.bu',           # MOL Group
    '4ig.hu', 'pannergy.hu', 'dhouse.hu',
]
for ticker in bse_attempts:
    url = (f"https://stooq.com/q/d/l/"
           f"?s={ticker}&d1=20200101&d2=20211231&i=d&apikey={apikey}")
    try:
        resp = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
        text = resp.text.strip()
        status = '✓ DATA' if text.startswith('Date') else f'✗ {repr(text[:60])}'
        print(f"  {ticker:20s}: {status}")
    except Exception as e:
        print(f"  {ticker:20s}: ERROR {e}")

# Active BSE firms - finding correct tickers
print()
print("="*55)
print("Active BSE firms - finding correct tickers")
print("="*55)
bse_active = [
    # AUTOW = Autowallis Nyrt
    'autowallis.hu', 'awallish.hu', 'autow.hu', 'autow',
    # DHOUSE = Duna House Holding Nyrt
    'dh.hu', 'dunahaz.hu', 'dunah.hu', 'dunahouse.hu', 'dhouse', 'dhz.hu',
]
for ticker in bse_active:
    url = (f"https://stooq.com/q/d/l/"
           f"?s={ticker}&d1=20220101&d2=20231231&i=d&apikey={apikey}")
    try:
        resp = requests.get(url, headers={'User-Agent': 'Mozilla/5.0'}, timeout=10)
        text = resp.text.strip()
        if text.startswith('Date'):
            rows = len(text.split('\n')) - 1
            print(f"  {ticker:20s}: ✓ DATA ({rows} rows)")
        else:
            print(f"  {ticker:20s}: ✗ {repr(text[:40])}")
    except Exception as e:
        print(f"  {ticker:20s}: ERROR {e}")
