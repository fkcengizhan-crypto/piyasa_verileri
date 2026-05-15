from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time
from datetime import datetime
import json
import os

# ----------------------------- HİSSE VERİLERİ -----------------------------
def turkce_sayi_cevir(deger):
    if isinstance(deger, str):
        deger = deger.replace('.', '')
        deger = deger.replace(',', '.')
    try:
        return float(deger)
    except ValueError:
        return None

def hisse_verilerini_cek():
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    # Chrome/Chromium binary konumu (GitHub Actions için)
    chrome_binary = os.environ.get("CHROME_BINARY_PATH")
    if chrome_binary:
        chrome_options.binary_location = chrome_binary

    # ChromeDriver konumu (ortam değişkeni varsa kullan, yoksa Selenium otomatik yönetir)
    driver_path = os.environ.get("CHROME_DRIVER_PATH")
    if driver_path:
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
    else:
        driver = webdriver.Chrome(options=chrome_options)

    driver.get("https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/default.aspx")
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    driver.quit()

    table = soup.find('table', id='DataTables_Table_0')
    if not table:
        raise Exception("Hisse tablosu bulunamadı!")

    rows = table.select('tbody tr')
    data = []
    for row in rows:
        cols = row.find_all('td')
        if len(cols) >= 2:
            hisse = cols[0].find('a').text.strip()
            fiyat = cols[1].get_text(strip=True)
            data.append([hisse, fiyat])

    df = pd.DataFrame(data, columns=["Hisse", "Son Fiyat (TL)"])
    df["Son Fiyat (TL)"] = df["Son Fiyat (TL)"].apply(turkce_sayi_cevir)
    return df

# ----------------------------- FON VERİLERİ -----------------------------
def fon_verilerini_cek():
    url = "https://www.tefas.gov.tr/api/funds/fonGnlBlgSiraliGetirDosya"
    headers = {
        "Accept": "*/*",
        "Authorization": "Bearer ST-tefaswebwse3irfmSBj4iRAzGPbAlS94Se",
        "Content-Type": "application/json",
        "Origin": "https://www.tefas.gov.tr",
        "Referer": "https://www.tefas.gov.tr/tr/fon-verileri",
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/147.0.0.0 Safari/537.36",
    }

    bugun = datetime.now().strftime("%Y%m%d")
    payload = {
        "dil": "TR",
        "fonTipi": "YAT",
        "fonKod": None,
        "fonGrup": None,
        "basTarih": bugun,
        "bitTarih": bugun,
        "fonTurKod": None,
        "fonUnvanTip": None,
        "kurucuKod": None,
        "fonTurAciklama": None,
        "sfonTurKod": None
    }

    response = requests.post(url, headers=headers, json=payload, timeout=30)
    response.raise_for_status()
    data = response.json()

    if "resultList" not in data or not data["resultList"]:
        raise Exception("Fon verisi alınamadı.")

    fonlar = data["resultList"]
    kayitlar = [[fon["fonKodu"], fon["fiyat"]] for fon in fonlar if fon.get("fiyat") is not None]
    df = pd.DataFrame(kayitlar, columns=["Fon Kodu", "Fiyat"])
    df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
    return df

# ----------------------------- ANA İŞLEM -----------------------------
if __name__ == "__main__":
    print("Hisse verileri çekiliyor...")
    df_hisse = hisse_verilerini_cek()
    print(f"{len(df_hisse)} hisse bulundu.")

    print("Fon verileri çekiliyor...")
    df_fon = fon_verilerini_cek()
    print(f"{len(df_fon)} fon bulundu.")

    # Excel: Aynı sayfada A1'den fon tablosu, boş sütun, ardından hisse tablosu
    with pd.ExcelWriter("piyasa_verileri.xlsx", engine="openpyxl") as writer:
        df_fon.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=0)
        df_hisse.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=len(df_fon.columns) + 1)
    print("Veriler 'piyasa_verileri.xlsx' dosyasına kaydedildi (tek sayfa: A1'den fon, boş sütun, ardından hisse).")

    # JSON çıktısı
    json_data = {
        "hisseler": df_hisse.to_dict(orient="records"),
        "fonlar": df_fon.to_dict(orient="records")
    }
    with open("piyasa_verileri.json", "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)
    print("Veriler 'piyasa_verileri.json' dosyasına kaydedildi (iki ayrı liste).")