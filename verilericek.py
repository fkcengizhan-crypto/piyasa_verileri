from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time
from datetime import datetime, timedelta
import json
import os

# ----------------------------- SELENIUM DRIVER -----------------------------
def selenium_driver_olustur():
    """Hem Windows hem GitHub Actions için ortak Selenium driver."""
    chrome_options = Options()
    chrome_options.add_argument("--headless")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")

    chrome_binary = os.environ.get("CHROME_BINARY_PATH")
    if chrome_binary:
        chrome_options.binary_location = chrome_binary

    driver_path = os.environ.get("CHROME_DRIVER_PATH")
    if driver_path:
        service = Service(driver_path)
        return webdriver.Chrome(service=service, options=chrome_options)
    else:
        return webdriver.Chrome(options=chrome_options)


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
    driver = selenium_driver_olustur()
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

    # Bugünden başla, 15 gün geriye git
    for i in range(15):
        tarih = (datetime.now() - timedelta(days=i)).strftime("%Y%m%d")
        print(f"Tarih deneniyor: {tarih}")

        payload = {
            "dil": "TR",
            "fonTipi": "YAT",
            "fonKod": None,
            "fonGrup": None,
            "basTarih": tarih,
            "bitTarih": tarih,
            "fonTurKod": None,
            "fonUnvanTip": None,
            "kurucuKod": None,
            "fonTurAciklama": None,
            "sfonTurKod": None
        }

        try:
            response = requests.post(url, headers=headers, json=payload, timeout=30)
            print(f"HTTP Durum Kodu: {response.status_code}")

            # HTTP hatası varsa log'la ve sonraki güne geç
            if response.status_code != 200:
                print(f"HTTP hatası, yanıt: {response.text[:300]}")
                continue

            data = response.json()

            # API iç hata mesajı varsa log'la
            if isinstance(data, dict) and ("error" in data or "message" in data):
                print(f"API mesajı: {data}")
                if "resultList" not in data or not data["resultList"]:
                    continue

            # resultList kontrolü
            if "resultList" not in data or not data["resultList"]:
                print(f"{tarih} için fon verisi bulunamadı.")
                continue

            # Veri bulundu, işle
            fonlar = data["resultList"]
            kayitlar = [[fon["fonKodu"], fon["fiyat"]] for fon in fonlar if fon.get("fiyat") is not None]
            df = pd.DataFrame(kayitlar, columns=["Fon Kodu", "Fiyat"])
            df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
            print(f"{tarih} tarihine ait {len(df)} fon verisi alındı.")
            return df

        except requests.exceptions.RequestException as e:
            print(f"İstek hatası ({tarih}): {e}")
            continue
        except json.JSONDecodeError as e:
            print(f"JSON ayrıştırma hatası ({tarih}): {e}")
            print(f"Ham yanıt: {response.text[:300]}")
            continue
        except Exception as e:
            print(f"Beklenmeyen hata ({tarih}): {e}")
            continue

    # 15 günde hiç veri bulunamadıysa boş DataFrame döndür
    print("UYARI: Son 15 günde fon verisi bulunamadı. Lütfen token'ı ve bağlantıyı kontrol edin.")
    return pd.DataFrame(columns=["Fon Kodu", "Fiyat"])


def bloomberg_verilerini_cek():
    driver = selenium_driver_olustur()
    driver.get("https://www.bloomberght.com/piyasalar")
    time.sleep(5)
    soup = BeautifulSoup(driver.page_source, "html.parser")
    driver.quit()

    slides = soup.select("div.swiper-slide[data-swiper-slide-index]")
    gorulmus = set()
    data = []
    for slide in slides:
        sembol  = slide.select_one("span.text-xs.text-ellipsis")
        fiyat   = slide.select_one("span.lastPrice")
        degisim = slide.select_one("span.percentChange")
        if sembol and fiyat and degisim:
            sembol_text = sembol.get_text(strip=True)
            if sembol_text in gorulmus:
                continue
            gorulmus.add(sembol_text)
            fiyat_str  = fiyat.get_text(strip=True).replace(".", "").replace(",", ".")
            degisim_str = degisim.get_text(strip=True).replace("%", "").replace(".", "").replace(",", ".")
            data.append([sembol_text, fiyat_str, degisim_str])

    df = pd.DataFrame(data, columns=["Sembol", "Fiyat", "Değişim%"])
    df["Fiyat"]    = pd.to_numeric(df["Fiyat"],    errors="coerce")
    df["Değişim%"] = pd.to_numeric(df["Değişim%"], errors="coerce")
    return df

# ----------------------------- ANA İŞLEM -----------------------------
if __name__ == "__main__":
    print("Hisse verileri çekiliyor...")
    df_hisse = hisse_verilerini_cek()
    print(f"{len(df_hisse)} hisse bulundu.")

    print("Fon verileri çekiliyor...")
    df_fon = fon_verilerini_cek()
    print(f"{len(df_fon)} fon bulundu.")

    print("Bloomberg verileri çekiliyor...")
    df_bloomberg = bloomberg_verilerini_cek()
    print(f"{len(df_bloomberg)} Bloomberg verisi bulundu.")

    # Excel: Fon (A) | boş (C) | Hisse (D) | boş (F) | Bloomberg (G)
    col_fon      = 0
    col_hisse    = len(df_fon.columns) + 1
    col_bloomberg = col_hisse + len(df_hisse.columns) + 1
    with pd.ExcelWriter("piyasa_verileri.xlsx", engine="openpyxl") as writer:
        df_fon.to_excel(writer,       sheet_name="Piyasa Verileri", index=False, startcol=col_fon)
        df_hisse.to_excel(writer,     sheet_name="Piyasa Verileri", index=False, startcol=col_hisse)
        df_bloomberg.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col_bloomberg)
    print("Veriler 'piyasa_verileri.xlsx' dosyasına kaydedildi.")

    # JSON çıktısı
    json_data = {
        "hisseler":  df_hisse.to_dict(orient="records"),
        "fonlar":    df_fon.to_dict(orient="records"),
        "bloomberg": df_bloomberg.to_dict(orient="records")
    }
    with open("piyasa_verileri.json", "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)
    print("Veriler 'piyasa_verileri.json' dosyasına kaydedildi.")
