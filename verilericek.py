from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.by import By
import pandas as pd
import requests
from requests.adapters import HTTPAdapter
from urllib3.util.retry import Retry
import time
from datetime import datetime, timedelta
import json
import os
import glob

# ----------------------------- GELİŞMİŞ SELENIUM DRIVER (indirme desteği ile) -----------------------------
def selenium_driver_olustur():
    chrome_options = Options()
    chrome_options.add_argument("--headless=new")
    chrome_options.add_argument("--no-sandbox")
    chrome_options.add_argument("--disable-dev-shm-usage")
    chrome_options.add_argument("--disable-gpu")
    chrome_options.add_argument("--disable-software-rasterizer")
    chrome_options.add_argument("--disable-features=VizDisplayCompositor")
    chrome_options.add_argument("--disable-features=NetworkService,NetworkServiceInProcess")
    chrome_options.add_argument("--disable-blink-features=AutomationControlled")
    chrome_options.add_argument("--disable-background-timer-throttling")
    chrome_options.add_argument("--disable-renderer-backgrounding")
    chrome_options.add_argument("--disable-backgrounding-occluded-windows")
    chrome_options.add_argument("--disable-breakpad")
    chrome_options.add_argument("--disable-component-extensions-with-background-pages")
    chrome_options.add_argument("--disable-default-apps")
    chrome_options.add_argument("--disable-notifications")
    chrome_options.add_argument("--disable-popup-blocking")
    chrome_options.add_argument("--no-first-run")
    chrome_options.add_argument("--safebrowsing-disable-auto-update")
    chrome_options.add_argument("--window-size=1920,1080")
    chrome_options.add_argument("--remote-debugging-port=0")
    chrome_options.add_argument("--disable-setuid-sandbox")
    chrome_options.add_argument("--disable-infobars")
    chrome_options.add_argument("--ignore-certificate-errors")
    chrome_options.add_argument(
        "user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 "
        "(KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36"
    )
    # İndirme ayarları
    download_dir = os.path.join(os.getcwd(), "downloads")
    os.makedirs(download_dir, exist_ok=True)
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    chrome_options.add_experimental_option("prefs", prefs)
    chrome_options.page_load_strategy = 'normal'

    chrome_binary = os.environ.get("CHROME_BINARY_PATH")
    if chrome_binary:
        chrome_options.binary_location = chrome_binary

    driver_path = os.environ.get("CHROME_DRIVER_PATH")
    if driver_path:
        service = Service(driver_path)
        driver = webdriver.Chrome(service=service, options=chrome_options)
    else:
        driver = webdriver.Chrome(options=chrome_options)

    driver.set_page_load_timeout(120)
    driver.set_script_timeout(120)
    return driver, download_dir

# ----------------------------- HİSSE VERİLERİ (Excel'e Aktar butonu ile) -----------------------------
def turkce_sayi_cevir(deger):
    if isinstance(deger, str):
        deger = deger.replace('.', '')
        deger = deger.replace(',', '.')
    try:
        return float(deger)
    except (ValueError, TypeError):
        return None

def hisse_verilerini_cek():
    # ... driver ayarları, butona tıklama, dosya indirme ...
    # Excel dosyasını oku
    df = pd.read_excel(downloaded_file)
    if df.shape[1] >= 2:
        df = df.iloc[:, [0, 1]]
        df.columns = ["Hisse", "Son Fiyat (TL)"]
        
        # ---- TEMİZLİK BAŞLIYOR ----
        import re
        def temizle_hisse(metin):
            if isinstance(metin, str):
                metin = re.sub(r'\s+', ' ', metin)      # Çoklu boşlukları tek boşluğa çevir
                metin = metin.strip()                   # Baş-son boşlukları sil
                metin = metin.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
                # İsterseniz sadece harf/rakam/nokta/kısa çizgi bırakabilirsiniz:
                # metin = re.sub(r'[^A-Za-z0-9\-\.]', '', metin)
            return metin
        
        df["Hisse"] = df["Hisse"].apply(temizle_hisse)
        df["Son Fiyat (TL)"] = df["Son Fiyat (TL)"].apply(turkce_sayi_cevir)
        df = df.dropna(subset=["Hisse"])
        df = df[df["Hisse"].str.strip() != ""]
        # ---- TEMİZLİK BİTTİ ----
        
        print(f"Excel'den {len(df)} hisse verisi alındı (temizlendi).")
        os.remove(downloaded_file)
        return df
    else:
        raise Exception("Excel dosyası beklenen sütunları içermiyor.")

# ----------------------------- FON VERİLERİ (değişmedi) -----------------------------
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
    session = requests.Session()
    retries = Retry(total=2, backoff_factor=1, status_forcelist=[500, 502, 503, 504])
    session.mount('https://', HTTPAdapter(max_retries=retries))

    for i in range(15):
        tarih = (datetime.now() - timedelta(days=i)).strftime("%Y%m%d")
        print(f"Fon verisi için tarih deneniyor: {tarih}")
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
            response = session.post(url, headers=headers, json=payload, timeout=(10, 60))
            if response.status_code != 200:
                print(f"HTTP hatası: {response.status_code}")
                continue
            data = response.json()
            if "resultList" not in data or not data["resultList"]:
                print(f"{tarih} için fon verisi bulunamadı.")
                continue
            fonlar = data["resultList"]
            kayitlar = [[fon["fonKodu"], fon["fiyat"]] for fon in fonlar if fon.get("fiyat") is not None]
            df = pd.DataFrame(kayitlar, columns=["Fon Kodu", "Fiyat"])
            df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
            print(f"{tarih} tarihine ait {len(df)} fon verisi alındı.")
            return df
        except Exception as e:
            print(f"{tarih} için hata: {e}")
            continue

    print("UYARI: Son 15 günde fon verisi bulunamadı, boş DataFrame dönülüyor.")
    return pd.DataFrame(columns=["Fon Kodu", "Fiyat"])

# ----------------------------- BLOOMBERG VERİLERİ (değişmedi) -----------------------------
def bloomberg_verilerini_cek():
    max_deneme = 3
    for deneme in range(1, max_deneme + 1):
        driver = None
        try:
            # Bloomberg için ayrı driver (indirme klasörü gerekmez)
            driver, _ = selenium_driver_olustur()
            print(f"Bloomberg verisi deneme {deneme}/{max_deneme}...")
            driver.get("https://www.bloomberght.com/piyasalar")
            WebDriverWait(driver, 90).until(
                lambda d: d.execute_script("return document.readyState") == "complete"
            )
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.swiper-slide[data-swiper-slide-index]"))
            )
            time.sleep(2)
            from bs4 import BeautifulSoup
            soup = BeautifulSoup(driver.page_source, "html.parser")
            slides = soup.select("div.swiper-slide[data-swiper-slide-index]")
            if not slides:
                raise Exception("Bloomberg slaytları bulunamadı.")
            gorulmus = set()
            data = []
            for slide in slides:
                sembol = slide.select_one("span.text-xs.text-ellipsis")
                fiyat = slide.select_one("span.lastPrice")
                degisim = slide.select_one("span.percentChange")
                if sembol and fiyat and degisim:
                    sembol_text = sembol.get_text(strip=True)
                    if sembol_text in gorulmus:
                        continue
                    gorulmus.add(sembol_text)
                    fiyat_str = fiyat.get_text(strip=True).replace(".", "").replace(",", ".")
                    degisim_str = degisim.get_text(strip=True).replace("%", "").replace(".", "").replace(",", ".")
                    data.append([sembol_text, fiyat_str, degisim_str])
            if not data:
                raise Exception("Hiç Bloomberg verisi çekilemedi (içerik boş).")
            df = pd.DataFrame(data, columns=["Sembol", "Fiyat", "Değişim%"])
            df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
            df["Değişim%"] = pd.to_numeric(df["Değişim%"], errors="coerce")
            print(f"Bloomberg'den {len(df)} veri alındı.")
            return df
        except Exception as e:
            print(f"Bloomberg deneme {deneme} başarısız: {e}")
            if deneme == max_deneme:
                print("Bloomberg verileri alınamadı, boş DataFrame ile devam ediliyor.")
                return pd.DataFrame(columns=["Sembol", "Fiyat", "Değişim%"])
            time.sleep(10)
        finally:
            if driver:
                driver.quit()

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

    # Excel çıktısı (aynı sayfada yan yana)
    col_fon = 0
    col_hisse = len(df_fon.columns) + 1 if not df_fon.empty else 0
    col_bloomberg = col_hisse + len(df_hisse.columns) + 1 if not df_hisse.empty else 0
    with pd.ExcelWriter("piyasa_verileri.xlsx", engine="openpyxl") as writer:
        if not df_fon.empty:
            df_fon.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col_fon)
        if not df_hisse.empty:
            df_hisse.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col_hisse)
        if not df_bloomberg.empty:
            df_bloomberg.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col_bloomberg)
    print("Veriler 'piyasa_verileri.xlsx' dosyasına kaydedildi.")

    json_data = {
        "hisseler": df_hisse.to_dict(orient="records"),
        "fonlar": df_fon.to_dict(orient="records"),
        "bloomberg": df_bloomberg.to_dict(orient="records")
    }
    with open("piyasa_verileri.json", "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)
    print("Veriler 'piyasa_verileri.json' dosyasına kaydedildi.")
