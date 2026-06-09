import undetected_chromedriver as uc
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
import re

# ----------------------------- DRIVER (undetected, headless KALDIRILDI - xvfb kullanılacak) -----------------------------
def selenium_driver_olustur():
    """
    GitHub Actions için optimize edilmiş driver. 
    Headless kullanılmıyor, xvfb ile çalışacak.
    """
    print("Driver oluşturuluyor (undetected-chromedriver)...")
    options = uc.ChromeOptions()
    # Headless kaldırıldı, çünkü xvfb kullanacağız
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    options.add_argument("--disable-gpu")
    options.add_argument("--window-size=1920,1080")
    # İndirme ayarları
    download_dir = os.path.join(os.getcwd(), "downloads")
    os.makedirs(download_dir, exist_ok=True)
    prefs = {
        "download.default_directory": download_dir,
        "download.prompt_for_download": False,
        "download.directory_upgrade": True,
        "safebrowsing.enabled": True
    }
    options.add_experimental_option("prefs", prefs)
    options.add_argument("--disable-blink-features=AutomationControlled")
    options.add_argument("user-agent=Mozilla/5.0 (X11; Linux x86_64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/148.0.0.0 Safari/537.36")
    
    driver = uc.Chrome(options=options, version_main=148)  # Chrome ana sürümünüz neyse onu yazın
    driver.set_page_load_timeout(120)
    driver.set_script_timeout(120)
    print("Driver hazır.")
    return driver, download_dir

# ----------------------------- HİSSE VERİLERİ (Excel'e Aktar butonu) -----------------------------
def turkce_sayi_cevir(deger):
    if isinstance(deger, str):
        deger = deger.replace('.', '').replace(',', '.')
    try:
        return float(deger)
    except:
        return None

def temizle_hisse(metin):
    if isinstance(metin, str):
        metin = re.sub(r'\s+', ' ', metin).strip()
        metin = metin.replace('\u200b', '').replace('\u200c', '').replace('\u200d', '')
        metin = metin.replace('\uFEFF', '').replace('\u00A0', ' ')
        return metin.strip()
    return metin

def hisse_verilerini_cek():
    max_deneme = 2
    for deneme in range(1, max_deneme + 1):
        driver = None
        try:
            driver, download_dir = selenium_driver_olustur()
            print(f"Hisse verisi deneme {deneme}/{max_deneme}...")
            driver.get("https://www.isyatirim.com.tr/tr-tr/analiz/hisse/Sayfalar/default.aspx")
            # Sayfada excel butonunu bekle (30 saniye)
            excel_btn = WebDriverWait(driver, 30).until(
                EC.element_to_be_clickable((By.CSS_SELECTOR, "a.excelimage"))
            )
            # İndirme öncesi mevcut dosyalar
            before = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
            excel_btn.click()
            # Yeni dosyayı bekle
            timeout = 30
            waited = 0
            downloaded_file = None
            while waited < timeout:
                after = set(glob.glob(os.path.join(download_dir, "*.xlsx")))
                new_files = after - before
                if new_files:
                    downloaded_file = max(new_files, key=os.path.getctime)
                    break
                time.sleep(1)
                waited += 1
            if not downloaded_file:
                raise Exception("Excel dosyası indirilemedi.")
            df = pd.read_excel(downloaded_file)
            if df.shape[1] >= 2:
                df = df.iloc[:, [0, 1]]
                df.columns = ["Hisse", "Son Fiyat (TL)"]
                df["Hisse"] = df["Hisse"].astype(str).apply(temizle_hisse)
                df["Son Fiyat (TL)"] = df["Son Fiyat (TL)"].apply(turkce_sayi_cevir)
                df = df.dropna(subset=["Hisse"])
                df = df[df["Hisse"].str.strip() != ""]
                print(f"Excel'den {len(df)} hisse alındı (temizlendi).")
                os.remove(downloaded_file)
                return df
            else:
                raise Exception("Excel sütun hatası")
        except Exception as e:
            print(f"Deneme {deneme} başarısız: {e}")
            if deneme == max_deneme:
                return pd.DataFrame(columns=["Hisse", "Son Fiyat (TL)"])
            time.sleep(5)
        finally:
            if driver:
                driver.quit()
    return pd.DataFrame(columns=["Hisse", "Son Fiyat (TL)"])

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
        print(f"Fon verisi deneniyor: {tarih}")
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
                continue
            data = response.json()
            if "resultList" in data and data["resultList"]:
                fonlar = data["resultList"]
                kayitlar = [[fon["fonKodu"], fon["fiyat"]] for fon in fonlar if fon.get("fiyat") is not None]
                df = pd.DataFrame(kayitlar, columns=["Fon Kodu", "Fiyat"])
                df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
                print(f"{tarih} tarihinde {len(df)} fon alındı.")
                return df
        except Exception as e:
            print(f"Hata: {e}")
            continue
    print("Fon verisi bulunamadı.")
    return pd.DataFrame(columns=["Fon Kodu", "Fiyat"])

# ----------------------------- BLOOMBERG VERİLERİ (undetected driver ile) -----------------------------
def bloomberg_verilerini_cek():
    from bs4 import BeautifulSoup
    max_deneme = 2
    for deneme in range(1, max_deneme + 1):
        driver = None
        try:
            driver, _ = selenium_driver_olustur()
            print(f"Bloomberg deneme {deneme}...")
            driver.get("https://www.bloomberght.com/piyasalar")
            WebDriverWait(driver, 30).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "div.swiper-slide[data-swiper-slide-index]"))
            )
            time.sleep(2)
            soup = BeautifulSoup(driver.page_source, "html.parser")
            slides = soup.select("div.swiper-slide[data-swiper-slide-index]")
            if not slides:
                raise Exception("Slayt yok")
            gorulmus = set()
            data = []
            for slide in slides:
                sembol = slide.select_one("span.text-xs.text-ellipsis")
                fiyat = slide.select_one("span.lastPrice")
                degisim = slide.select_one("span.percentChange")
                if sembol and fiyat and degisim:
                    s = sembol.get_text(strip=True)
                    if s not in gorulmus:
                        gorulmus.add(s)
                        f = fiyat.get_text(strip=True).replace(".", "").replace(",", ".")
                        d = degisim.get_text(strip=True).replace("%", "").replace(".", "").replace(",", ".")
                        data.append([s, f, d])
            if not data:
                raise Exception("Veri yok")
            df = pd.DataFrame(data, columns=["Sembol", "Fiyat", "Değişim%"])
            df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
            df["Değişim%"] = pd.to_numeric(df["Değişim%"], errors="coerce")
            print(f"Bloomberg'den {len(df)} veri alındı.")
            return df
        except Exception as e:
            print(f"Bloomberg deneme {deneme} hata: {e}")
            if deneme == max_deneme:
                return pd.DataFrame(columns=["Sembol", "Fiyat", "Değişim%"])
            time.sleep(5)
        finally:
            if driver:
                driver.quit()

# ----------------------------- ANA -----------------------------
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

    # Excel çıktısı
    with pd.ExcelWriter("piyasa_verileri.xlsx", engine="openpyxl") as writer:
        col = 0
        if not df_fon.empty:
            df_fon.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col)
            col += len(df_fon.columns) + 1
        if not df_hisse.empty:
            df_hisse.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col)
            col += len(df_hisse.columns) + 1
        if not df_bloomberg.empty:
            df_bloomberg.to_excel(writer, sheet_name="Piyasa Verileri", index=False, startcol=col)
    print("Excel kaydedildi.")

    with open("piyasa_verileri.json", "w", encoding="utf-8") as f:
        json.dump({
            "hisseler": df_hisse.to_dict(orient="records"),
            "fonlar": df_fon.to_dict(orient="records"),
            "bloomberg": df_bloomberg.to_dict(orient="records")
        }, f, ensure_ascii=False, indent=4)
    print("JSON kaydedildi.")
