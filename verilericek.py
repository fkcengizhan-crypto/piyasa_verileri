from selenium import webdriver
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.chrome.service import Service
from bs4 import BeautifulSoup
import pandas as pd
import requests
import time
import re
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
def _tefas_token_al(session: requests.Session, user_agent: str) -> str | None:
    """
    TEFAS ana sayfasını ziyaret ederek dinamik Bearer token'ı bulmaya çalışır.
    Token JavaScript içinde veya meta tag'de gömülü olabilir.
    Bulunamazsa None döner.
    """
    try:
        resp = session.get(
            "https://www.tefas.gov.tr/tr/fon-verileri",
            headers={"User-Agent": user_agent},
            timeout=30,
        )
        # Sayfada "ST-..." formatındaki token'ı ara
        m = re.search(r'["\']?(ST-[A-Za-z0-9]+)["\']?', resp.text)
        if m:
            token = m.group(1)
            print(f"  TEFAS token bulundu: {token[:20]}...")
            return token
    except Exception as e:
        print(f"  TEFAS ana sayfa ziyareti başarısız: {e}")
    return None


def _tefas_api_cagir(session: requests.Session, headers: dict, tarih: str) -> list | None:
    """
    Verilen tarih için TEFAS API'sini çağırır.
    Başarılıysa resultList döner, aksi halde None döner.
    """
    url = "https://www.tefas.gov.tr/api/funds/fonGnlBlgSiraliGetirDosya"
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
        "sfonTurKod": None,
    }
    try:
        response = session.post(url, headers=headers, json=payload, timeout=30)
        response.raise_for_status()
        data = response.json()
        result = data.get("resultList")
        if result:
            return result
    except Exception as e:
        print(f"  TEFAS API çağrısı başarısız ({tarih}): {e}")
    return None


def fon_verilerini_cek():
    """
    TEFAS'tan fon verilerini çeker.
    - Önce dinamik olarak Bearer token almaya çalışır.
    - Token bulunamazsa token'sız dener.
    - Bugünün verisi yoksa (tatil/hafta sonu) geriye doğru en fazla 5 iş günü dener.
    """
    user_agent = (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/147.0.0.0 Safari/537.36"
    )
    session = requests.Session()
    session.headers.update({"User-Agent": user_agent})

    # Dinamik token almayı dene
    token = _tefas_token_al(session, user_agent)

    api_headers = {
        "Accept": "application/json, text/plain, */*",
        "Content-Type": "application/json",
        "Origin": "https://www.tefas.gov.tr",
        "Referer": "https://www.tefas.gov.tr/tr/fon-verileri",
        "User-Agent": user_agent,
    }
    if token:
        api_headers["Authorization"] = f"Bearer {token}"
    else:
        print("  Token bulunamadı, token'sız deneniyor...")

    # Bugünden geriye 5 güne kadar dene (hafta sonu / tatil durumu)
    fonlar = None
    for gun_geriye in range(6):
        hedef_tarih = (datetime.now() - timedelta(days=gun_geriye)).strftime("%Y%m%d")
        print(f"  {hedef_tarih} tarihi için TEFAS verisi deneniyor...")
        fonlar = _tefas_api_cagir(session, api_headers, hedef_tarih)
        if fonlar:
            print(f"  {hedef_tarih} tarihinde {len(fonlar)} fon bulundu.")
            break

    if not fonlar:
        raise Exception(
            "TEFAS'tan fon verisi alınamadı. "
            "Son 5 gün için API yanıtı boş geldi. "
            "Endpoint veya token formatı değişmiş olabilir."
        )

    kayitlar = [
        [fon["fonKodu"], fon["fiyat"]]
        for fon in fonlar
        if fon.get("fiyat") is not None
    ]
    df = pd.DataFrame(kayitlar, columns=["Fon Kodu", "Fiyat"])
    df["Fiyat"] = pd.to_numeric(df["Fiyat"], errors="coerce")
    return df


# ----------------------------- BLOOMBERG VERİLERİ -----------------------------
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
            fiyat_str   = fiyat.get_text(strip=True).replace(".", "").replace(",", ".")
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
    col_fon       = 0
    col_hisse     = len(df_fon.columns) + 1
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
        "bloomberg": df_bloomberg.to_dict(orient="records"),
    }
    with open("piyasa_verileri.json", "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=4)
    print("Veriler 'piyasa_verileri.json' dosyasına kaydedildi.")
