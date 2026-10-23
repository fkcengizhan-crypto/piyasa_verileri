import requests
import pandas as pd
from datetime import datetime, timedelta, timezone
from openpyxl import load_workbook
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment, Font
import json
import gzip
import re
import socket
import ssl
import base64
import os
import struct
from pathlib import Path
from bs4 import BeautifulSoup

BASE_DIR = Path(__file__).resolve().parent


# ========== YARDIMCI FONKSİYONLAR ==========
def _yuvarla(deger, basamak=2):
    """Sayıyı güvenli biçimde yuvarlar; boş/geçersiz değerde None döner."""
    if deger is None:
        return None
    try:
        if isinstance(deger, str) and not deger.strip():
            return None
        if pd.isna(deger):
            return None
        return round(float(deger), basamak)
    except (TypeError, ValueError):
        return None


def _degisim_hesapla(yeni, eski, basamak=2):
    """(yeni - eski) / eski * 100. Önceki değer yoksa veya sıfırsa None döner."""
    try:
        yeni = float(yeni)
        eski = float(eski)
    except (TypeError, ValueError):
        return None
    if not eski:
        return None
    return round((yeni - eski) / eski * 100, basamak)


def _json_sayi(deger):
    """JSON çıktısı için güvenli sayı dönüşümü."""
    if deger is None:
        return None
    try:
        if pd.isna(deger):
            return None
    except (TypeError, ValueError):
        pass
    if isinstance(deger, (int, float)):
        return float(deger)
    return _sayi_cevir(str(deger))


# ========== HİSSE FONKSİYONLARI (TradingView) ==========
def fetch_all_stocks():
    url = "https://scanner.tradingview.com/turkey/scan?label-product=screener-stock"
    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}

    payload = {
        "columns": ["name", "description", "close", "change"],
        "filter": [],
        "ignore_unknown_fields": False,
        "options": {"lang": "tr"},
        "range": [0, 999],
        "sort": {"sortBy": "market_cap_basic", "sortOrder": "desc"},
        "markets": ["turkey"],
        "filter2": {
            "operator": "and",
            "operands": [
                {
                    "operation": {
                        "operator": "or",
                        "operands": [
                            {
                                "operation": {
                                    "operator": "and",
                                    "operands": [
                                        {"expression": {"left": "type", "operation": "equal", "right": "stock"}},
                                        {"expression": {"left": "typespecs", "operation": "has", "right": ["common"]}}
                                    ]
                                }
                            },
                            {
                                "operation": {
                                    "operator": "and",
                                    "operands": [
                                        {"expression": {"left": "type", "operation": "equal", "right": "stock"}},
                                        {"expression": {"left": "typespecs", "operation": "has", "right": ["preferred"]}}
                                    ]
                                }
                            },
                            {
                                "operation": {
                                    "operator": "and",
                                    "operands": [
                                        {"expression": {"left": "type", "operation": "equal", "right": "dr"}}
                                    ]
                                }
                            },
                            {
                                "operation": {
                                    "operator": "and",
                                    "operands": [
                                        {"expression": {"left": "type", "operation": "equal", "right": "fund"}},
                                        {"expression": {"left": "typespecs", "operation": "has_none_of", "right": ["etf", "mutual"]}}
                                    ]
                                }
                            }
                        ]
                    }
                },
                {
                    "expression": {"left": "typespecs", "operation": "has_none_of", "right": ["pre-ipo"]}
                }
            ]
        }
    }

    all_stocks = []

    def extract_stocks(data):
        count_before = len(all_stocks)
        for item in data.get("data", []):
            try:
                d = item.get("d", [])
                if len(d) < 3:
                    continue
                if isinstance(d[0], str) and d[0]:
                    code = d[0]
                    name = d[1] if d[1] else ""
                    price = d[2]
                    degisim = d[3] if len(d) > 3 else None
                    all_stocks.append({
                        "Hisse Kodu": code,
                        "Hisse Adı": name,
                        "Hisse Fiyatı": price,
                        "Günlük Değişim (%)": _yuvarla(degisim)
                    })
            except Exception as e:
                print(f"  ⚠ Hisse atlandı: {e}")
        added = len(all_stocks) - count_before
        if added > 0:
            print(f"  → {added} adet hisse bulundu.")

    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        total = data.get("totalCount", 0)
        extract_stocks(data)

        page_size = 999
        for start in range(page_size, total, page_size):
            payload["range"] = [start, start + page_size - 1]
            print(f"Sayfa {start // page_size + 1} çekiliyor...", end=" ")
            try:
                resp = requests.post(url, json=payload, headers=headers, timeout=30)
                resp.raise_for_status()
                page_data = resp.json()
                extract_stocks(page_data)
            except Exception as e:
                print(f"\n  ⚠ Sayfa hatası: {e}")

    except Exception as e:
        print(f"Ana hata: {e}")
        if all_stocks:
            print(f"{len(all_stocks)} hisseyle devam ediliyor...")
            return all_stocks
        return None

    return all_stocks if all_stocks else None


# ========== FON FONKSİYONLARI (TEFAS) ==========
def fon_turu_esleme_al():
    url = "https://www.tefas.gov.tr/api/fund-returns/export"
    payload = {
        "format": "json",
        "listingType": "return",
        "fundType": "YAT",
        "locale": "tr",
        "filters": {
            "kurucuKodu": None, "fonTurKod": None, "fonGrubu": None,
            "fonTurAciklama": None, "sfonTurKod": None, "islem": 1,
            "basTarih": "20260601", "bitTarih": "20260601",
            "calismaTipi": 1,
            "donemGetiri1a": "0", "donemGetiri3a": "0", "donemGetiri6a": "0",
            "donemGetiriyb": "0", "donemGetiri1y": "0", "donemGetiri3y": "0",
            "donemGetiri5y": "0", "getiriOrani": "1"
        },
        "columns": ["fonKodu", "fonTurAciklama"]
    }
    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        df = pd.DataFrame(data)
        df = df[["fonKodu", "fonTurAciklama"]].drop_duplicates()
        return dict(zip(df["fonKodu"], df["fonTurAciklama"]))
    except Exception as e:
        print(f"Fon türü alınamadı: {e}")
        return {}

def son_fiyatlari_cek(tarih_str):
    url = "https://www.tefas.gov.tr/api/funds/fonGnlBlgSiraliGetir"
    headers = {"Content-Type": "application/json", "User-Agent": "Mozilla/5.0"}
    payload = {
        "fonTipi": "YAT",
        "basTarih": tarih_str,
        "bitTarih": tarih_str,
        "basSira": 1,
        "bitSira": 5000,
        "dil": "TR"
    }
    try:
        resp = requests.post(url, json=payload, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json()
        return data.get("resultList", [])
    except Exception as e:
        print(f"Fiyat çekme hatası: {e}")
        return []

def _fiyat_haritasi(tarih_str):
    """Belirtilen tarihin fon kodu → fiyat haritasını döndürür."""
    veri = son_fiyatlari_cek(tarih_str)
    return {
        r.get("fonKodu"): r.get("fiyat")
        for r in veri
        if r.get("fonKodu") and r.get("fiyat") is not None
    }


def fetch_all_funds():
    tarih_denenecek = []
    bugun = datetime.now().date()
    for i in range(3):
        tarih_denenecek.append(bugun.strftime("%Y%m%d"))
        bugun = bugun - timedelta(days=1)

    tur_map = fon_turu_esleme_al()

    sonuclar = []
    veri_tarihi = None
    for tarih_str in tarih_denenecek:
        veri = son_fiyatlari_cek(tarih_str)
        if veri:
            print(f"  → {len(veri)} fon bulundu ({tarih_str}).")
            sonuclar = veri
            veri_tarihi = tarih_str
            break
        else:
            print("✗ Veri yok.")

    if not sonuclar:
        print("Son 3 günde hiç veri bulunamadı.")
        return None

    # Günlük değişim için bir önceki işlem gününün fiyatları (hafta sonu/tatil toleransı: 10 gün)
    onceki_fiyatlar = {}
    gun = datetime.strptime(veri_tarihi, "%Y%m%d").date() - timedelta(days=1)
    for _ in range(10):
        harita = _fiyat_haritasi(gun.strftime("%Y%m%d"))
        if harita:
            onceki_fiyatlar = harita
            print(f"  → Önceki fiyatlar {gun.isoformat()} tarihinden alındı.")
            break
        gun = gun - timedelta(days=1)

    if not onceki_fiyatlar:
        print("  ⚠ Önceki gün fiyatları bulunamadı, günlük değişim boş kalacak.")

    df = pd.DataFrame(sonuclar)
    df = df[["fonKodu", "fonUnvan", "fiyat"]].rename(columns={
        "fonKodu": "Fon Kodu",
        "fonUnvan": "Fon Adı",
        "fiyat": "Fon Fiyatı"
    })
    df["Günlük Değişim (%)"] = [
        _degisim_hesapla(fiyat, onceki_fiyatlar.get(kod))
        for kod, fiyat in zip(df["Fon Kodu"], df["Fon Fiyatı"])
    ]
    df["Fon Türü"] = df["Fon Kodu"].map(tur_map).fillna("Bilinmiyor")

    df = df[df["Fon Türü"] != "Bilinmiyor"]
    df = df[~df["Fon Türü"].str.contains("Serbest", na=False)]

    if df.empty:
        print("Filtreleme sonrası hiç fon kalmadı.")
        return None

    df = df[["Fon Kodu", "Fon Adı", "Fon Türü", "Fon Fiyatı", "Günlük Değişim (%)"]]
    return df


# ========== BIST TLREF ENDEKSİ FONLARI (FINTABLES / BYF) ==========
# Fintables canli akisina standart kutuphane ile (socket + ssl) baglanir.
# Harici paket GEREKTIRMEZ (workflow'da ek kurulum gerekmez).
def _ws_baglan(host, path, timeout=15):
    ctx = ssl.create_default_context()
    raw = socket.create_connection((host, 443), timeout=timeout)
    sock = ctx.wrap_socket(raw, server_hostname=host)
    key = base64.b64encode(os.urandom(16)).decode("ascii")
    istek = (
        "GET " + path + " HTTP/1.1\r\n"
        "Host: " + host + "\r\n"
        "Upgrade: websocket\r\n"
        "Connection: Upgrade\r\n"
        "Sec-WebSocket-Key: " + key + "\r\n"
        "Sec-WebSocket-Version: 13\r\n"
        "User-Agent: Mozilla/5.0\r\n"
        "\r\n"
    )
    sock.sendall(istek.encode("ascii"))
    f = sock.makefile("rb")
    durum = f.readline()
    if b"101" not in durum:
        sock.close()
        raise ConnectionError("WebSocket el sikismasi basarisiz")
    while True:
        satir = f.readline()
        if satir in (b"\r\n", b"\n", b""):
            break
    return sock, f


def _ws_frame_oku(f):
    bas = f.read(2)
    if len(bas) < 2:
        raise ConnectionError("baglanti kapandi")
    opcode = bas[0] & 0x0F
    uzunluk = bas[1] & 0x7F
    if uzunluk == 126:
        uzunluk = struct.unpack(">H", f.read(2))[0]
    elif uzunluk == 127:
        uzunluk = struct.unpack(">Q", f.read(8))[0]
    veri = f.read(uzunluk) if uzunluk else b""
    return opcode, veri


def _ws_gonder(sock, metin):
    veri = metin.encode("utf-8")
    maske = os.urandom(4)
    cerceve = bytearray([0x81])
    if len(veri) < 126:
        cerceve.append(0x80 | len(veri))
    elif len(veri) < 65536:
        cerceve.append(0x80 | 126)
        cerceve += struct.pack(">H", len(veri))
    else:
        cerceve.append(0x80 | 127)
        cerceve += struct.pack(">Q", len(veri))
    cerceve += maske
    cerceve += bytes(veri[i] ^ maske[i % 4] for i in range(len(veri)))
    sock.sendall(bytes(cerceve))


def fetch_byf_fonlar():
    """
    Fintables canli akisindan NPTLR.F ve OPTLR.F fonlarinin son fiyatini ceker.
    tbl_fonfiyatlari tablosunun sonuna eklenecek satirlari dondurur.
    Günlük değişim, mesajdaki C (son) ve P (önceki kapanış) değerlerinden hesaplanır.
    Doner: DataFrame [Fon Kodu, Fon Adı, Fon Türü, Fon Fiyatı, Günlük Değişim (%)] veya None
    """
    kodlar = ["NPTLR.F", "OPTLR.F"]
    fon_adi = "BIST TLREF ENDEKSİ (TL) BORSA YATIRIM FONU"
    fon_turu = "Bist TLref Endeksi"

    topics = []
    for kod in kodlar:
        topics += [kod + "/C", kod + "/P", kod + "/T"]

    fiyatlar = {}
    onceki_fiyatlar = {}
    sock = None
    try:
        sock, akis = _ws_baglan("markets.fintables.com", "/vessel", timeout=15)
        _ws_gonder(sock, json.dumps({"type": "subscribe", "topics": topics}))
        for _ in range(20):
            opcode, veri = _ws_frame_oku(akis)
            if opcode == 0x8:
                break
            if opcode not in (0x1, 0x2):
                continue
            try:
                data = json.loads(veri.decode("utf-8"))
            except Exception:
                continue
            if data.get("type") == "subscribe_success":
                for item in data.get("data", []):
                    kod = item.get("code")
                    if kod and item.get("C") is not None:
                        fiyatlar[kod] = item.get("C")
                        onceki_fiyatlar[kod] = item.get("P")  # önceki kapanış
                break
    except Exception as e:
        print(f"  ⚠ Fintables fon verisi alınamadı: {str(e)[:60]}")
    finally:
        if sock is not None:
            try:
                sock.close()
            except Exception:
                pass

    if not fiyatlar:
        return None

    satirlar = []
    for kod in kodlar:
        if kod in fiyatlar:
            satirlar.append({
                "Fon Kodu": kod,
                "Fon Adı": fon_adi,
                "Fon Türü": fon_turu,
                "Fon Fiyatı": fiyatlar[kod],
                "Günlük Değişim (%)": _degisim_hesapla(fiyatlar[kod], onceki_fiyatlar.get(kod))
            })
        else:
            print(f"  ⚠ {kod} için fiyat bulunamadı.")

    if not satirlar:
        return None
    return pd.DataFrame(satirlar)


# ========== DOVIZ.COM VERİLERİ (STATİK HTML) ==========
def format_price(price_str):
    """
    Ham fiyat string'ini düzenler:
    - Eğer zaten virgül ile ondalık ayrılmışsa (örn: 14.421,15) aynen döndür.
    - Eğer nokta ile ayrılmışsa son noktayı virgüle çevir (örn: 14.421.15 -> 14.421,15)
    - Eğer sadece nokta yoksa aynen döndür (örn: 65.820 -> 65.820)
    """
    if not price_str:
        return "Veri yok"
    
    # Önce varsa dolar işaretini temizle
    price_str = price_str.replace('$', '').strip()
    
    # Eğer zaten virgül varsa (Türk formatı) direkt döndür
    if ',' in price_str:
        return price_str
    
    # Nokta ile ayrılmışsa son noktayı virgül yap
    parts = price_str.split('.')
    if len(parts) > 1:
        # Son parça ondalık kısım, diğerleri binlik
        integer_part = '.'.join(parts[:-1])
        decimal_part = parts[-1]
        return f"{integer_part},{decimal_part}"
    else:
        # Nokta yoksa aynen döndür (tam sayı)
        return price_str

def _sayi_cevir(metin):
    """'%0,05', '1.234,56' gibi Türkçe biçimli metinleri sayıya çevirir."""
    if metin is None:
        return None
    metin = str(metin).replace('%', '').replace('$', '').strip()
    if not metin:
        return None
    if ',' in metin:
        metin = metin.replace('.', '').replace(',', '.')
    try:
        return float(metin)
    except ValueError:
        return None


def _api_info_degisim(html, asset_key):
    """
    Sayfadaki `var apiInfo = [...]` bloğundan ilgili varlığın günlük % değişimini çeker.
    Bulunamazsa None döner; çağıran taraf data-socket-attr="c" ile yedekler.
    """
    eslesme = re.search(r'var\s+apiInfo\s*=\s*(\[.*?\])\s*;', html, re.S)
    if not eslesme:
        return None
    try:
        veriler = json.loads(eslesme.group(1))
    except Exception:
        return None
    for item in veriler:
        if item.get('assetKey') == asset_key:
            return _yuvarla(item.get('changeRate'))
    return None


def fetch_doviz_data():
    """
    Doviz.com'dan istenen tüm verileri çeker.
    Döner: List[Dict] - {'Sembol': ..., 'Fiyat': ..., 'Günlük Değişim (%)': ...}
    """
    
    # URL, sembol ve HTML içindeki varlık anahtarı eşleştirmesi
    urls = {
        "BIST 100": ("https://borsa.doviz.com/endeksler/xu100-bist-100", "XU100"),
        "Dolar": ("https://kur.doviz.com/serbest-piyasa/amerikan-dolari", "USD"),
        "Euro": ("https://kur.doviz.com/serbest-piyasa/euro", "EUR"),
        "BTCUSD": ("https://www.doviz.com/kripto-paralar/bitcoin", "d-bitcoin"),
        "Tahvil": ("https://www.doviz.com/tahvil/tr-5-yillik-tahvil", "TAHVIL5Y"),
        "Brent Petrol": ("https://www.doviz.com/emtia/brent-petrol", "BRENT"),
    }

    # Altın verileri: altin.doviz.com/<sembol> alt sayfaları artık boş dönüyor.
    # Bu yüzden hepsi altin.doviz.com ana sayfasındaki tablodan tek seferde çekilir.
    altin_semboller = {
        "Gr Altın": "gram-altin",
        "Çey.Altın": "ceyrek-altin",
        "Tam Altın": "tam-altin",
        "14 Ayar Bilezik": "14-ayar-altin",
        "18 Ayar Bilezik": "18-ayar-altin",
        "22 Ayar Bilezik": "22-ayar-bilezik",
        "Ons Altın": "ons",
    }
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }
    sonuclar = []
    
    for sembol, (url, asset_key) in urls.items():
        try:
            print(f"  → {sembol} çekiliyor...", end=" ")
            resp = requests.get(url, headers=headers, timeout=15)
            resp.raise_for_status()
            
            soup = BeautifulSoup(resp.text, 'html.parser')
            fiyat_raw = None
            
            # 1. Yöntem: data-socket-key ve data-socket-attr="s" olan div'leri ara
            div = soup.find('div', {'data-socket-key': True, 'data-socket-attr': 's'})
            if div:
                fiyat_raw = div.text.strip()
            else:
                # 2. Yöntem: Sayfa başlığından veya h1'den fiyat bilgisini çek
                h1 = soup.find('h1')
                if h1:
                    h1_text = h1.text
                    match = re.search(r'Son\s*\([^)]*\)\s*([\d.,]+)', h1_text)
                    if match:
                        fiyat_raw = match.group(1)
                
                # 3. Yöntem: Meta description'dan dene
                if not fiyat_raw:
                    meta = soup.find('meta', {'name': 'description'})
                    if meta and meta.get('content'):
                        content = meta.get('content')
                        match = re.search(r'([\d.,]+)\s*(?:seviyesinde|dolar|TL|$)', content)
                        if match:
                            fiyat_raw = match.group(1)
                
                # 4. Yöntem: Sayfa başlığı (title) içinden fiyat ara
                if not fiyat_raw:
                    title = soup.find('title')
                    if title:
                        title_text = title.text
                        match = re.search(r'([\d.,]+)\s*(?:TL|$)', title_text)
                        if match:
                            fiyat_raw = match.group(1)
            
            # Günlük değişim: önce apiInfo.changeRate, yoksa data-socket-attr="c"
            degisim = _api_info_degisim(resp.text, asset_key)
            if degisim is None:
                el = soup.find(attrs={"data-socket-key": asset_key, "data-socket-attr": "c"})
                if el and el.get_text(strip=True):
                    degisim = _yuvarla(_sayi_cevir(el.get_text(strip=True)))

            if fiyat_raw:
                # Formatı düzelt (dolar işareti temizlenir, nokta/virgül düzenlenir)
                fiyat = format_price(fiyat_raw)
                sonuclar.append({"Sembol": sembol, "Fiyat": fiyat, "Günlük Değişim (%)": degisim})
                print("✅")
            else:
                sonuclar.append({"Sembol": sembol, "Fiyat": "Bulunamadı", "Günlük Değişim (%)": degisim})
                print("❌ (bulunamadı)")
                
        except Exception as e:
            print(f"❌ Hata: {str(e)[:50]}")
            sonuclar.append({"Sembol": sembol, "Fiyat": "Hata", "Günlük Değişim (%)": None})

    # Altın fiyatlarını ana sayfadan tek seferde çek
    altin_fiyatlar = {}
    altin_degisim = {}
    try:
        resp = requests.get("https://altin.doviz.com/", headers=headers, timeout=15)
        resp.raise_for_status()
        soup = BeautifulSoup(resp.text, 'html.parser')
        for key in set(altin_semboller.values()):
            # Önce "s" (son), yoksa "ask" (satış), yoksa "bid" (alış) değerini kullan
            for attr in ("s", "ask", "bid"):
                el = soup.find(attrs={"data-socket-key": key, "data-socket-attr": attr})
                if el and el.get_text(strip=True):
                    altin_fiyatlar[key] = el.get_text(strip=True)
                    break
            c = soup.find(attrs={"data-socket-key": key, "data-socket-attr": "c"})
            if c and c.get_text(strip=True):
                altin_degisim[key] = _yuvarla(_sayi_cevir(c.get_text(strip=True)))
    except Exception:
        pass

    for sembol, key in altin_semboller.items():
        print(f"  → {sembol} çekiliyor...", end=" ")
        fiyat_raw = altin_fiyatlar.get(key)
        degisim = altin_degisim.get(key)
        if fiyat_raw:
            sonuclar.append({
                "Sembol": sembol,
                "Fiyat": format_price(fiyat_raw),
                "Günlük Değişim (%)": degisim
            })
            print("✅")
        else:
            sonuclar.append({"Sembol": sembol, "Fiyat": "Bulunamadı", "Günlük Değişim (%)": degisim})
            print("❌ (bulunamadı)")
    
    return sonuclar


# ========== ANA PROGRAM ==========
if __name__ == "__main__":
    print("=== BIST HİSSELERİ ÇEKİLİYOR ===")
    stocks = fetch_all_stocks()
    
    print("\n=== FON FİYATLARI ÇEKİLİYOR ===")
    funds = fetch_all_funds()

    print("\n=== BIST TLREF ENDEKSİ FONLARI (FINTABLES) ÇEKİLİYOR ===")
    byf_fonlar = fetch_byf_fonlar()
    if byf_fonlar is not None and not byf_fonlar.empty:
        print(f"  → {len(byf_fonlar)} fon fiyatı bulundu.")
        if funds is not None:
            funds_tablo = pd.concat([funds, byf_fonlar], ignore_index=True)
        else:
            funds_tablo = byf_fonlar
    else:
        funds_tablo = funds

    print("\n=== DOVIZ.COM VERİLERİ ÇEKİLİYOR ===")
    doviz = fetch_doviz_data()
    if doviz:
        print(f"  → {len(doviz)} veri bulundu.")
    else:
        print("  ⚠ Doviz.com verisi alınamadı.")

    if stocks is None and funds_tablo is None and doviz is None:
        print("Hiç veri alınamadı, çıkılıyor.")
        exit()

    dosya = BASE_DIR / "piyasa_verileri.xlsx"
    
    stocks_startcol = None
    doviz_startcol = None
    bos_sutunlar = []  # tablolar arasına eklenen boş sütunlar (0 tabanlı)

    # Excel yazma (Tablolar arasında 1 sütun boşluk)
    with pd.ExcelWriter(dosya, engine='openpyxl') as writer:
        current_col = 0
        
        if funds_tablo is not None:
            funds_tablo.to_excel(writer, sheet_name="Fiyatlar", startrow=0, startcol=current_col, index=False)
            current_col += len(funds_tablo.columns)  # Fon Kodu, Fon Adı, Fon Türü, Fon Fiyatı, Günlük Değişim (%)
            bos_sutunlar.append(current_col)
            current_col += 1  # ⬅️ BOŞ SÜTUN
        
        if stocks is not None:
            stocks_startcol = current_col
            df_stocks = pd.DataFrame(stocks)
            df_stocks = df_stocks.reindex(columns=["Hisse Kodu", "Hisse Adı", "Hisse Fiyatı", "Günlük Değişim (%)"])
            df_stocks.to_excel(writer, sheet_name="Fiyatlar", startrow=0, startcol=current_col, index=False)
            current_col += len(df_stocks.columns)
            bos_sutunlar.append(current_col)
            current_col += 1  # ⬅️ BOŞ SÜTUN
            
        if doviz is not None:
            doviz_startcol = current_col
            df_doviz = pd.DataFrame(doviz)
            df_doviz = df_doviz.reindex(columns=["Sembol", "Fiyat", "Günlük Değişim (%)"])
            df_doviz.to_excel(writer, sheet_name="Fiyatlar", startrow=0, startcol=current_col, index=False)
            current_col += len(df_doviz.columns)

    # Tablo stilleri ve biçimlendirme
    wb = load_workbook(dosya)
    ws = wb["Fiyatlar"]

    DEGISIM_FORMATI = '+0.00"%";-0.00"%";0.00"%"'

    def _degisim_bicimle(sutun_harfi, ilk_satir, son_satir):
        """Değişim sütununu sağa yaslar, % biçimi verir ve pozitif/negatif renklendirir."""
        for row in range(ilk_satir, son_satir + 1):
            cell = ws[f'{sutun_harfi}{row}']
            cell.alignment = Alignment(horizontal='right')
            cell.number_format = DEGISIM_FORMATI
            deger = cell.value
            if isinstance(deger, (int, float)):
                if deger > 0:
                    cell.font = Font(color="FF2E7D32")
                elif deger < 0:
                    cell.font = Font(color="FFC62828")

    def _fiyat_sagla(sutun_harfi, ilk_satir, son_satir):
        for row in range(ilk_satir, son_satir + 1):
            cell = ws[f'{sutun_harfi}{row}']
            cell.alignment = Alignment(horizontal='right')

    # 1. FON TABLOSU
    if funds_tablo is not None:
        max_row_funds = len(funds_tablo) + 1
        tablo_ref_funds = f"A1:{get_column_letter(len(funds_tablo.columns))}{max_row_funds}"
        tbl_funds = Table(displayName="tbl_fonfiyatlari", ref=tablo_ref_funds)
        tbl_funds.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(tbl_funds)
        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 45
        ws.column_dimensions['C'].width = 30
        ws.column_dimensions['D'].width = 12
        ws.column_dimensions['E'].width = 14
        # Fon fiyat sütunu (D) sağa yasla, değişim sütunu (E) biçimlendir
        _fiyat_sagla('D', 2, max_row_funds)
        _degisim_bicimle('E', 2, max_row_funds)

    # 2. HİSSE TABLOSU
    if stocks is not None and stocks_startcol is not None:
        first_col_letter = get_column_letter(stocks_startcol + 1)
        last_col_letter = get_column_letter(stocks_startcol + 4)
        max_row_stocks = len(stocks) + 1
        tablo_ref_stocks = f"{first_col_letter}1:{last_col_letter}{max_row_stocks}"
        tbl_stocks = Table(displayName="tbl_hissefiyatlari", ref=tablo_ref_stocks)
        tbl_stocks.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(tbl_stocks)
        ws.column_dimensions[first_col_letter].width = 12
        ws.column_dimensions[get_column_letter(stocks_startcol + 2)].width = 40
        ws.column_dimensions[get_column_letter(stocks_startcol + 3)].width = 14
        ws.column_dimensions[last_col_letter].width = 14
        # Hisse fiyat sütunu (3. sütun) sağa yasla, değişim sütunu (4. sütun) biçimlendir
        price_col_letter = get_column_letter(stocks_startcol + 3)
        _fiyat_sagla(price_col_letter, 2, max_row_stocks)
        _degisim_bicimle(last_col_letter, 2, max_row_stocks)

    # 3. DOVIZ TABLOSU (Sembol, Fiyat, Günlük Değişim)
    if doviz is not None and doviz_startcol is not None:
        first_col_letter = get_column_letter(doviz_startcol + 1)
        last_col_letter = get_column_letter(doviz_startcol + 3)
        max_row_doviz = len(doviz) + 1
        tablo_ref_doviz = f"{first_col_letter}1:{last_col_letter}{max_row_doviz}"
        tbl_doviz = Table(displayName="tbl_doviz", ref=tablo_ref_doviz)
        tbl_doviz.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium9",
            showFirstColumn=False,
            showLastColumn=False,
            showRowStripes=True,
            showColumnStripes=False
        )
        ws.add_table(tbl_doviz)
        ws.column_dimensions[first_col_letter].width = 18
        ws.column_dimensions[get_column_letter(doviz_startcol + 2)].width = 16
        ws.column_dimensions[last_col_letter].width = 14
        # Fiyat sütunu (2. sütun) sağa yasla, değişim sütunu (3. sütun) biçimlendir
        price_col_letter = get_column_letter(doviz_startcol + 2)
        _fiyat_sagla(price_col_letter, 2, max_row_doviz)
        _degisim_bicimle(last_col_letter, 2, max_row_doviz)

    # Boş sütun genişliklerini ayarla
    for col_idx in bos_sutunlar:
        try:
            ws.column_dimensions[get_column_letter(col_idx + 1)].width = 3
        except Exception:
            pass

    wb.save(dosya)
    print(f"\n✔ '{dosya}' kaydedildi.")

    # ========== JSON DOSYASI ==========
    turkiye_saati = datetime.now(timezone(timedelta(hours=3)))
    # DataFrame içindeki NaN değerler JSON'da null olarak yazılsın (geçersiz NaN üretmemek için)
    fon_kayitlari = (
        json.loads(funds_tablo.to_json(orient='records', force_ascii=False))
        if funds_tablo is not None else []
    )
    json_data = {
        "guncelleme_zamani": turkiye_saati.isoformat(),   # ISO formatında +03:00 eklenir
        "fonlar": fon_kayitlari,
        "hisseler": stocks if stocks is not None else [],
        "doviz": doviz if doviz is not None else []
    }

    json_dosya = BASE_DIR / "piyasa_verileri.json"
    with open(json_dosya, "w", encoding="utf-8") as f:
        json.dump(json_data, f, ensure_ascii=False, indent=2)

    print(f"✔ '{json_dosya}' kaydedildi.")

    # ========== HIZLI DOSYA (gzip'li kompakt JSON) ==========
    # Uygulama yalnızca bu dosyayı indirip doğrudan jsonDecode eder; XLSX
    # ayrıştırma maliyeti ortadan kalkar. Temel veriler ayrıca BIST xlsx
    # dosyasından alınmaya devam eder.
    def _hisse_kaydi(h):
        kod = str(h.get("Hisse Kodu") or "").strip()
        fiyat = _json_sayi(h.get("Hisse Fiyatı"))
        if not kod or fiyat is None or fiyat <= 0:
            return None
        item = {"k": kod, "p": round(fiyat, 6)}
        ad = str(h.get("Hisse Adı") or "").strip()
        if ad:
            item["n"] = ad
        degisim = _json_sayi(h.get("Günlük Değişim (%)"))
        if degisim is not None:
            item["ch"] = round(degisim, 4)
        return item

    def _fon_kaydi(f):
        kod = str(f.get("Fon Kodu") or "").strip()
        fiyat = _json_sayi(f.get("Fon Fiyatı"))
        if not kod or fiyat is None or fiyat <= 0:
            return None
        item = {"k": kod, "p": round(fiyat, 6)}
        ad = str(f.get("Fon Adı") or "").strip()
        tur = str(f.get("Fon Türü") or "").strip()
        if ad:
            item["n"] = ad
        if tur:
            item["t"] = tur
        degisim = _json_sayi(f.get("Günlük Değişim (%)"))
        if degisim is not None:
            item["ch"] = round(degisim, 4)
        return item

    def _doviz_kaydi(d):
        sembol = str(d.get("Sembol") or "").strip()
        fiyat = _json_sayi(d.get("Fiyat"))
        if not sembol or fiyat is None or fiyat <= 0:
            return None
        item = {"k": sembol, "p": round(fiyat, 6)}
        degisim = _json_sayi(d.get("Günlük Değişim (%)"))
        if degisim is not None:
            item["ch"] = round(degisim, 4)
        return item

    hizli_data = {
        "v": 1,
        "updated": turkiye_saati.strftime("%d.%m.%Y %H:%M"),
        "stocks": [x for h in (stocks or []) if (x := _hisse_kaydi(h))],
        "funds": [x for f in fon_kayitlari if (x := _fon_kaydi(f))],
        "market": [x for d in (doviz or []) if (x := _doviz_kaydi(d))],
    }
    hizli_data["stocks"].sort(key=lambda x: x["k"])

    hizli_dosya = BASE_DIR / "piyasa_verileri.json.gz"
    with gzip.open(hizli_dosya, "wt", encoding="utf-8", compresslevel=9) as f:
        json.dump(hizli_data, f, ensure_ascii=False, separators=(",", ":"))

    print(
        f"✔ '{hizli_dosya}' kaydedildi "
        f"({hizli_dosya.stat().st_size / 1024:.1f} KB, "
        f"{len(hizli_data['stocks'])} hisse, {len(hizli_data['funds'])} fon)."
    )
