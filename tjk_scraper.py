"""
TJK Gunluk Yaris Sonuclari - HTML Scraper v4
Gercek CSS class yapisina gore:
  SONUCNO   -> derece
  AtAdi3    -> at adi (parantez oncesi)
  SahipAdi  -> sahip
  AntronorAdi -> antrenor
  race-share dl/dt/dd -> ikramiye
"""

import requests
from bs4 import BeautifulSoup
from datetime import datetime
import re
import logging

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/124.0.0.0 Safari/537.36"
    ),
    "Referer": "https://www.tjk.org/",
    "Accept-Language": "tr-TR,tr;q=0.9",
}

SEHIRLER = [
    {"id": "1",  "ad": "Istanbul"},
    {"id": "2",  "ad": "Izmir"},
    {"id": "3",  "ad": "Ankara"},
    {"id": "4",  "ad": "Bursa"},
    {"id": "5",  "ad": "Adana"},
    {"id": "6",  "ad": "Kocaeli"},
    {"id": "7",  "ad": "Diyarbakir"},
    {"id": "8",  "ad": "Elazig"},
    {"id": "9",  "ad": "Sanliurfa"},
]


def tarih_formatla(tarih_str):
    d = datetime.strptime(tarih_str, "%Y-%m-%d")
    return d.strftime("%d/%m/%Y")


def sehir_html_cek(sehir_id, sehir_ad, tarih_ddmmyyyy):
    url = (
        f"https://www.tjk.org/TR/YarisSever/Info/Sehir/GunlukYarisSonuclari"
        f"?SehirId={sehir_id}"
        f"&QueryParameter_Tarih={tarih_ddmmyyyy.replace('/', '%2F')}"
        f"&SehirAdi={sehir_ad}"
        f"&Era=today"
    )
    try:
        r = requests.get(url, headers=HEADERS, timeout=15)
        if r.status_code == 200 and len(r.text) > 5000:
            return r.text
    except Exception as e:
        logging.warning(f"Sehir cekilemedi ({sehir_ad}): {e}")
    return None


def para_donustur(metin):
    """630.000 -> 630000.0"""
    if not metin:
        return 0.0
    temiz = re.sub(r'[^\d.]', '', str(metin)).strip()
    if not temiz:
        return 0.0
    if re.match(r'^\d{1,3}(\.\d{3})+$', temiz):
        temiz = temiz.replace('.', '')
    try:
        return float(temiz)
    except:
        return 0.0


def race_share_ikramiye(race_share_div):
    """race-share div -> {"1": 630000.0, "2": 252000.0, ...}"""
    ikramiyeler = {}
    if not race_share_div:
        return ikramiyeler
    dl = race_share_div.find('dl')
    if not dl:
        return ikramiyeler
    for dt, dd in zip(dl.find_all('dt'), dl.find_all('dd')):
        sira = re.sub(r'[^\d]', '', dt.get_text(strip=True))
        span = dd.find('span', class_='tlsymbol')
        if span:
            span.decompose()
        tutar = para_donustur(dd.get_text(strip=True))
        if sira and tutar > 0:
            ikramiyeler[sira] = tutar
    return ikramiyeler


def at_adi_temizle(metin):
    """'ARGİTHANLI ŞENGÜL(2)KG SK...' -> 'ARGİTHANLI ŞENGÜL'"""
    if not metin:
        return "-"
    # Parantez ve sonrasini kes
    metin = re.sub(r'\(.*', '', metin).strip()
    # Kose parantez ve sonrasini kes
    metin = re.sub(r'\[.*', '', metin).strip()
    return metin.strip() or "-"


def html_parse_et(html, sehir_ad, hedef_sahipler):
    soup = BeautifulSoup(html, 'html.parser')
    sonuclar = []
    hedef_upper = {s.upper().strip(): s for s in hedef_sahipler}

    race_shares = soup.find_all('div', class_='race-share')

    for race_share in race_shares:
        ikramiyeler = race_share_ikramiye(race_share)

        # Kosu bilgisi
        kosuno = "-"
        mesafe = "-"
        zemin = "-"
        onceki = race_share.find_previous(['h2', 'h3'])
        if onceki:
            baslik = onceki.get_text()
            m = re.search(r'(\d+)\.\s*[Kk]o[sş]u', baslik)
            if m:
                kosuno = m.group(1)
            m2 = re.search(r'(\d{3,4})\s*[Mm]', baslik)
            if m2:
                mesafe = m2.group(0).strip()
            if 'Çim' in baslik:
                zemin = 'Çim'
            elif 'Kum' in baslik:
                zemin = 'Kum'

        sonraki_tablo = race_share.find_next('table')
        if not sonraki_tablo:
            continue

        for satir in sonraki_tablo.find_all('tr'):
            # Sahip
            sahip_td = satir.find('td', class_=re.compile('SahipAdi', re.I))
            if not sahip_td:
                continue

            sahip_a = sahip_td.find('a')
            sahip_metin = (
                sahip_a.get_text(strip=True) if sahip_a
                else sahip_td.get_text(strip=True)
            ).upper().strip()

            eslesen = None
            for hedef, orijinal in hedef_upper.items():
                if hedef in sahip_metin or sahip_metin in hedef:
                    eslesen = orijinal
                    break
            if not eslesen:
                continue

            # At adi - AtAdi3 class'i
            at_td = satir.find('td', class_=re.compile('AtAdi', re.I))
            if at_td:
                at_adi = at_adi_temizle(at_td.get_text(strip=True))
            else:
                at_adi = "-"

            # Derece - SONUCNO class'i
            sonuc_td = satir.find('td', class_=re.compile('SONUCNO|SonucNo', re.I))
            if sonuc_td:
                derece = sonuc_td.get_text(strip=True)
            else:
                derece = "-"

            # Antrenor
            ant_td = satir.find('td', class_=re.compile('Antron|Antren', re.I))
            if ant_td:
                ant_a = ant_td.find('a')
                antrenor = (
                    ant_a.get_text(strip=True) if ant_a
                    else ant_td.get_text(strip=True)
                )
            else:
                antrenor = "-"

            # Ikramiye
            ikramiye = ikramiyeler.get(derece.strip(), 0.0)

            sonuclar.append({
                "AtAdi":       at_adi,
                "KosuNo":      kosuno,
                "HipodromAdi": sehir_ad,
                "Derece":      derece.strip(),
                "IkramiyeTL":  ikramiye,
                "AntrenorAdi": antrenor,
                "AtSahibiAdi": eslesen,
                "MesafeAdi":   mesafe,
                "ZeminAdi":    zemin,
                "GrupAdi":     "-",
            })

    return sonuclar


def gunluk_sonuclari_cek(tarih_str, hedef_sahipler):
    tarih_ddmmyyyy = tarih_formatla(tarih_str)
    tum_sonuclar = {s: [] for s in hedef_sahipler}

    for sehir in SEHIRLER:
        logging.info(f"Taranıyor: {sehir['ad']}")
        html = sehir_html_cek(sehir["id"], sehir["ad"], tarih_ddmmyyyy)
        if not html:
            logging.info(f"  -> {sehir['ad']} bu tarihte yok")
            continue

        logging.info(f"  -> {sehir['ad']} HTML alindi ({len(html):,} byte)")
        bulunanlar = html_parse_et(html, sehir["ad"], hedef_sahipler)

        for sonuc in bulunanlar:
            eslesen = sonuc["AtSahibiAdi"]
            if eslesen in tum_sonuclar:
                tum_sonuclar[eslesen].append(sonuc)
                logging.info(
                    f"     ✅ {eslesen}: {sonuc['AtAdi']} "
                    f"- {sonuc['Derece']}. derece "
                    f"- {sonuc['IkramiyeTL']:,.0f} TL"
                )

    return tum_sonuclar
