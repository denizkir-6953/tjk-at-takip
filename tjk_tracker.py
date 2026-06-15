"""
TJK At Sahibi Koşu Sonuçları Takip Botu
========================================
Belirtilen at sahiplerinin günlük koşu sonuçlarını TJK.org'dan çeker,
Excel raporu oluşturur, e-posta atar ve Telegram'a bildirim gönderir.

Kurulum:
  pip install requests schedule openpyxl

Kullanım:
  python tjk_tracker.py              # Sürekli çalışır (zamanlanmış)
  python tjk_tracker.py --test       # Bugünü hemen çek ve gönder
  python tjk_tracker.py --tarih 2024-06-15
"""

import requests
import json
import schedule
import time
import logging
import argparse
import os
import tempfile
from datetime import datetime, date
from typing import Optional

from tjk_excel import excel_olustur
from tjk_email import html_rapor_olustur, email_gonder_smtp

# ─────────────────────────────────────────
#  AYARLAR
# ─────────────────────────────────────────

CONFIG_FILE = "config.json"

DEFAULT_CONFIG = {
    # ── Telegram ─────────────────────────
    "telegram_bot_token":  "BOT_TOKEN_BURAYA",
    "telegram_chat_ids":   ["CHAT_ID_BURAYA"],
    "telegram_aktif":      True,

    # ── E-posta (SMTP) ───────────────────
    "email_aktif":         True,
    "smtp": {
        "sunucu":    "smtp.gmail.com",
        "port":      587,
        "kullanici": "senin@gmail.com",
        "sifre":     "GMAIL_UYGULAMA_SIFRESI"
    },
    "email_alicilar": [
        "alici1@example.com",
        "alici2@example.com"
    ],

    # ── Excel ────────────────────────────
    "excel_aktif":         True,
    "excel_klasor":        "raporlar",   # Oluşturulan xlsx'ler burada saklanır

    # ── At Sahipleri ─────────────────────
    "at_sahipleri": [
        {"isim": "AT SAHİBİ 1 TAM İSİM", "id": ""},
        {"isim": "AT SAHİBİ 2 TAM İSİM", "id": ""}
    ],

    # ── Genel ────────────────────────────
    "bildirim_saati":      "21:00",
    "komisyon_yuzdesi":    5.0,
    "log_seviyesi":        "INFO"
}


def config_yukle() -> dict:
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, "r", encoding="utf-8") as f:
            cfg = json.load(f)
        for k, v in DEFAULT_CONFIG.items():
            cfg.setdefault(k, v)
        return cfg
    with open(CONFIG_FILE, "w", encoding="utf-8") as f:
        json.dump(DEFAULT_CONFIG, f, ensure_ascii=False, indent=2)
    print(f"⚠️  config.json oluşturuldu. Lütfen düzenleyin: {CONFIG_FILE}")
    return DEFAULT_CONFIG


# ─────────────────────────────────────────
#  TJK API
# ─────────────────────────────────────────

TJK_BASE = "https://www.tjk.org"
SESSION  = requests.Session()
SESSION.headers.update({
    "User-Agent":      "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36",
    "Accept":          "application/json, text/plain, */*",
    "Accept-Language": "tr-TR,tr;q=0.9",
    "Referer":         "https://www.tjk.org/",
    "X-Requested-With":"XMLHttpRequest",
})


def tjk_at_sahibi_id_bul(isim: str) -> Optional[str]:
    url = f"{TJK_BASE}/TR/YarisSever/Query/GetAtSahibiList"
    payload = {
        "sMode": "GetAtSahibiList",
        "jsonString": json.dumps({
            "AtSahibiAdi": isim,
            "QueryParameter": {"Value": "1", "DataType": "Number"}
        })
    }
    try:
        r = SESSION.post(url, data=payload, timeout=15)
        r.raise_for_status()
        data  = r.json()
        liste = data.get("Value") or data.get("value") or []
        if liste:
            return str(liste[0].get("AtSahibiId") or liste[0].get("Id", ""))
    except Exception as e:
        logging.warning(f"At sahibi ID araması başarısız ({isim}): {e}")
    return None


def tjk_kos_sonuclari_cek(at_sahibi_id: str, tarih: str) -> list:
    url = f"{TJK_BASE}/TR/YarisSever/Query/GetAtSahipKosuSonuclari"
    payload = {
        "sMode": "GetAtSahipKosuSonuclari",
        "jsonString": json.dumps({
            "AtSahibiId": at_sahibi_id,
            "QueryParameter": {"Value": tarih, "DataType": "DateTime"}
        })
    }
    try:
        r    = SESSION.post(url, data=payload, timeout=20)
        r.raise_for_status()
        data = r.json()
        return _normalize(data.get("Value") or data.get("value") or [])
    except Exception as e:
        logging.error(f"TJK API hatası (id={at_sahibi_id}, tarih={tarih}): {e}")
        return []


def _normalize(liste: list) -> list:
    mapping = {
        "AtAdi":       ["AtAdi","ATADI","AtAdı","HorseNameTR"],
        "KosuNo":      ["KosuNo","KosuSira","RaceNo"],
        "HipodromAdi": ["HipodromAdi","HipodromAd","Track","Hipodrom"],
        "Derece":      ["Derece","Sonuc","Place"],
        "IkramiyeTL":  ["IkramiyeTL","Ikramiye","Prize","ToplamIkramiye"],
        "AntrenorAdi": ["AntrenorAdi","AntrenorAd","Trainer","AntrenörAdı"],
        "AtSahibiAdi": ["AtSahibiAdi","AtSahibi","Owner"],
        "MesafeAdi":   ["MesafeAdi","Mesafe","Distance"],
        "ZeminAdi":    ["ZeminAdi","Zemin","Surface"],
        "GrupAdi":     ["GrupAdi","Grup","Group","KosuGrubu"],
    }
    result = []
    for item in liste:
        row = {}
        for hedef, kaynaklar in mapping.items():
            for k in kaynaklar:
                if k in item and item[k] not in (None, "", "-"):
                    row[hedef] = item[k]
                    break
            row.setdefault(hedef, "-")
        try:
            row["IkramiyeTL"] = float(
                str(row["IkramiyeTL"])
                .replace(".", "").replace(",", ".").replace("₺","").strip()
            )
        except (ValueError, TypeError):
            row["IkramiyeTL"] = 0.0
        result.append(row)
    return result


# ─────────────────────────────────────────
#  TELEGRAM
# ─────────────────────────────────────────

def _para(t: float) -> str:
    return f"{t:,.2f}".replace(",","X").replace(".",",").replace("X",".") + " ₺"


def telegram_mesaj_olustur(sahip: str, tarih: str,
                            sonuclar: list, komisyon: float) -> str:
    gun   = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
    emoji = {"1":"🥇","2":"🥈","3":"🥉"}
    satirlar = [
        f"🏇 *TJK Koşu Sonuçları*",
        f"📅 {gun}  |  👤 *{sahip}*",
        "─"*32,
    ]
    if not sonuclar:
        satirlar.append("❌ Bu tarihte koşuya çıkan at bulunamadı.")
    else:
        for s in sonuclar:
            d   = str(s.get("Derece","-")).strip()
            ik  = _para(s["IkramiyeTL"]) if s["IkramiyeTL"] > 0 else "—"
            satirlar += [
                f"\n{emoji.get(d,'🔵')} *{s['AtAdi']}*",
                f"   📍 {s['HipodromAdi']}  Koşu No: {s['KosuNo']}",
                f"   📏 {s['MesafeAdi']}  |  🌱 {s['ZeminAdi']}",
                f"   👨‍🏫 Antrenör: {s['AntrenorAdi']}",
                f"   🏅 Derece: *{d}*  |  💰 İkramiye: *{ik}*",
            ]
        toplam = sum(s["IkramiyeTL"] for s in sonuclar)
        birinci= sum(1 for s in sonuclar if str(s.get("Derece","")).strip()=="1")
        saturlar_ek = [
            "\n" + "─"*32,
            f"📊 *ÖZET*",
            f"   Toplam at: {len(sonuclar)}  |  Birinci: {birinci}",
            f"   Toplam ikramiye: *{_para(toplam)}*",
            f"   %{komisyon:.0f} komisyon: *{_para(toplam*komisyon/100)}*",
        ]
        satirlar.extend(saturlar_ek)
    return "\n".join(satirlar)


def telegram_gonder(token: str, chat_ids: list, mesaj: str) -> bool:
    ok = True
    for cid in chat_ids:
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        try:
            r = requests.post(url, json={
                "chat_id": cid, "text": mesaj, "parse_mode": "Markdown"
            }, timeout=10)
            r.raise_for_status()
            logging.info(f"✅ Telegram → chat_id={cid}")
        except Exception as e:
            logging.error(f"❌ Telegram hatası ({cid}): {e}")
            ok = False
    return ok


# ─────────────────────────────────────────
#  ANA İŞ AKIŞI
# ─────────────────────────────────────────

def gunku_raporu_isle(cfg: dict, tarih: Optional[str] = None):
    if tarih is None:
        tarih = date.today().strftime("%Y-%m-%d")
    logging.info(f"📡 Rapor işleniyor: {tarih}")

    komisyon = cfg["komisyon_yuzdesi"]

    # ── 1. Tüm sahipler için veri çek ─────────────────────
    tum_sonuclar = {}   # { isim: [sonuclar] }

    for sahip in cfg["at_sahipleri"]:
        isim = sahip["isim"]
        sid  = sahip.get("id", "").strip()

        if not sid:
            logging.info(f"🔍 ID aranıyor: {isim}")
            sid = tjk_at_sahibi_id_bul(isim) or ""
            if sid:
                sahip["id"] = sid

        if not sid:
            logging.warning(f"⚠️  ID bulunamadı, atlanıyor: {isim}")
            continue

        sonuclar = tjk_kos_sonuclari_cek(sid, tarih)
        tum_sonuclar[isim] = sonuclar
        logging.info(f"   {isim}: {len(sonuclar)} sonuç")

    if not tum_sonuclar:
        logging.warning("Hiç veri bulunamadı, bildirim gönderilmiyor.")
        return

    # ── 2. Excel oluştur ───────────────────────────────────
    excel_dosyasi = None
    if cfg.get("excel_aktif", True):
        klasor = cfg.get("excel_klasor", "raporlar")
        os.makedirs(klasor, exist_ok=True)
        excel_dosyasi = os.path.join(
            klasor, f"TJK_Rapor_{tarih.replace('-','')}.xlsx"
        )
        try:
            excel_olustur(tum_sonuclar, tarih, komisyon, excel_dosyasi)
            logging.info(f"📊 Excel oluşturuldu: {excel_dosyasi}")
        except Exception as e:
            logging.error(f"❌ Excel oluşturulamadı: {e}")
            excel_dosyasi = None

    # ── 3. E-posta gönder ──────────────────────────────────
    if cfg.get("email_aktif", True):
        gun_str = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
        konu    = f"🏇 TJK Koşu Raporu — {gun_str}"
        html    = html_rapor_olustur(tum_sonuclar, tarih, komisyon)

        ekler = []
        if excel_dosyasi and os.path.exists(excel_dosyasi):
            ekler.append((
                excel_dosyasi,
                f"TJK_Rapor_{tarih.replace('-','')}.xlsx"
            ))

        email_gonder_smtp(
            smtp_cfg  = cfg["smtp"],
            alicilar  = cfg["email_alicilar"],
            konu      = konu,
            html_icerik = html,
            ekler     = ekler
        )

    # ── 4. Telegram gönder ─────────────────────────────────
    if cfg.get("telegram_aktif", True):
        token    = cfg["telegram_bot_token"]
        chat_ids = cfg["telegram_chat_ids"]
        for isim, sonuclar in tum_sonuclar.items():
            mesaj = telegram_mesaj_olustur(isim, tarih, sonuclar, komisyon)
            print("\n" + mesaj)
            telegram_gonder(token, chat_ids, mesaj)


def zamanlanmis_calistir(cfg: dict):
    saat = cfg.get("bildirim_saati", "21:00")
    schedule.every().day.at(saat).do(gunku_raporu_isle, cfg=cfg)
    logging.info(f"⏰ Zamanlayıcı: her gün {saat}")
    while True:
        schedule.run_pending()
        time.sleep(30)


# ─────────────────────────────────────────
#  GİRİŞ NOKTASI
# ─────────────────────────────────────────

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="TJK At Sahibi Koşu Takip Botu")
    parser.add_argument("--test",  action="store_true")
    parser.add_argument("--tarih", type=str)
    args = parser.parse_args()

    cfg = config_yukle()

    log_level = getattr(logging, cfg.get("log_seviyesi","INFO").upper(), logging.INFO)
    logging.basicConfig(
        level=log_level,
        format="%(asctime)s  %(levelname)-8s  %(message)s",
        datefmt="%Y-%m-%d %H:%M:%S",
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler("tjk_tracker.log", encoding="utf-8"),
        ]
    )

    if args.test or args.tarih:
        gunku_raporu_isle(cfg, tarih=args.tarih)
    else:
        zamanlanmis_calistir(cfg)
