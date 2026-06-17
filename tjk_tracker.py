"""
TJK At Sahibi Kosu Takip Botu
Kullanim:
  python3 tjk_tracker.py --test
  python3 tjk_tracker.py --tarih 2026-06-15
"""

import requests
import json
import logging
import argparse
import os
from datetime import datetime, date
from typing import Optional

from tjk_scraper import gunluk_sonuclari_cek
from tjk_excel import excel_olustur
from tjk_email import html_rapor_olustur, email_gonder_smtp

CONFIG_FILE = "config.json"


def config_yukle():
    with open(CONFIG_FILE, "r", encoding="utf-8") as f:
        return json.load(f)


def _para(t):
    return f"{t:,.2f}".replace(",","X").replace(".",",").replace("X",".") + " ₺"


def telegram_mesaj_olustur(sahip, tarih, sonuclar, komisyon):
    gun = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
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
            d = str(s.get("Derece","-")).strip()
            ik = _para(s["IkramiyeTL"]) if s["IkramiyeTL"] > 0 else "—"
            satirlar += [
                f"\n{emoji.get(d,'🔵')} *{s['AtAdi']}*",
                f"   📍 {s['HipodromAdi']}  Koşu No: {s['KosuNo']}",
                f"   📏 {s['MesafeAdi']}  |  🌱 {s['ZeminAdi']}",
                f"   👨‍🏫 Antrenör: {s['AntrenorAdi']}",
                f"   🏅 Derece: *{d}*  |  💰 İkramiye: *{ik}*",
            ]
        toplam = sum(s["IkramiyeTL"] for s in sonuclar)
        birinci = sum(1 for s in sonuclar if str(s.get("Derece","")).strip()=="1")
        satirlar += [
            "\n" + "─"*32,
            f"📊 *ÖZET*",
            f"   Toplam at: {len(sonuclar)}  |  Birinci: {birinci}",
            f"   Toplam ikramiye: *{_para(toplam)}*",
            f"   %{komisyon:.0f} komisyon: *{_para(toplam*komisyon/100)}*",
        ]
    return "\n".join(satirlar)


def telegram_gonder(token, chat_ids, mesaj):
    for cid in chat_ids:
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        try:
            r = requests.post(url, json={
                "chat_id": cid, "text": mesaj, "parse_mode": "Markdown"
            }, timeout=10)
            r.raise_for_status()
            logging.info(f"✅ Telegram → {cid}")
        except Exception as e:
            logging.error(f"❌ Telegram hatası ({cid}): {e}")


def gunku_raporu_isle(cfg, tarih=None):
    if tarih is None:
        tarih = date.today().strftime("%Y-%m-%d")

    logging.info(f"📡 Rapor işleniyor: {tarih}")
    komisyon = cfg["komisyon_yuzdesi"]
    sahip_listesi = [s["isim"] for s in cfg["at_sahipleri"]]

    # Veri cek
    tum_sonuclar = gunluk_sonuclari_cek(tarih, sahip_listesi)

    # Excel
    excel_dosyasi = None
    if cfg.get("excel_aktif", True):
        klasor = cfg.get("excel_klasor", "raporlar")
        os.makedirs(klasor, exist_ok=True)
        excel_dosyasi = os.path.join(klasor, f"TJK_Rapor_{tarih.replace('-','')}.xlsx")
        try:
            excel_olustur(tum_sonuclar, tarih, komisyon, excel_dosyasi)
            logging.info(f"📊 Excel: {excel_dosyasi}")
        except Exception as e:
            logging.error(f"❌ Excel hatası: {e}")
            excel_dosyasi = None

    # Email
    if cfg.get("email_aktif", True):
        gun_str = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")
        konu = f"🏇 TJK Koşu Raporu — {gun_str}"
        html = html_rapor_olustur(tum_sonuclar, tarih, komisyon)
        ekler = []
        if excel_dosyasi and os.path.exists(excel_dosyasi):
            ekler.append((excel_dosyasi, f"TJK_Rapor_{tarih.replace('-','')}.xlsx"))
        email_gonder_smtp(cfg["smtp"], cfg["email_alicilar"], konu, html, ekler)

    # Telegram
    if cfg.get("telegram_aktif", True):
        token = cfg["telegram_bot_token"]
        chat_ids = cfg["telegram_chat_ids"]
        for sahip, sonuclar in tum_sonuclar.items():
            mesaj = telegram_mesaj_olustur(sahip, tarih, sonuclar, komisyon)
            print("\n" + mesaj)
            telegram_gonder(token, chat_ids, mesaj)


if __name__ == "__main__":
    parser = argparse.ArgumentParser()
    parser.add_argument("--test", action="store_true")
    parser.add_argument("--tarih", type=str)
    args = parser.parse_args()

    cfg = config_yukle()

    logging.basicConfig(
        level=logging.INFO,
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
        gunku_raporu_isle(cfg)
