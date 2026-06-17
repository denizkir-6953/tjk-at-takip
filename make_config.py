import json, os

cfg = {
    "telegram_bot_token": os.environ.get("TG_TOKEN", ""),
    "telegram_chat_ids": json.loads(os.environ.get("TG_CHATS", "[]")),
    "telegram_aktif": True,
    "email_aktif": True,
    "smtp": {
        "sunucu": "smtp.gmail.com",
        "port": 587,
        "kullanici": os.environ.get("SMTP_USER", ""),
        "sifre": os.environ.get("SMTP_PASS", "")
    },
    "email_alicilar": json.loads(os.environ.get("EMAIL_LIST", "[]")),
    "at_sahipleri": [{"isim": s} for s in json.loads(os.environ.get("AT_LIST", "[]"))],
    "bildirim_saati": "21:00",
    "komisyon_yuzdesi": 5.0,
    "excel_aktif": True,
    "excel_klasor": "raporlar",
    "log_seviyesi": "INFO"
}

with open("config.json", "w", encoding="utf-8") as f:
    json.dump(cfg, f, ensure_ascii=False, indent=2)

print("config.json olusturuldu")
print(open("config.json", encoding="utf-8").read()[:300])
