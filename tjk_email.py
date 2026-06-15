"""
TJK E-posta Gönderim Modülü
Desteklenen gönderim yöntemleri:
  - SMTP (Gmail, Outlook, özel sunucu)
  - SendGrid API (opsiyonel, daha güvenilir)
"""

import smtplib
import logging
import os
from email.mime.multipart import MIMEMultipart
from email.mime.text      import MIMEText
from email.mime.base      import MIMEBase
from email                import encoders
from datetime             import datetime
from typing               import Optional


# ──────────────────────────────────────────────
#  HTML E-posta Şablonu
# ──────────────────────────────────────────────

def html_rapor_olustur(
    tum_sonuclar: dict,
    tarih: str,
    komisyon_yuzdesi: float
) -> str:
    """Tüm at sahipleri için HTML formatlı e-posta içeriği üretir."""
    gun_str = datetime.strptime(tarih, "%Y-%m-%d").strftime("%d.%m.%Y")

    derece_emoji = {"1": "🥇", "2": "🥈", "3": "🥉"}

    tablolar_html = ""
    for sahip_isim, sonuclar in tum_sonuclar.items():
        toplam_ik  = sum(s["IkramiyeTL"] for s in sonuclar)
        birinci    = sum(1 for s in sonuclar if str(s.get("Derece","")).strip()=="1")
        komisyon   = toplam_ik * komisyon_yuzdesi / 100

        # Sonuç satırları
        satirlar_html = ""
        for s in sonuclar:
            d = str(s.get("Derece", "-")).strip()
            emoji = derece_emoji.get(d, "🔵")

            if   d == "1": row_bg = "#D4EDDA"
            elif d == "2": row_bg = "#FFF3CD"
            elif d == "3": row_bg = "#FFE5D0"
            else:          row_bg = "#FFFFFF"

            ik_str = _para(s.get("IkramiyeTL", 0))

            satirlar_html += f"""
            <tr style="background:{row_bg}">
              <td style="padding:6px 10px;text-align:center">{s.get("KosuNo","-")}</td>
              <td style="padding:6px 10px;font-weight:bold">{emoji} {s.get("AtAdi","-")}</td>
              <td style="padding:6px 10px">{s.get("HipodromAdi","-")}</td>
              <td style="padding:6px 10px;text-align:center">{s.get("MesafeAdi","-")}</td>
              <td style="padding:6px 10px;text-align:center">{s.get("ZeminAdi","-")}</td>
              <td style="padding:6px 10px">{s.get("AntrenorAdi","-")}</td>
              <td style="padding:6px 10px;text-align:center;font-weight:bold">{d}</td>
              <td style="padding:6px 10px;text-align:right;font-weight:bold">{ik_str}</td>
            </tr>"""

        if not satirlar_html:
            satirlar_html = """
            <tr>
              <td colspan="8" style="padding:12px;text-align:center;color:#888">
                Bu tarihte koşuya çıkan at bulunamadı.
              </td>
            </tr>"""

        tablolar_html += f"""
        <div style="margin-bottom:30px">
          <h2 style="color:#1A3A5C;border-bottom:3px solid #2E6DA4;padding-bottom:6px">
            👤 {sahip_isim}
          </h2>
          <table width="100%" cellspacing="0" cellpadding="0"
                 style="border-collapse:collapse;font-size:13px;font-family:Arial,sans-serif">
            <thead>
              <tr style="background:#2E6DA4;color:#fff">
                <th style="padding:8px 10px">Koşu No</th>
                <th style="padding:8px 10px">At Adı</th>
                <th style="padding:8px 10px">Hipodrom</th>
                <th style="padding:8px 10px">Mesafe</th>
                <th style="padding:8px 10px">Zemin</th>
                <th style="padding:8px 10px">Antrenör</th>
                <th style="padding:8px 10px">Derece</th>
                <th style="padding:8px 10px">İkramiye</th>
              </tr>
            </thead>
            <tbody>
              {satirlar_html}
            </tbody>
            <tfoot>
              <tr style="background:#1A3A5C;color:#fff;font-weight:bold">
                <td colspan="6" style="padding:8px 10px;text-align:right">
                  Toplam ({len(sonuclar)} at, {birinci} birinci)
                </td>
                <td colspan="2" style="padding:8px 10px;text-align:right">
                  {_para(toplam_ik)}
                </td>
              </tr>
              <tr style="background:#E8F5E9">
                <td colspan="6" style="padding:8px 10px;text-align:right;color:#1A3A5C;font-weight:bold">
                  %{komisyon_yuzdesi:.0f} Komisyon
                </td>
                <td colspan="2" style="padding:8px 10px;text-align:right;color:#1A3A5C;font-weight:bold">
                  {_para(komisyon)}
                </td>
              </tr>
            </tfoot>
          </table>
        </div>"""

    return f"""<!DOCTYPE html>
<html lang="tr">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width,initial-scale=1">
  <title>TJK Koşu Raporu {gun_str}</title>
</head>
<body style="margin:0;padding:0;background:#F5F5F5;font-family:Arial,sans-serif">
  <table width="100%" cellpadding="0" cellspacing="0"
         style="max-width:800px;margin:20px auto;background:#fff;
                border-radius:8px;box-shadow:0 2px 8px rgba(0,0,0,0.1)">
    <!-- BAŞLIK -->
    <tr>
      <td style="background:#1A3A5C;padding:20px 30px;border-radius:8px 8px 0 0">
        <h1 style="margin:0;color:#fff;font-size:22px">
          🏇 TJK Günlük Koşu Raporu
        </h1>
        <p style="margin:5px 0 0;color:#A9C4E4;font-size:14px">
          📅 {gun_str}
        </p>
      </td>
    </tr>
    <!-- İÇERİK -->
    <tr>
      <td style="padding:25px 30px">
        {tablolar_html}
      </td>
    </tr>
    <!-- ALT BİLGİ -->
    <tr>
      <td style="background:#F5F7FA;padding:15px 30px;
                 border-top:1px solid #E0E0E0;border-radius:0 0 8px 8px;
                 text-align:center;color:#888;font-size:11px">
        Bu rapor otomatik olarak oluşturulmuştur. Kaynak: tjk.org
      </td>
    </tr>
  </table>
</body>
</html>"""


def _para(tutar: float) -> str:
    return f"{tutar:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".") + " ₺"


# ──────────────────────────────────────────────
#  SMTP Gönderici
# ──────────────────────────────────────────────

def email_gonder_smtp(
    smtp_cfg: dict,
    alicilar: list,
    konu: str,
    html_icerik: str,
    ekler: Optional[list] = None   # [(dosya_yolu, dosya_adi), ...]
) -> bool:
    """
    smtp_cfg örneği (Gmail):
    {
      "sunucu":  "smtp.gmail.com",
      "port":    587,
      "kullanici": "sen@gmail.com",
      "sifre":   "UYGULAMA_SIFRESI"
    }

    Gmail için App Password oluşturma:
      myaccount.google.com → Güvenlik → 2 adımlı doğrulama
      → Uygulama şifreleri → "Mail" seç → 16 haneli şifreyi kopyala
    """
    try:
        msg = MIMEMultipart("alternative")
        msg["Subject"] = konu
        msg["From"]    = smtp_cfg["kullanici"]
        msg["To"]      = ", ".join(alicilar)

        # HTML içerik
        msg.attach(MIMEText(html_icerik, "html", "utf-8"))

        # Ekler (Excel dosyası vb.)
        if ekler:
            for dosya_yolu, dosya_adi in ekler:
                if not os.path.exists(dosya_yolu):
                    logging.warning(f"Ek bulunamadı: {dosya_yolu}")
                    continue
                with open(dosya_yolu, "rb") as f:
                    ek = MIMEBase("application", "octet-stream")
                    ek.set_payload(f.read())
                encoders.encode_base64(ek)
                ek.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=dosya_adi
                )
                msg.attach(ek)

        sunucu = smtp_cfg["sunucu"]
        port   = int(smtp_cfg.get("port", 587))

        with smtplib.SMTP(sunucu, port) as server:
            server.ehlo()
            server.starttls()
            server.login(smtp_cfg["kullanici"], smtp_cfg["sifre"])
            server.sendmail(
                smtp_cfg["kullanici"],
                alicilar,
                msg.as_bytes()
            )

        logging.info(f"✅ E-posta gönderildi → {alicilar}")
        return True

    except Exception as e:
        logging.error(f"❌ E-posta gönderilemedi: {e}")
        return False
