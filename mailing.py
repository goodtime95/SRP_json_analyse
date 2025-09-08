
"""
Created on Wed May 28 15:40:15 2025

@author: victor.bontemps
"""
"===================================================================================================================================="
""" On importe les librairies et les inputs du script"""
"===================================================================================================================================="

""" Librairies"""
import json
import win32com.client as win32
from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import os
from dotenv import load_dotenv
import markdown

""" Inputs """
stamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y_%m_%d")
out_dir = Path("output")
out_dir.mkdir(parents=True, exist_ok=True)
out_dir_json =  out_dir / "json_folder"
outpath_json = out_dir_json / f"gpt_srp_analyze_{stamp}.json" # adapte le chemin
load_dotenv()
to_email = os.getenv("EMAIL_TO")
subject = "Overview SRP data"

"===================================================================================================================================="
""" On ecrit le mail"""
"===================================================================================================================================="

# 1) Lire le JSON et construire un texte simple
with open(outpath_json, "r", encoding="utf-8") as f:
    data = json.load(f)

"""
Mail SRP avec rendu HTML Outlook, spacing affiné et message de dispo en fin
"""
import os
import json
from pathlib import Path
from datetime import datetime
from zoneinfo import ZoneInfo

import win32com.client as win32
from dotenv import load_dotenv
import markdown  # pip install markdown

# ---------- Inputs ----------
stamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y_%m_%d")
out_dir = Path("output")
out_dir_json = out_dir / "json_folder"
outpath_json = out_dir_json / f"gpt_srp_analyze_{stamp}.json"
outfile = out_dir / f"srp_data_output_{stamp}.xlsx"

load_dotenv()
from_email = os.getenv("EMAIL_FROM")
to_email = os.getenv("EMAIL_TO")
subject = "Overview SRP data"

# ---------- Lecture JSON ----------
with open(outpath_json, "r", encoding="utf-8") as f:
    data = json.load(f)

# ---------- Helpers ----------
def md_to_html(md_text: str) -> str:
    """
    Convertit le Markdown -> HTML, puis affine les marges des listes/paragraphes
    pour un rendu propre dans Outlook (styles inline).
    """
    md_text = md_text.replace("\r\n", "\n").strip()

    html = markdown.markdown(
        md_text,
        extensions=["extra", "sane_lists", "nl2br"]
    )

    # Normaliser les séparateurs
    html = html.replace(
        "<hr />",
        '<hr style="border:none;border-top:1px solid #E1E1E1;margin:14px 0;">'
    )

    # Affiner les marges des listes/éléments (Outlook-friendly)
    html = (
        html.replace("<ul>",  '<ul style="margin:6px 0 0 20px;padding:0;">')
            .replace("<ol>",  '<ol style="margin:6px 0 0 20px;padding:0;">')
            .replace("<li>",  '<li style="margin:4px 0 4px 2px;">')
            .replace("<p>",   '<p style="margin:6px 0;">')
            .replace("<h3>",  '<h3 style="margin:10px 0 6px 0;font-size:14pt;line-height:1.25;color:#1F3A5F;">')
            .replace("<h4>",  '<h4 style="margin:8px 0 4px 0;font-size:12.5pt;line-height:1.25;color:#1F3A5F;">')
    )

    return html

def section_block(title: str, content_md: str) -> str:
    content_html = md_to_html(content_md)
    # Table-based layout: rendus Outlook robustes
    return f"""
    <!-- Section -->
    <tr>
      <td style="padding:0;">
        <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
               style="border-collapse:collapse;background:#FFFFFF;border:1px solid #EEE;border-radius:12px;">
          <tr>
            <td style="padding:14px 16px 4px 16px;">
              <div style="font-family:Calibri,Arial,sans-serif;font-size:18px;line-height:1.25;
                          color:#0B1F33;font-weight:700;">
                {title}
              </div>
            </td>
          </tr>
          <tr>
            <td style="padding:2px 16px 12px 16px;">
              <div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.45;color:#202124;">
                {content_html}
              </div>
            </td>
          </tr>
        </table>
      </td>
    </tr>
    <!-- Fine séparation minimale entre sections -->
    <tr><td style="height:10px;line-height:10px;">&nbsp;</td></tr>
    """

# ---------- Construire le corps HTML ----------
sections_html = []
iterable = data.items() if isinstance(data, dict) else enumerate(data)
for section, content in iterable:
    sections_html.append(section_block(str(section), str(content)))

# Bloc de dispo final UNIQUEMENT à la fin
closing_note = f"""
<tr>
  <td style="padding:8px 0 0 0;">
    <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0"
           style="border-collapse:collapse;background:#FFFFFF;border:1px solid #EEE;border-radius:12px;">
      <tr>
        <td style="padding:14px 16px;">
          <div style="font-family:Calibri,Arial,sans-serif;font-size:11pt;line-height:1.5;color:#202124;">
            Je reste disponible pour tout renseignement complémentaire ou pour approfondir une section en particulier.
          </div>
        </td>
      </tr>
    </table>
  </td>
</tr>
"""

html_body = f"""
<html>
  <body style="margin:0;padding:0;background:#F6F7F9;">
    <center style="width:100%;background:#F6F7F9;">
      <table role="presentation" width="100%" cellpadding="0" cellspacing="0" border="0">
        <tr><td align="center" style="padding:18px 12px;">
          <table role="presentation" cellpadding="0" cellspacing="0" border="0" width="760"
                 style="width:760px;max-width:760px;background:#F6F7F9;">
            <!-- Header compact -->
            <tr>
              <td style="padding:0 0 10px 0;">
                <div style="font-family:Calibri,Arial,sans-serif;font-size:20px;line-height:1.25;
                            color:#0B1F33;font-weight:700;">
                  Overview SRP – {stamp}
                </div>
              </td>
            </tr>

            {''.join(sections_html)}

            {closing_note}

            <!-- Footer -->
            <tr>
              <td style="padding:10px 0 0 0;">
                <div style="font-family:Calibri,Arial,sans-serif;font-size:10pt;color:#8A8F98;">
                  Généré par l'équipe d'ingénierie X-Asset avec l'aide de l'IA • Eleva Solutions
                </div>
              </td>
            </tr>
          </table>
        </td></tr>
      </table>
    </center>
  </body>
</html>
"""
# Créer et envoyer le mail
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # MailItem
mail.To = to_email
mail.Subject = subject
mail.HTMLBody = html_body  # ⚡ HTML au lieu de texte brut
mail.Attachments.Add(Source=str(outfile.resolve()))
mail.Save()  # ou mail.Send()


