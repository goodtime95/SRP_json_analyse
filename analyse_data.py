from datetime import datetime
from zoneinfo import ZoneInfo
from pathlib import Path
import os
import pandas as pd
from openai import OpenAI
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from dotenv import load_dotenv
import sys
import re
import html
import win32com.client as win32
import markdown


# --- Timestamp et output dir ---
stamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y_%m_%d")
out_dir = Path("output")
out_dir.mkdir(parents=True, exist_ok=True)

outfile = out_dir / f"srp_data_output_{stamp}.xlsx"

csv_map = {
    "Interest Rates": out_dir / f"ir_products_{stamp}.csv",
    "Credit": out_dir / f"credit_products_{stamp}.csv",
    "EQD": out_dir / f"eqd_products_{stamp}.csv",
    "Other": out_dir / f"other_products_{stamp}.csv",
}

# Charger automatiquement le .env
load_dotenv()

# Récupérer les variables
api_key = os.getenv("OPENAI_API_KEY")
from_email = os.getenv("EMAIL_FROM")
to_email = os.getenv("EMAIL_TO")

PROMPT_SALES = (

    "Tu es analyste sur un desk de produits structurés. Tu reçois un fichier de deals avec les "
    "caractéristiques des produits (nom, émetteur, sous-jacent, volume, maturité, description etc.). "
    "Ta mission : produire une synthèse rapide et actionnable pour les vendeurs, à inclure dans un mail avec le fichier. "
    "Contraintes : Format desk-to-sales (percutant, lisible en diagonale). Moins de 10 minutes de lecture → structuré en blocs clairs avec bullet points. "
    "Hiérarchise : commence par flux/volumes globaux, sous-jacents dominants, produits et enfin analyse synthétique "
    "Évite les listes trop longues → regroupe par thème/sous-thème "
    "Structure attendue : Flux & volumes (tickets, montants, tenor moyen, taille typique des deals, émetteurs actifs). "
    "Sous-jacents dominants (triés par familles). "
    "Produits & structuration (type de note, wrapper, caractéristiques coupon/call)."
    "Analyse synthétique : Analyse des émissions remarquables et des flux majoritaire en une phrase"

)

# client = OpenAI(api_key=api_key)

# def summarize_csv(csv_path: Path, category: str, model: str = "gpt-4.1-mini") -> str:
#     """
#     Lit le CSV puis l'envoie en texte dans le prompt (pas d'upload de fichier).
#     On limite la taille pour rester compact.
#     """
#     df = pd.read_csv(csv_path)

#     # Limiter la taille: on garde les 2000 premières lignes puis on tronque à ~100k caractères
#     csv_text = df.head(2000).to_csv(index=False)
#     csv_text = csv_text[:100000]

#     resp = client.responses.create(
#         model=model,
#         input=(
#             f"{PROMPT_SALES}\n\n"
#             f"Catégorie: {category}.\n"
#             f"Voici les deals au format CSV (échantillon si trop volumineux) :\n\n{csv_text}"
#         ),
#         temperature=0.2,
#     )

#     # Extraction simple du texte
#     if hasattr(resp, "output_text") and resp.output_text:
#         return resp.output_text.strip()

#     # Fallback minimal
#     try:
#         return resp.choices[0].message.content.strip()
#     except Exception:
#         return "Synthèse indisponible."

# summaries = {}
# for cat, path in csv_map.items():
#     summaries[cat] = summarize_csv(path, cat)

# print(summaries)

def md_to_html(md_text: str) -> str:
    """
    Convertit du Markdown en HTML avec des extensions adaptées
    (listes "propres", tableaux, sauts de ligne, etc.).
    """
    return markdown.markdown(
        md_text or "",
        extensions=["extra", "sane_lists", "nl2br", "tables"]
    )
def summaries_to_html_email(summaries: dict, title: str = "Synthèses Desk-to-Sales") -> str:
    """
    Crée un HTML compatible Outlook :
    - Bandeau titre
    - Sommaire cliquable (ancres)
    - Un bloc par clé (ex: 'Interest Rates', 'Credit', ...)
    """
    header_html = f"""
<div style="font-family:Segoe UI, Arial, sans-serif; font-size:13px; color:#222;">
  <div style="background:#0b5cff; color:#fff; padding:16px; border-radius:6px 6px 0 0;">
    <h2 style="margin:0; font-size:18px;">{html.escape(title)}</h2>
  </div>
  <div style="border:1px solid #e5e5e5; border-top:none; border-radius:0 0 6px 6px; padding:16px;">
"""

    # Sommaire
    keys = list(summaries.keys())
    toc_items = []
    for k in keys:
        anchor = re.sub(r'[^a-z0-9]+', '-', k.lower()).strip('-')
        toc_items.append(
            f'<li style="margin:4px 0;"><a href="#{anchor}" style="color:#0b5cff; text-decoration:none;">{html.escape(k)}</a></li>'
        )
    toc_html = f"""
    <div style="margin-bottom:12px;">
      <p style="margin:6px 0 4px 0; color:#555;"><strong>Sommaire</strong></p>
      <ul style="margin:4px 0 12px 18px; padding:0;">
        {''.join(toc_items)}
      </ul>
    </div>
    <hr style="border:none;border-top:1px solid #eee;margin:12px 0 16px;">
"""

    # Sections
    sections_html = []
    for k in keys:
        anchor = re.sub(r'[^a-z0-9]+', '-', k.lower()).strip('-')
        content_html = md_to_html(summaries.get(k, ""))
        section_block = f"""
    <a id="{anchor}"></a>
    <h3 style="margin:18px 0 8px; font-size:16px;">{html.escape(k)}</h3>
    <div style="margin:0 0 12px;">
      {content_html}
    </div>
"""
        sections_html.append(section_block)

    footer_html = """
    <div style="margin-top:16px; font-size:12px; color:#666;">
      — Généré automatiquement.
    </div>
  </div>
</div>
"""
    return header_html + toc_html + "\n".join(sections_html) + footer_html

# --- 3) Envoi / sauvegarde via Outlook ---
def create_mail_from_summaries(
    summaries: dict,
    subject: str,
    recipient: str,
    title: str = "Synthèses Desk-to-Sales",
    send: bool = False,
    cc: str = "",
    bcc: str = ""
):
    """
    Construit le HTML depuis 'summaries' puis crée le mail Outlook.
    """
    html_body = summaries_to_html_email(summaries, title=title)

    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    mail.To = recipient
    if cc:
        mail.CC = cc
    if bcc:
        mail.BCC = bcc
    mail.Subject = subject
    mail.HTMLBody = html_body  # HTMLBody (pas besoin de fallback)
    if send:
        mail.Send()
    else:
        mail.Save()


create_mail_from_summaries(
        summaries=summaries,
        subject="Desk-to-Sales – Synthèses & Recos",
        recipient="mail.address@gmail.com",
        title="Desk-to-Sales – Synthèses (IR, Credit, EQD, Other)",
        send=False  # True pour envoyer
    )
