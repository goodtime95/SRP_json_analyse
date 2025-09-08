
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

""" Inputs """
stamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y_%m_%d")
out_dir = Path("output")
out_dir.mkdir(parents=True, exist_ok=True)
out_dir_json =  out_dir / "json_folder"
outpath_json = out_dir_json / f"gpt_srp_analyze_{stamp}.json" # adapte le chemin
to_email = os.getenv("EMAIL_TO")
subject = "Overview SRP data"

"===================================================================================================================================="
""" On ecrit le mail"""
"===================================================================================================================================="


# 1) Lire le JSON et construire un texte simple
with open(outpath_json, "r", encoding="utf-8") as f:
    data = json.load(f)

lines = []
for section, content in (data.items() if isinstance(data, dict) else enumerate(data)):
    lines.append(f"## {section}\n{content}\n")
body = "\n".join(lines)

# 2) Créer et envoyer le mail via Outlook
outlook = win32.Dispatch("Outlook.Application")
mail = outlook.CreateItem(0)  # 0 = MailItem
mail.To = to_email
mail.Subject = subject
mail.Body = body  # texte brut (simple et robuste)
mail.Display()

