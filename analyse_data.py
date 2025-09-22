"""
Created on Wed May 28 15:40:15 2025

@author: victor.bontemps
"""
"===================================================================================================================================="
""" On importe les librairies et les inputs du script"""
"===================================================================================================================================="

""" Librairies"""

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
import markdown
import win32com.client as win32
import json


""" Inputs """
stamp = datetime.now(ZoneInfo("Europe/Paris")).strftime("%Y_%m_%d")
out_dir = Path("output")
out_dir.mkdir(parents=True, exist_ok=True)
out_dir_csv =  out_dir / "csv_folder"
out_dir_json =  out_dir / "json_folder"
outfile = out_dir / f"srp_data_output_{stamp}.xlsx"
csv_map = {
    "Interest Rates": out_dir_csv / f"ir_products_{stamp}.csv",
    "Credit": out_dir_csv / f"credit_products_{stamp}.csv",
    "EQD": out_dir_csv / f"eqd_products_{stamp}.csv",
    "Other": out_dir_csv / f"other_products_{stamp}.csv",
}

load_dotenv()
api_key = os.getenv("OPENAI_API_KEY")

PROMPT_SALES = (

    "Tu es expert des produits structurés. Tu reçois un fichier de deals avec les "
    "caractéristiques des produits (nom, émetteur, sous-jacent, volume, maturité, description etc.). "
    "Ta mission : produire une synthèse rapide du fichier pour les vendeurs."
    "Contraintes : Format desk-to-sales (percutant, lisible en diagonale). Moins de 3 minutes de lecture et structuré en blocs clairs avec des bullet points. "
    "Hiérarchise : commence par flux/volumes globaux, sous-jacents dominants, produits et enfin analyse synthétique "
    "Évite les listes trop longues, regroupe par thème/sous-thème "
    "Structure attendue : Flux & volumes (montants, tenor moyen, émetteurs actifs). "
    "Sous-jacents dominants (triés par familles). "
    "Produits & structuration (type de note, wrapper, caractéristiques coupon/call)."
    "Analyse synthétique : Analyse des émissions remarquables et des flux majoritaire en une sule phrase"
    "Dans ton livrable n'ajoute aucune phrases de politesse ou autres ajouts formels, car tu es en charge d'une section d'un mail plus long" 

)

"===================================================================================================================================="
""" On utilise l'API Chat GPT pour analyzer les données"""
"===================================================================================================================================="

client = OpenAI(api_key=api_key)

def summarize_csv(csv_path: Path, category: str, model: str = "gpt-4.1-mini") -> str:
    """
    Lit le CSV puis l'envoie en texte dans le prompt (pas d'upload de fichier).
    On limite la taille pour rester compact.
    """
    df = pd.read_csv(csv_path)

    # Limiter la taille: on garde les 2000 premières lignes puis on tronque à ~100k caractères
    csv_text = df.head(2000).to_csv(index=False)
    csv_text = csv_text[:100000]

    resp = client.responses.create(
        model=model,
        input=(
            f"{PROMPT_SALES}\n\n"
            f"Catégorie: {category}.\n"
            f"Voici les deals au format CSV (échantillon si trop volumineux) :\n\n{csv_text}"
        ),
        temperature=0.2,
    )

    # Extraction simple du texte
    if hasattr(resp, "output_text") and resp.output_text:
        return resp.output_text.strip()

    # Fallback minimal
    try:
        return resp.choices[0].message.content.strip()
    except Exception:
        return "Synthèse indisponible."

summaries = {}
for cat, path in csv_map.items():
    summaries[cat] = summarize_csv(path, cat)

"===================================================================================================================================="
""" On sauvegarde l'analyse de l'API dans un json"""
"===================================================================================================================================="
# Sauvegarde
outpath_json = out_dir_json / f"gpt_srp_analyze_{stamp}.json" # adapte le chemin
with outpath_json.open("w", encoding="utf-8") as f:
    json.dump(summaries, f, ensure_ascii=False, indent=2)

print(f"Saved in: {outpath_json}")

