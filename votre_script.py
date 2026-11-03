import os
import re
import requests
import openpyxl
from datetime import datetime, timedelta

APIFY_TOKEN = os.environ.get("APIFY_TOKEN")
EXCEL_FILE = "reporting_airbnb.xlsx"  # ← Mets ici le VRAI nom de ton fichier Excel

def get_listing_id(url):
    match = re.search(r'/rooms/(\d+)', str(url))
    return match.group(1) if match else None

def call_apify(listing_id):
    url = f"https://api.apify.com/v2/acts/tri_angle~airbnb-scraper/run-sync-get-dataset-items?token={APIFY_TOKEN}"
    payload = {
        "listingIds": [listing_id],
        "maxReviews": 0,
        "includeReviews": False,
        "calendarMonths": 2
    }
    try:
        response = requests.post(url, json=payload, timeout=120)
        if response.status_code == 200:
            return response.json()
        else:
            print(f"Erreur API {response.status_code}: {response.text}")
            return []
    except Exception as e:
        print(f"Exception: {e}")
        return []

def main():
    if not os.path.exists(EXCEL_FILE):
        print(f"❌ Fichier {EXCEL_FILE} non trouvé dans le repo")
        exit(1)

    wb = openpyxl.load_workbook(EXCEL_FILE)
    print(f"✅ Fichier Excel chargé. Onglets: {wb.sheetnames}")

    # Adapte les noms d'onglets selon ton fichier
    ws_params = wb["paramètre"] if "paramètre" in wb.sheetnames else wb.active
    ws_synth = wb["synthèse"] if "synthèse" in wb.sheetnames else wb.active

    for row in range(2, ws_params.max_row + 1):
        url = ws_params.cell(row=row, column=1).value
        if not url:
            continue

        listing_id = get_listing_id(url)
        print(f"🔍 Traitement listing {listing_id}...")

        data = call_apify(listing_id)
        print(f"   → {len(data)} données reçues")

        # Ici tu rempliras les cellules de l'onglet synthèse

    wb.save(EXCEL_FILE)
    print("✅ Fichier sauvegardé")

if __name__ == "__main__":
    main()
