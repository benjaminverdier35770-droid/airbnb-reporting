import os
import re
import requests
from datetime import datetime, timedelta
from openpyxl import load_workbook
from collections import defaultdict

# Configuration
APIFY_TOKEN = os.environ['APIFY_TOKEN']
EXCEL_FILE = 'reporting_airbnb.xlsx'  # nom de ton fichier Excel dans le repo

# Charger le fichier Excel
print(f"📂 Chargement du fichier Excel...")
wb = load_workbook(EXCEL_FILE)
print(f"✅ Onglets trouvés : {wb.sheetnames}")

# Récupérer les IDs Airbnb depuis l'onglet "paramètre" (D2 à D7)
ws_param = wb['paramètre']
listing_ids = []
for row in range(2, 8):
    cell_value = ws_param.cell(row=row, column=4).value  # colonne D
    if cell_value and 'airbnb' in str(cell_value).lower():
        match = re.search(r'/rooms/(\d+)', str(cell_value))
        if match:
            listing_ids.append({
                'id': match.group(1),
                'row': row,
                'name': ws_param.cell(row=row, column=1).value or f"Listing {match.group(1)}"
            })
            print(f"✅ Listing trouvé : {match.group(1)} (ligne {row})")

print(f"📊 Total listings : {len(listing_ids)}")

# Fonction pour appeler l'API Apify
def get_airbnb_data(listing_id):
    url = f"https://api.apify.com/v2/acts/dtrungtin~airbnb-scraper/run-sync-get-dataset-items?token={APIFY_TOKEN}"
    payload = {
        "startUrls": [{"url": f"https://www.airbnb.fr/rooms/{listing_id}"}],
        "maxListings": 1,
        "includeReviews": False,
        "includeCalendar": True
    }

    try:
        response = requests.post(url, json=payload, timeout=60)
        if response.status_code == 200:
            data = response.json()
            if data and len(data) > 0:
                return data[0]
        print(f"❌ Erreur API {response.status_code} pour {listing_id}")
        return None
    except Exception as e:
        print(f"❌ Exception pour {listing_id}: {e}")
        return None

# Traiter chaque listing
ws_synthese = wb['synthèse']
today = datetime.now().date()

for listing in listing_ids:
    lid = listing['id']
    print(f"\n🏠 Traitement listing {lid}...")

    data = get_airbnb_data(lid)
    if not data:
        print(f"   ⚠️ Pas de données pour {lid}")
        continue

    # Extraire le calendrier (disponibilité et prix)
    calendar = data.get('calendar', []) or data.get('calendarMonths', [])

    # Calculer les taux de remplissage
    def get_occupancy(days_ahead):
        end_date = today + timedelta(days=days_ahead)
        booked = 0
        total = 0
        for day in calendar:
            day_date = datetime.strptime(day.get('date', ''), '%Y-%m-%d').date() if day.get('date') else None
            if day_date and today <= day_date <= end_date:
                total += 1
                if not day.get('available', True):
                    booked += 1
        return round((booked / total * 100), 1) if total > 0 else 0

    # Remplir les cellules (à adapter selon ta structure exacte)
    # Hypothèse : ligne = position du listing, colonnes = métriques
    row_idx = listing['row']  # ou calcule la ligne dans synthèse

    # Taux de remplissage
    ws_synthese.cell(row=row_idx, column=2, value=get_occupancy(0))   # Aujourd'hui
    ws_synthese.cell(row=row_idx, column=3, value=get_occupancy(1))   # Demain
    ws_synthese.cell(row=row_idx, column=4, value=get_occupancy(2))   # Après-demain
    ws_synthese.cell(row=row_idx, column=5, value=get_occupancy(7))   # 7 jours
    ws_synthese.cell(row=row_idx, column=6, value=get_occupancy(30))  # 30 jours

    # Tarifs 7 prochains jours
    prices_7d = []
    for i in range(7):
        day_date = today + timedelta(days=i)
        for day in calendar:
            if day.get('date') == day_date.strftime('%Y-%m-%d'):
                prices_7d.append(day.get('price', 0))
                break

    for i, price in enumerate(prices_7d[:7]):
        ws_synthese.cell(row=row_idx, column=7+i, value=price)

    # Tarifs moyens par jour de semaine (8-38 jours)
    weekday_prices = defaultdict(list)
    for i in range(8, 39):
        day_date = today + timedelta(days=i)
        weekday = day_date.weekday()  # 0=Lundi, 6=Dimanche
        for day in calendar:
            if day.get('date') == day_date.strftime('%Y-%m-%d'):
                weekday_prices[weekday].append(day.get('price', 0))
                break

    weekday_names = ['Lundi', 'Mardi', 'Mercredi', 'Jeudi', 'Vendredi', 'Samedi', 'Dimanche']
    for wd in range(7):
        avg_price = round(sum(weekday_prices[wd]) / len(weekday_prices[wd]), 2) if weekday_prices[wd] else 0
        ws_synthese.cell(row=row_idx, column=14+wd, value=avg_price)

    print(f"   ✅ Données enregistrées pour {lid}")

# Sauvegarder
wb.save(EXCEL_FILE)
print(f"\n💾 Fichier sauvegardé : {EXCEL_FILE}")
