import os
import re
import requests
from datetime import datetime, timedelta
from openpyxl import load_workbook
from collections import defaultdict

APIFY_TOKEN = os.environ['APIFY_TOKEN']
EXCEL_FILE = 'reporting_airbnb.xlsx'

print("Chargement du fichier Excel...")
wb = load_workbook(EXCEL_FILE)
print("Onglets trouves :", wb.sheetnames)

ws_param = wb['paramètre']
listing_ids = []
for row in range(2, 8):
    cell_value = ws_param.cell(row=row, column=4).value
    if cell_value and 'airbnb' in str(cell_value).lower():
        match = re.search(r'/rooms/(\d+)', str(cell_value))
        if match:
            listing_ids.append(match.group(1))
            print("Listing trouve :", match.group(1), "(ligne", row, ")")

print("Total listings :", len(listing_ids))

def get_airbnb_data(listing_id):
    url = f"https://api.apify.com/v2/acts/dtrungtin~airbnb-scraper/run-sync-get-dataset-items?token={APIFY_TOKEN}"

    # CORRECTION : startUrls comme tableau de strings, pas d'objets
    payload = {
        "startUrls": [f"https://www.airbnb.fr/rooms/{listing_id}"],
        "maxItems": 1,
        "includeCalendar": True,
        "includeReviews": False
    }

    try:
        response = requests.post(url, json=payload, timeout=120)
        if response.status_code == 200:
            data = response.json()
            if isinstance(data, list) and len(data) > 0:
                return data[0]
            elif isinstance(data, dict):
                return data
        print(f"Erreur API {response.status_code} pour {listing_id}")
        print(f"Detail : {response.text[:500]}")  # Log pour voir l'erreur exacte
        return None
    except Exception as e:
        print("Exception pour", listing_id, ":", e)
        return None

ws_synthese = wb['synthèse']
today = datetime.now().date()

for i, listing_id in enumerate(listing_ids):
    print("Traitement listing", listing_id, "...")
    data = get_airbnb_data(listing_id)
    if not data:
        print("Pas de donnees pour", listing_id)
        continue

    # Extraction du calendrier (structure variable selon l'acteur)
    calendar = data.get('calendar', []) or data.get('calendarMonths', []) or data.get('availabilityCalendar', [])
    all_days = []
    if calendar and isinstance(calendar, list):
        if isinstance(calendar[0], dict) and 'days' in calendar[0]:
            for month in calendar:
                all_days.extend(month.get('days', []))
        else:
            all_days = calendar

    def get_occupancy(days_list):
        if not days_list:
            return 0
        booked = sum(1 for d in days_list if not d.get('available', True))
        return round((booked / len(days_list)) * 100, 1)

    today_str = today.strftime('%Y-%m-%d')
    tomorrow = (today + timedelta(days=1)).strftime('%Y-%m-%d')
    day_after = (today + timedelta(days=2)).strftime('%Y-%m-%d')

    days_today = [d for d in all_days if d.get('date') == today_str]
    days_tomorrow = [d for d in all_days if d.get('date') == tomorrow]
    days_day_after = [d for d in all_days if d.get('date') == day_after]

    end_7 = today + timedelta(days=7)
    end_30 = today + timedelta(days=30)
    days_7 = [d for d in all_days if today <= datetime.strptime(d.get('date', '1900-01-01'), '%Y-%m-%d').date() <= end_7]
    days_30 = [d for d in all_days if today <= datetime.strptime(d.get('date', '1900-01-01'), '%Y-%m-%d').date() <= end_30]

    row_idx = i + 2

    ws_synthese.cell(row=row_idx, column=2, value=get_occupancy(days_today))
    ws_synthese.cell(row=row_idx, column=3, value=get_occupancy(days_tomorrow))
    ws_synthese.cell(row=row_idx, column=4, value=get_occupancy(days_day_after))
    ws_synthese.cell(row=row_idx, column=5, value=get_occupancy(days_7))
    ws_synthese.cell(row=row_idx, column=6, value=get_occupancy(days_30))

    for j in range(7):
        target_date = (today + timedelta(days=j)).strftime('%Y-%m-%d')
        day_data = next((d for d in all_days if d.get('date') == target_date), None)
        price = day_data.get('price', 0) if day_data else 0
        ws_synthese.cell(row=row_idx, column=7+j, value=price)

    weekday_prices = defaultdict(list)
    for j in range(8, 39):
        target_date = today + timedelta(days=j)
        day_data = next((d for d in all_days if d.get('date') == target_date.strftime('%Y-%m-%d')), None)
        if day_data and day_data.get('price'):
            weekday_prices[target_date.weekday()].append(day_data['price'])

    for wd in range(7):
        avg = round(sum(weekday_prices[wd]) / len(weekday_prices[wd]), 2) if weekday_prices[wd] else 0
        ws_synthese.cell(row=row_idx, column=14+wd, value=avg)

    print("Donnees enregistrees pour", listing_id)

wb.save(EXCEL_FILE)
print("Fichier sauvegarde :", EXCEL_FILE)
