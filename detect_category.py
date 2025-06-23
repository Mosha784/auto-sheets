import json
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

SHEET_URL = 'https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit'
TAB_NAME = 'Missing In Form'

CATEGORY_KEYWORDS = {
    "Electronics": [
        "tv", "لابتوب", "موبايل", "mobile", "هاتف", "تلفزيون", "لاب توب", "سماعة", "airpod", "هيدفون", "الكترونيات",
        "tablet", "سامسونج", "xiaomi", "smart", "iphone", "ساعة ذكية", "سماعات"
    ],
    "Home": [
        "sofa", "كنبة", "table", "كرسي", "كرسى", "chair", "سرير", "bed", "ثلاجة", "fridge", "microwave", "ميكروويف",
        "cooker", "موقد", "خلاط", "blender", "ستارة", "curtain", "أثاث", "furniture", "منزلية"
    ],
    "Consumables": [
        "شامبو", "shampoo", "صابون", "soap", "شيكولاتة", "chocolate", "عصير", "juice", "مياه", "water", "اكل", "food",
        "مستهلك", "consumable", "tea", "قهوة", "snack", "وجبة"
    ],
    "Leisure": [
        "كرة", "ball", "عجلة", "bicycle", "لعبة", "puzzle", "lego", "game", "leisure", "سكيت", "سباحة", "tennis",
        "camp", "outdoor", "رياضة", "خيمة", "رحلات"
    ],
    "Fashion": [
        "قميص", "shirt", "تيشيرت", "t-shirt", "بنطلون", "pant", "dress", "فستان", "حذاء", "shoe", "جاكيت", "jacket",
        "fashion", "ملابس", "skirt", "جيب", "jean", "حقيبة", "bag", "scarf", "حزام", "belt", "قبعة", "hat", "sneaker"
    ],
}

def extract_name_from_link(link):
    if not link:
        return ""
    m = re.search(r'/dp/([^/?]+)', link)
    if m:
        return m.group(1).replace('-', ' ')
    m = re.search(r'/p/([^/?]+)', link)
    if m:
        return m.group(1).replace('-', ' ')
    # كود عام - اسم آخر جزء من اللينك
    name = link.split('/')[-1].split('?')[0].replace('-', ' ')
    if len(name) < 3 or name.isdigit():
        return ""
    return name

def detect_category(text):
    if not text:
        return ""
    text = text.lower()
    for cat, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if re.search(r'\b' + re.escape(kw) + r'\b', text):
                return cat
    return "Unknown"

ws = client.open_by_url(SHEET_URL).worksheet(TAB_NAME)
data = ws.get_all_values()

header = data[0]
rows = data[1:]

count = 0
for idx, row in enumerate(rows, start=2):  # start=2 لأن أول صف هيدر
    name = row[3] if len(row) > 3 else ""  # D
    link = row[7] if len(row) > 7 else ""  # H
    category = row[8] if len(row) > 8 else ""  # I

    if not category.strip():
        name_from_link = extract_name_from_link(link)
        merged_text = f"{name} {name_from_link} {link}"
        cat = detect_category(merged_text)
        ws.update_cell(idx, 9, cat)
        print(f"Row {idx}: {name[:25]}... | {name_from_link[:25]}... → {cat}")
        count += 1

print(f"\n✅ تم تحديث {count} صف في التصنيف تلقائيًا (باستخدام الاسم واللينك).")
