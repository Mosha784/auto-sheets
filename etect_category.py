import json
import re
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# إعدادات الاتصال
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

SHEET_URL = 'https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit'
TAB_NAME = 'Missing In Form'

# الكلمات المفتاحية لكل تصنيف
CATEGORY_KEYWORDS = {
    "Electronics": ["tv", "laptop", "mobile", "phone", "airpod", "headphone", "electronics", "tablet", "samsung", "xiaomi", "smart", "iphone"],
    "Home": ["sofa", "table", "chair", "bed", "home", "fridge", "microwave", "cooker", "blender", "curtain", "furniture"],
    "Consumables": ["shampoo", "soap", "chips", "biscuit", "water", "juice", "snack", "food", "consumable", "tea", "coffee", "chocolate"],
    "Leisure": ["ball", "bicycle", "puzzle", "lego", "game", "leisure", "skate", "swim", "tennis", "camp", "outdoor", "sport"],
    "Fashion": ["shirt", "t-shirt", "pant", "dress", "shoe", "jacket", "fashion", "skirt", "jean", "bag", "scarf", "belt", "hat", "sneaker"],
}

def detect_category(text):
    if not text:
        return ""
    text = text.lower()
    for cat, keywords in CATEGORY_KEYWORDS.items():
        for kw in keywords:
            if re.search(r'\b' + re.escape(kw) + r'\b', text):
                return cat
    return "Unknown"

# افتح التاب وجب البيانات
ws = client.open_by_url(SHEET_URL).worksheet(TAB_NAME)
data = ws.get_all_values()

header = data[0]
rows = data[1:]

count = 0
for idx, row in enumerate(rows, start=2):  # start=2 عشان رقم الصف في الشيت
    name = row[3] if len(row) > 3 else ""  # عمود D (اسم المنتج)
    link = row[7] if len(row) > 7 else ""  # عمود H (لينك المنتج)
    category = row[8] if len(row) > 8 else ""  # عمود I (التصنيف)

    # فقط الصفوف اللي التصنيف فيها فاضي
    if not category.strip():
        cat = detect_category(f"{name} {link}")
        ws.update_cell(idx, 9, cat)
        print(f"Row {idx}: {name[:30]}... → {cat}")
        count += 1

print(f"\n✅ تم تحديث {count} صف في التصنيف تلقائيًا.")
