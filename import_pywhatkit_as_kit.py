import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# تحميل بيانات الخدمة من ملف خارجي
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

SRC_URL = 'https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit'
DST_URL = 'https://docs.google.com/spreadsheets/d/1CmFB5Om8xO2qr-nRY8YdhSP07m3JgnhlyF54uSpB6Gg/edit'

src_ws = client.open_by_url(SRC_URL).worksheet('Sheet15')
dst_ws = client.open_by_url(DST_URL).worksheet('request')

NUM_COLS = 22  # H:AC = 22 عمود = A:V في الوجهة


def row_key(row):
    """مفتاح مقارنة موحّد للصف: 22 خلية نصية متقصّصة"""
    padded = list(row) + [''] * (NUM_COLS - len(row))
    return tuple(str(c).strip() for c in padded[:NUM_COLS])


# بدون مسح المصدر (Sheet15 معادلات) — منع التكرار بمقارنة كل صف
# بالصفوف الموجودة فعلًا في الوجهة A:V وتخطّي المطابق.
all_values = src_ws.get('H:AC')

src_rows = [
    row for row in all_values[1:]
    if any(str(c).strip() for c in row)  # تخطي الصفوف الفاضية تمامًا
]

dst_existing_raw = dst_ws.get('A:V')
existing = {row_key(row) for row in dst_existing_raw if any(str(c).strip() for c in row)}

new_rows = []
for row in src_rows:
    key = row_key(row)
    if key not in existing:
        new_rows.append(list(key))  # صف موحّد الطول 22 خلية
        existing.add(key)           # حتى لا يتكرر نفس الصف داخل نفس التشغيل

if not new_rows:
    print('ℹ️ لا توجد صفوف جديدة في Sheet15 — كل الصفوف موجودة بالفعل في الوجهة.')
else:
    col_b = dst_ws.col_values(2)
    first_empty = len(col_b) + 1

    start_row = first_empty
    end_row = start_row + len(new_rows) - 1
    dest_range = f'A{start_row}:V{end_row}'

    dst_ws.update(values=new_rows, range_name=dest_range)
    print(f'✅ تم لصق {len(new_rows)} صفًّا جديدًا في النطاق {dest_range} (بدون مسح المصدر — المعادلات كما هي).')
