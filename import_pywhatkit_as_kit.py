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

all_values = src_ws.get('H:AC')
last_src_row = len(all_values)  # آخر صف فيه بيانات في النطاق المصدر

# تخطي الصفوف الفاضية تمامًا (كان بينسخ صفوف فاضية) + توحيد طول الصفوف
data_rows = [
    row + [''] * (NUM_COLS - len(row))
    for row in all_values[1:]
    if any(str(c).strip() for c in row)
]

if not data_rows:
    # كان بيضرب error لو مفيش بيانات (نطاق عكسي مثل A5:V4)
    print('ℹ️ لا توجد بيانات جديدة في Sheet15 H:AC — لا شيء للنسخ.')
else:
    col_b = dst_ws.col_values(2)
    first_empty = len(col_b) + 1

    start_row = first_empty
    end_row = start_row + len(data_rows) - 1
    dest_range = f'A{start_row}:V{end_row}'

    # استخدام الوسائط المسماة (الترتيب القديم deprecated في gspread الحديث)
    dst_ws.update(values=data_rows, range_name=dest_range)

    # مسح المصدر بعد النسخ — بدون هذا كانت كل الصفوف تتنسخ من جديد كل 30 دقيقة
    # لو مش عايز المسح (عايز تحتفظ بالبيانات في Sheet15) علّق السطر التالي،
    # لكن ساعتها لازم تعمل آلية تانية لتمييز الصفوف المنسوخة وإلا هترجع مشكلة التكرار.
    src_ws.batch_clear([f'H2:AC{last_src_row}'])

    print(f'✅ تم لصق {len(data_rows)} صفًّا في النطاق {dest_range} وتم مسح المصدر لمنع التكرار.')
