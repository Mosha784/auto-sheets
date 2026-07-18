import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# تحميل بيانات الخدمة من ملف خارجي
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

# بدون مسح المصدر (المصدر معادلات) — منع التكرار بيتم بمقارنة القيم
# بالموجود فعلًا في عمود الوجهة وتخطّي أي قيمة موجودة قبل كده.
# ⚠️ ملحوظة: لو ممكن تتكرر نفس القيمة بشكل مشروع (قيمتين متطابقتين لصفقتين مختلفتين)
# هتتنقل مرة واحدة بس — لو ده وارد عندك قولّي نعمل مفتاح مقارنة أدق.

# ---------- شيت 1 ----------
SHEET_URL1 = 'https://docs.google.com/spreadsheets/d/1GqiFgV5EesNaL3JGWkMPQfw7GjTpYXakK8Dtz9jPt2A/edit'
src_tab1 = 'Raw Data Missing'
dst_tab1 = 'PR Follow UP'

# نقل عمود A (المصدر) إلى C (الوجهة) — القيم الجديدة فقط
src_ws1 = client.open_by_url(SHEET_URL1).worksheet(src_tab1)
raw1 = [v.strip() for v in src_ws1.col_values(1)[1:] if v.strip()]  # عمود A من الصف 2

dst_ws1 = client.open_by_url(SHEET_URL1).worksheet(dst_tab1)
dst_col1 = dst_ws1.col_values(3)
existing1 = {v.strip() for v in dst_col1 if v.strip()}

src_data1 = []
for v in raw1:
    if v not in existing1:
        src_data1.append([v])
        existing1.add(v)  # حتى لا تتكرر نفس القيمة داخل نفس التشغيل

first_empty_row1 = len(dst_col1) + 1
end_row1 = first_empty_row1 + len(src_data1) - 1
if src_data1:
    dst_ws1.update(values=src_data1, range_name=f"C{first_empty_row1}:C{end_row1}")
    print(f"✅ [شيت 1] نقل {len(src_data1)} قيمة جديدة من '{src_tab1}' إلى '{dst_tab1}' في العمود C بدءًا من الصف {first_empty_row1} (بدون مسح المصدر)")
else:
    print("ℹ️ [شيت 1] لا توجد قيم جديدة — كل قيم المصدر موجودة بالفعل في الوجهة")

# ---------- شيت 2 ----------
SHEET_URL2 = 'https://docs.google.com/spreadsheets/d/1-Uw-4PNDuedzPmMXqnzPkNerVaTRBJ2wm88wfxJYZV8/edit?pli=1'
src_tab2 = 'Raw Data Missing'
dst_tab2 = 'Raw Data Iraq'

# نقل عمود E (المصدر) إلى B (الوجهة) — القيم الجديدة فقط
# ⚠️ التعليق الأصلي كان يقول "عمود D" لكن الكود يقرأ col_values(5) = عمود E.
# لو المقصود فعلاً عمود D غيّر الرقم إلى 4.
src_ws2 = client.open_by_url(SHEET_URL2).worksheet(src_tab2)
raw2 = [v.strip() for v in src_ws2.col_values(5)[1:] if v.strip()]

dst_ws2 = client.open_by_url(SHEET_URL2).worksheet(dst_tab2)
dst_col2 = dst_ws2.col_values(2)
existing2 = {v.strip() for v in dst_col2 if v.strip()}

src_data2 = []
for v in raw2:
    if v not in existing2:
        src_data2.append([v])
        existing2.add(v)

first_empty_row2 = len(dst_col2) + 1
end_row2 = first_empty_row2 + len(src_data2) - 1
if src_data2:
    dst_ws2.update(values=src_data2, range_name=f"B{first_empty_row2}:B{end_row2}")
    print(f"✅ [شيت 2] نقل {len(src_data2)} قيمة جديدة من '{src_tab2}' إلى '{dst_tab2}' في العمود B بدءًا من الصف {first_empty_row2} (بدون مسح المصدر)")
else:
    print("ℹ️ [شيت 2] لا توجد قيم جديدة — كل قيم المصدر موجودة بالفعل في الوجهة")
