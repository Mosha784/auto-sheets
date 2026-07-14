import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# تحميل بيانات الخدمة من ملف خارجي
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

# ---------- شيت 1 ----------
SHEET_URL1 = 'https://docs.google.com/spreadsheets/d/1GqiFgV5EesNaL3JGWkMPQfw7GjTpYXakK8Dtz9jPt2A/edit'
src_tab1 = 'Raw Data Missing'
dst_tab1 = 'PR Follow UP'

# نقل عمود A إلى C
src_ws1 = client.open_by_url(SHEET_URL1).worksheet(src_tab1)
raw1 = src_ws1.col_values(1)
last_row1 = len(raw1)
src_values1 = [v for v in raw1[1:] if v.strip()]  # عمود A من الصف 2
src_data1 = [[v] for v in src_values1]

dst_ws1 = client.open_by_url(SHEET_URL1).worksheet(dst_tab1)
dst_col1 = dst_ws1.col_values(3)
first_empty_row1 = len(dst_col1) + 1
end_row1 = first_empty_row1 + len(src_data1) - 1
if src_data1:
    dst_ws1.update(values=src_data1, range_name=f"C{first_empty_row1}:C{end_row1}")
    # مسح المصدر بعد النقل — بدون هذا كانت نفس البيانات تتنقل مرة تانية كل 30 دقيقة
    src_ws1.batch_clear([f"A2:A{last_row1}"])
    print(f"✅ [شيت 1] نقل {len(src_data1)} صف من '{src_tab1}' إلى '{dst_tab1}' في العمود C بدءًا من الصف {first_empty_row1} (وتم مسح المصدر)")
else:
    print("⚠️ [شيت 1] لا يوجد بيانات للنقل")

# ---------- شيت 2 ----------
SHEET_URL2 = 'https://docs.google.com/spreadsheets/d/1-Uw-4PNDuedzPmMXqnzPkNerVaTRBJ2wm88wfxJYZV8/edit?pli=1'
src_tab2 = 'Raw Data Missing'
dst_tab2 = 'Raw Data Iraq'

# نقل عمود E إلى B
# ⚠️ ملحوظة: الكود الأصلي كان مكتوب فيه تعليق "عمود D" لكنه يقرأ col_values(5) = عمود E.
# لو المقصود فعلاً عمود D غيّر الرقم إلى 4 وغيّر batch_clear إلى D2:D
src_ws2 = client.open_by_url(SHEET_URL2).worksheet(src_tab2)
raw2 = src_ws2.col_values(5)
last_row2 = len(raw2)
src_values2 = [v for v in raw2[1:] if v.strip()]
src_data2 = [[v] for v in src_values2]

dst_ws2 = client.open_by_url(SHEET_URL2).worksheet(dst_tab2)
dst_col2 = dst_ws2.col_values(2)
first_empty_row2 = len(dst_col2) + 1
end_row2 = first_empty_row2 + len(src_data2) - 1
if src_data2:
    dst_ws2.update(values=src_data2, range_name=f"B{first_empty_row2}:B{end_row2}")
    src_ws2.batch_clear([f"E2:E{last_row2}"])
    print(f"✅ [شيت 2] نقل {len(src_data2)} صف من '{src_tab2}' إلى '{dst_tab2}' في العمود B بدءًا من الصف {first_empty_row2} (وتم مسح المصدر)")
else:
    print("⚠️ [شيت 2] لا يوجد بيانات للنقل")
