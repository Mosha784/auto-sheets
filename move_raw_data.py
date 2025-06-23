import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# تحميل بيانات الخدمة من ملف خارجي
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

# رابط الشيت
SHEET_URL = 'https://docs.google.com/spreadsheets/d/1GqiFgV5EesNaL3JGWkMPQfw7GjTpYXakK8Dtz9jPt2A/edit'

# التابات
src_tab = 'Raw Data Missing'
dst_tab = 'PR Follow UP'

# قراءة العمود A من Raw Data Missing (بدءًا من الصف الثاني)
src_ws = client.open_by_url(SHEET_URL).worksheet(src_tab)
src_values = src_ws.col_values(1)[1:]  # [1:] عشان نهمل الهيدر

# حذف القيم الفارغة
src_values = [v for v in src_values if v.strip()]

if not src_values:
    print("⚠️ لا يوجد بيانات للنقل")
    exit(0)

# تجهيز القيم بشكل عمودي (كل قيمة داخل list وحدها)
src_data = [[v] for v in src_values]

# جلب العمود C من PR Follow UP لمعرفة آخر صف فيه داتا
dst_ws = client.open_by_url(SHEET_URL).worksheet(dst_tab)
dst_col = dst_ws.col_values(3)
first_empty_row = len(dst_col) + 1

# تحديث البيانات الجديدة في آخر الصفوف
end_row = first_empty_row + len(src_data) - 1
dst_ws.update(f"C{first_empty_row}:C{end_row}", src_data)

print(f"✅ تم نقل {len(src_data)} صف من '{src_tab}' إلى '{dst_tab}' في العمود C، بدءًا من الصف {first_empty_row}")
