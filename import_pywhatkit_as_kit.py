import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# تحميل بيانات الخدمة من ملف خارجي
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope  = [
    'https://spreadsheets.google.com/feeds',
    'https://www.googleapis.com/auth/drive'
]
creds  = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)

SRC_URL = 'https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit'
DST_URL = 'https://docs.google.com/spreadsheets/d/1CmFB5Om8xO2qr-nRY8YdhSP07m3JgnhlyF54uSpB6Gg/edit'

src_ws = client.open_by_url(SRC_URL).worksheet('Sheet15')
dst_ws = client.open_by_url(DST_URL).worksheet('request')

all_values = src_ws.get('H:AC')
data_rows  = all_values[1:]

col_b       = dst_ws.col_values(2)
first_empty = len(col_b) + 1

start_row  = first_empty
end_row    = start_row + len(data_rows) - 1
dest_range = f'A{start_row}:V{end_row}'

dst_ws.update(dest_range, data_rows)
print(f'✅ تم لصق {len(data_rows)} صفّاً في النطاق {dest_range} دون إضافة أعمدة جديدة.')
