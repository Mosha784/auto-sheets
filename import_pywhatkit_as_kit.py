import gspread
from oauth2client.service_account import ServiceAccountCredentials

service_account_info = {
    "type": "service_account",
    "project_id": "hallowed-valve-343119",
    "private_key_id": "cf66bf72c23ae29054a8b92af78d22eae1b68819",
    "private_key": "-----BEGIN PRIVATE KEY-----\nMIIEvQIBADANBgkqhkiG9w0BAQEFAASCBKcwggSjAgEAAoIBAQCqucuCHhqd/pJj\nHKANxa1TuJAEEj1yMFMhd4Yq0luapqJIEnMvEK6jZV0DSXCeAha+izo73/LWQHDn\nsU3CVVrGtpGqH1uKhnTcO2GLEmWQ8nHNn4LQKp3yYlssRwO/qWyR9ptxSYK4Xv+k\n/FvjwbHPMkl+93v3IouTWbj6zAHUBgW3xb+B+9i/zw0ehpVCBR/gp8iG0siD3pEi\ngR1qIZjVX+ObP3IL+pWHROrY/hs+QsD6yq26A1LiTgMQLAR1owUBfsyWElILXcpC\n1OsnCirG2pLZmF5q68Ohww+4oK6JWhzEb54bZnVLTIYw1IC46GzoACkD41X9OmAe\ntUIynGu7AgMBAAECggEAJtafillI2tp3+N4hNyDaPmqFMLfpjJmbv8hOGF3EgxkX\nX+f6liFoaTl9AGtrmDaHcA+CTu6ycrU0OjEmrGf4f6420wnRLGFMInHLzfSAcIoH\nA60e+DZJukNP1HHPU4G6djYwxIPhngnWhHT4foao6abZ+21XoTAVqo7FuyA/5jic\nq6/kxqumsISj89+Gqyn0EadW22dRKqg7cApUjMo8sN3I5KlJz+VTpe2yS/M+WSoa\nSlb24H00OTD+CFeObZwgBlosgv9togOMh+We6JEiL98ok8F2/FLQ5tABiJcNliuY\nOGoIckyEgP9MWHv6mSpNHgWh47ucQTPNvl/63EOxMQKBgQDsZyCSnVKkMj4KrXZp\ngA5lMkm7hS9NjMUYsPD74nQswWKj/tE2t90vHswFYyAE8m0/UQ1RBptarb1gTJOJ\nj7WmFqfzX0O08KSiuMP0uh/6HoxCXJb9bNIhkT6PIXB/jlIrgwxrTr+8zi8kPseu\nIST96rNYf5u3h/j77D05Gh28EwKBgQC44ONB1PoFsN0wPp/LsTHCMQfK/SvS+772\nHL45ezKo3G71OCZYtGgREI05BRAijJg7ry2wpCgztnEFO6lCimWtCVTMhZrmxt/8\nS6XNgQ4zl/QOiV4ffU4acD8gTa/c4gtaSg3uMjwKqNXKn2a6SyqBVutAWBcyiuPC\nfw45jC62uQKBgQDlDD6JD5kksfFe0xapvYM1FXZPFAny73OAKuAyjQTW4EA8eQYo\nKBlMMGCoz5QUdvbWpCds3CPlxfR4u3kvjWgIlmb/7MtjIs3BQ5fJJBUbeEGZgrBg\ntvEZyOp+L34aeMCwm/aKefBYdMVELve1hTOcOayvEGTFfB8Hp6riCqXItQKBgEd4\nMIJHievfRnKbEv0UX+75M1EGdAWY6maMEAF6ncfnh0Fm1nQeMci/BEkRqv4gKc2Q\n1/HcU+pB0gk62iDuDYZKAC0cTRh/syD+QXdjN5E8Yc2ozukPcL0JvW2Ier7B56+c\nxyvY4ZshT5yH6JeF7UWYy1LRew4/4PJUWbRne7uJAoGAbQeyLa4izXRf4+4d05k3\nTZsIM1ZRY6gU8pqeZqxUefWnTa315/dQMyvm60yP2vu05cxiLvKD11roUKqbQB9f\nLZLVTCozRrCYhiUsvlcJr8NH2iJm8EMETuWqWumi5LLkiFGMmvuIdGP3/pI7vx6Z\n9gKsKlDpb8hplfbgRv8BtXM=\n-----END PRIVATE KEY-----\n",
    "client_email": "n8n-service-689@hallowed-valve-343119.iam.gserviceaccount.com",
    "client_id": "103820270853123050546",
    "auth_uri": "https://accounts.google.com/o/oauth2/auth",
    "token_uri": "https://oauth2.googleapis.com/token",
    "auth_provider_x509_cert_url": "https://www.googleapis.com/oauth2/v1/certs",
    "client_x509_cert_url": "https://www.googleapis.com/robot/v1/metadata/x509/n8n-service-689@hallowed-valve-343119.iam.gserviceaccount.com",
    "universe_domain": "googleapis.com"
}

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
