import datetime
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
from google.auth.transport.requests import Request as GARequest

# ---------- حساب الخدمة ----------
SERVICE_ACCOUNT_FILE = r"C:\Users\mosha\Downloads\service_account.json"
SCOPES = [
    "https://www.googleapis.com/auth/spreadsheets.readonly",
    "https://www.googleapis.com/auth/forms.body",
]

# ---------- التهيئة (أضِف أو عدِّل هنا) ----------
CONFIGS = [
    {
        "form_id": "12z1B1EfSX8zPKNiiqz7T6NeyCroESfYy_zvFvARMKes",
        "question_title": "Description",
        "spreadsheet_id": "10RYC8r3yqpkHFZtRa7ITD_hSfS7FdCXXHV6kzy5l-68",
        "sheet_range": "MS - Gulf & Electronic!E:E",
    },
    {
        "form_id": "1DIQwmPdDVQ7iHNBw3cDf6R8meXoUTRKG0VHH0uVsVho",
        "question_title": "Description",
        "spreadsheet_id": "10RYC8r3yqpkHFZtRa7ITD_hSfS7FdCXXHV6kzy5l-68",
        "sheet_range": "MS - KSA!E:E",
    },
    {
        "form_id": "1v8aG2IvOPxmGA7jBHQtL4rM5Ye7B-xyXYa6ZDoCFx64",
        "question_title": "Description",
        "spreadsheet_id": "10RYC8r3yqpkHFZtRa7ITD_hSfS7FdCXXHV6kzy5l-68",
        "sheet_range": "MS - UAE!E:E",
    },
]

# ---------- الاعتماد ----------
def get_creds():
    creds = Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )
    creds.refresh(GARequest())
    return creds


# ---------- جلب القيم الفريدة من الشيت ----------
def fetch_unique_values(creds, spreadsheet_id, sheet_range):
    """يقرأ القيم من العمود المحدَّد متجاوزًا صفّ العناوين (الصف 1)."""
    sheets = build("sheets", "v4", credentials=creds)
    rows = (
        sheets.spreadsheets()
        .values()
        .get(spreadsheetId=spreadsheet_id, range=sheet_range)
        .execute()
        .get("values", [])
    )

    # تخطَّ الصف الأول لأنه هيدر
    rows = rows[1:]

    seen, unique = set(), []
    for r in rows:
        if r and (v := r[0].strip()) and v not in seen:
            seen.add(v)
            unique.append(v)
    return unique


# ---------- إيجاد index عنصر الـ Dropdown ----------
def find_dropdown_index(creds, form_id, title):
    forms = build("forms", "v1", credentials=creds)
    items = forms.forms().get(formId=form_id).execute().get("items", [])
    for idx, item in enumerate(items):
        cq = (
            item.get("questionItem", {})
            .get("question", {})
            .get("choiceQuestion")
        )
        if item.get("title") == title and cq and cq.get("type") == "DROP_DOWN":
            return idx
    raise ValueError(f"Dropdown '{title}' غير موجود في النموذج {form_id}")


# ---------- تحديث خيارات Dropdown ----------
def update_dropdown(creds, form_id, item_index, title, options):
    forms = build("forms", "v1", credentials=creds)
    api_opts = [{"value": opt} for opt in options]
    body = {
        "requests": [
            {
                "updateItem": {
                    "location": {"index": item_index},
                    "item": {
                        "title": title,
                        "questionItem": {
                            "question": {
                                "choiceQuestion": {
                                    "type": "DROP_DOWN",
                                    "options": api_opts,
                                    "shuffle": False,
                                }
                            }
                        },
                    },
                    "updateMask": (
                        "title,questionItem.question.choiceQuestion.options"
                    ),
                }
            }
        ]
    }
    forms.forms().batchUpdate(formId=form_id, body=body).execute()
    print(
        f"✅ [{title}] – {form_id[:10]}… محدث بـ {len(options)} خيار"
        f"  ({datetime.datetime.now():%Y-%m-%d %H:%M:%S})"
    )


# ---------- تشغيل لجميع التكوينات ----------
def main():
    creds = get_creds()
    for cfg in CONFIGS:
        opts = fetch_unique_values(
            creds, cfg["spreadsheet_id"], cfg["sheet_range"]
        )
        if not opts:
            print(f"⚠️ لا قيم في {cfg['sheet_range']}")
            continue
        idx = find_dropdown_index(
            creds, cfg["form_id"], cfg["question_title"]
        )
        update_dropdown(
            creds,
            cfg["form_id"],
            idx,
            cfg["question_title"],
            opts,
        )


if __name__ == "__main__":
    main()
