import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
import time
import re

# 1. إعداد الوصول لجوجل شيت
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")

pending_updates = []


def queue_update(row_num, url):
    pending_updates.append({'range': f'G{row_num}', 'values': [[url]]})
    if len(pending_updates) >= 10:
        flush_updates()


def flush_updates():
    global pending_updates
    if pending_updates:
        worksheet.batch_update(pending_updates)
        pending_updates = []


def get_clean_image_address(page):
    """محاكاة عملية 'Copy Image Address' بدقة"""
    try:
        img_address = page.evaluate('''() => {
            // محاولة إيجاد الصورة النشطة في معرض علي بابا
            const mainImg = document.querySelector('.main-image-thumb-item img') || 
                            document.querySelector('.module-pdp-main-image img') ||
                            document.querySelector('.image-viewer img') ||
                            document.querySelector('.detail-main-image');
            
            if (mainImg) {
                let src = mainImg.getAttribute('data-src') || mainImg.getAttribute('src');
                if (src && src.includes('.jpg')) return src;
            }

            // البحث عن أول صورة JPG كبيرة (تجنب اللوجو PNG)
            const allImages = Array.from(document.querySelectorAll('img'));
            const productImg = allImages.find(img => 
                img.src.includes('.jpg') && 
                !img.src.includes('tps-') && 
                img.width > 250
            );
            
            return productImg ? productImg.src : null;
        }''')

        if img_address:
            # تنظيف الرابط للحصول على الصورة الأصلية عالية الجودة
            img_address = re.sub(r'_\d+x\d+.*$', '', img_address)
            if img_address.startswith('//'):
                img_address = "https:" + img_address
            return img_address
    except Exception as e:
        # كان except: بدون نوع — بيخفي أي خطأ حقيقي
        print(f"⚠️ evaluate error: {e}")
        return None
    return None


# --- التنفيذ ---
print("🚀 Starting Copy Image Address Script...")
data = worksheet.get_all_values()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    page = context.new_page()

    for idx in range(1, len(data)):
        img_val = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''

        if (not img_val or not img_val.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Processing {link[:40]}...")
            try:
                page.goto(link, timeout=60000, wait_until="domcontentloaded")
                time.sleep(5)  # انتظار تحميل الصور

                final_address = get_clean_image_address(page)

                if final_address:
                    queue_update(idx + 1, final_address)
                    print(f"✅ Success: {final_address}")
                else:
                    print(f"❌ Failed to find image address")
            except Exception as e:
                print(f"⚠️ Error: {str(e)}")

    browser.close()

flush_updates()
print("🎉 Script Finished.")
