import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
import time
import re
import requests

# 1. إعداد الوصول لجوجل شيت
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")

def clean_url(url):
    """تنظيف الرابط للحصول على الصورة الأصلية بجودة عالية"""
    if not url: return None
    # إزالة لاحقات المقاسات مثل _300x300.jpg
    url = re.sub(r'_\d+x\d+.*$', '', url)
    if url.startswith('//'): url = "https:" + url
    return url

def get_image_fast(link):
    """جلب الصورة عبر طلب HTTP سريع لتخطي حماية المتصفحات واللوجو"""
    headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
    try:
        response = requests.get(link, headers=headers, timeout=10)
        if response.status_code == 200:
            # البحث عن وسم og:image مع استبعاد الـ png (اللوجو)
            # نركز على الروابط التي تنتهي بـ .jpg
            match = re.search(r'<meta[^>]*property="og:image"[^>]*content="([^"]+\.jpg[^"]*)"', response.text)
            if match:
                img_url = match.group(1)
                if "tps-" not in img_url:
                    return clean_url(img_url)
    except: pass
    return None

# --- بداية التنفيذ الرئيسي ---
print("🔍 Starting Extraction Process...")
data = worksheet.get_all_values()

# استخدام Playwright فقط كخيار بديل (Fallback)
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    page = context.new_page()

    for idx in range(1, len(data)):
        # G هو العمود 7، H هو العمود 8
        img_val = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''
        
        if (not img_val or not img_val.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Processing {link[:50]}...")
            
            # 1. محاولة الحل السريع أولاً (أسرع وأدق لعلي بابا)
            img_url = get_image_fast(link)
            
            # 2. إذا فشل، نستخدم المتصفح
            if not img_url:
                try:
                    page.goto(link, timeout=60000, wait_until="domcontentloaded")
                    time.sleep(5)
                    # استخراج og:image مع فلتر الـ jpg والـ tps
                    img_url = page.evaluate('''() => {
                        const meta = document.querySelector('meta[property="og:image"]');
                        if (meta && meta.content.includes(".jpg") && !meta.content.includes("tps-")) {
                            return meta.content;
                        }
                        return null;
                    }''')
                except: pass

            if img_url:
                worksheet.update_cell(idx+1, 7, clean_url(img_url))
                print(f"✅ Success: {img_url}")
            else:
                print(f"❌ Failed to find product image")
                
    browser.close()
print("🎉 Task Completed Successfully.")
