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
    """تنظيف الرابط للحصول على أعلى جودة وتصحيح البروتوكول"""
    if not url: return None
    url = re.sub(r'_\d+x\d+.*$', '', url) # إزالة أحجام التصغير
    if url.startswith('//'): url = "https:" + url
    return url

def get_product_image(link, page):
    """استراتيجية متعددة الطبقات لضمان جلب صورة المنتج وليس اللوجو"""
    
    # الطبقة الأولى: محاولة سريعة عبر وسم og:image باستخدام Requests
    # هذه الطريقة تتخطى حماية المتصفحات في كثير من الأحيان
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        res = requests.get(link, headers=headers, timeout=10)
        if res.status_code == 200:
            # البحث عن رابط ينتهي بـ .jpg (صور المنتجات) ويستبعد .png (اللوجو)
            match = re.search(r'property="og:image"\s+content="([^"]+\.jpg[^"]*)"', res.text)
            if match:
                img = match.group(1)
                if "tps-" not in img: return clean_url(img)
    except: pass

    # الطبقة الثانية: استخدام Playwright لاستخراج الصورة من داخل الـ DOM
    try:
        # البحث عن وسوم og:image أو الصور داخل معرض الصور الرئيسي
        img_url = page.evaluate('''() => {
            // 1. فحص الـ Meta tags
            const og = document.querySelector('meta[property="og:image"]');
            if (og && og.content.includes(".jpg") && !og.content.includes("tps-")) return og.content;
            
            // 2. فحص الصور الرئيسية في الصفحة
            const imgs = Array.from(document.querySelectorAll('img'));
            const productImg = imgs.find(i => 
                i.src.includes(".jpg") && 
                !i.src.includes("logo") && 
                !i.src.includes("tps-") &&
                (i.width > 200 || i.className.includes("main") || i.className.includes("detail"))
            );
            return productImg ? productImg.src : null;
        }''')
        if img_url: return clean_url(img_url)
    except: pass
    
    return None

# --- التنفيذ الرئيسي ---
print("🔍 Starting Final Extraction Process...")
data = worksheet.get_all_values()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    # استخدام User-Agent حديث لتجنب اكتشاف "البوت"
    context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36")
    page = context.new_page()

    for idx in range(1, len(data)):
        img_cell = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''
        
        if (not img_cell or not img_cell.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Processing {link[:50]}...")
            try:
                # محاولة جلب الصورة (سواء عبر طلب سريع أو متصفح)
                page.goto(link, timeout=60000, wait_until="domcontentloaded")
                time.sleep(3) # وقت بسيط لفك الحماية
                
                final_img = get_product_image(link, page)
                
                if final_img:
                    worksheet.update_cell(idx+1, 7, final_img)
                    print(f"✅ Success: {final_img}")
                else:
                    print(f"❌ Failed to find product image")
            except Exception as e:
                print(f"⚠️ Error on Row {idx+1}")
    
    browser.close()
print("🎉 Task Completed Successfully.")
