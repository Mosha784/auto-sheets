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

def clean_final_url(url):
    if not url: return None
    # إزالة أحجام التصغير للحصول على الصورة الأصلية
    url = re.sub(r'_\d+x\d+.*$', '', url)
    if url.startswith('//'): url = "https:" + url
    return url

def try_extract_methods(link, page):
    """تجربة 5 طرق مختلفة لاستخراج صورة المنتج"""
    
    # --- الطريقة 1: الطلب المباشر (Fast HTTP Request) ---
    try:
        headers = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36"}
        res = requests.get(link, headers=headers, timeout=10)
        if res.status_code == 200:
            # البحث عن og:image بشرط أن تكون JPG ولا تحتوي على وسم اللوجو tps-
            match = re.search(r'property="og:image"\s+content="([^"]+\.jpg[^"]*)"', res.text)
            if match and "tps-" not in match.group(1):
                return clean_final_url(match.group(1))
    except: pass

    # --- الطريقة 2: البحث في JSON الصفحة (Scripts Parsing) ---
    try:
        script_data = page.evaluate('''() => {
            const scripts = Array.from(document.querySelectorAll('script'));
            return scripts.map(s => s.innerText).join(' ');
        }''')
        # البحث عن روابط الصور داخل مصفوفة الصور في السكريبت
        img_match = re.search(r'(https:[^"]+?\.jpg)', script_data)
        if img_match and "tps-" not in img_match.group(1):
            return clean_final_url(img_match.group(1).replace('\\u002F', '/'))
    except: pass

    # --- الطريقة 3: فحص الصور الرئيسية (Gallery Selectors) ---
    selectors = [
        "img.main-image", ".module-pdp-main-image img", 
        ".image-viewer img", "img.detail-main-image"
    ]
    for selector in selectors:
        try:
            img = page.query_selector(selector)
            if img:
                src = img.get_attribute("src") or img.get_attribute("data-src")
                if src and ".jpg" in src.lower() and "tps-" not in src:
                    return clean_final_url(src)
        except: continue

    # --- الطريقة 4: البحث عن أكبر صورة JPG في الصفحة ---
    try:
        best_img = page.evaluate('''() => {
            const imgs = Array.from(document.querySelectorAll('img'));
            const filtered = imgs.filter(i => i.src.includes('.jpg') && !i.src.includes('tps-') && !i.src.includes('logo'));
            if (filtered.length === 0) return null;
            // ترتيب الصور حسب الحجم التقديري
            filtered.sort((a, b) => (b.width * b.height) - (a.width * a.height));
            return filtered[0].src;
        }''')
        if best_img: return clean_final_url(best_img)
    except: pass

    # --- الطريقة 5: الـ Meta Tag عبر Playwright ---
    try:
        meta_img = page.get_attribute('meta[property="og:image"]', "content")
        if meta_img and ".jpg" in meta_img.lower() and "tps-" not in meta_img:
            return clean_final_url(meta_img)
    except: pass

    return None

# --- دورة العمل الأساسية ---
print("🚀 Starting Multi-Method Extraction...")
data = worksheet.get_all_values()

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    page = context.new_page()

    for idx in range(1, len(data)):
        img_cell = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''
        
        if (not img_cell or not img_cell.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Attempting {link[:40]}...")
            try:
                page.goto(link, timeout=60000, wait_until="domcontentloaded")
                time.sleep(5) # انتظار رندرة الصفحة
                
                final_url = try_extract_methods(link, page)
                
                if final_url:
                    worksheet.update_cell(idx+1, 7, final_url)
                    print(f"✅ Method Success: {final_url}")
                else:
                    print(f"❌ All 5 methods failed for Row {idx+1}")
            except:
                print(f"⚠️ Connection Error on Row {idx+1}")
    
    browser.close()
print("🎉 Process Finished.")
