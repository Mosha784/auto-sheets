import json
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright
import time
import re
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

# 1. إعداد الوصول لجوجل شيت
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")

def clean_url(url):
    if not url: return url
    # تنظيف روابط علي بابا للحصول على الصورة الأصلية
    url = re.sub(r'_\d+x\d+.*$', '', url)
    if url.startswith('//'): url = "https:" + url
    return url

def smart_get_image_url(link, page):
    if not link: return None

    # روابط مباشرة
    if "drive.google.com" in link:
        match = re.search(r"/d/([^/]+)", link)
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}" if match else None
    
    # منطق علي بابا و 1688 المحدث
    if "alibaba.com" in link or "1688.com" in link:
        # البحث في الـ Meta Tags أولاً (الأكثر دقة)
        # نستخدم evaluate للحصول على الـ content مباشرة من الـ DOM
        og_image = page.evaluate('''() => {
            const meta = document.querySelector('meta[property="og:image"]');
            return meta ? meta.content : null;
        }''')
        
        # إذا وجدنا رابط ينتهي بـ .jpg فهو المنتج (تجنب الـ .png لأنه اللوجو)
        if og_image and ".jpg" in og_image.lower():
            return clean_url(og_image)
        
        # محاولة بديلة: البحث عن أول صورة كبيرة .jpg في الـ Gallery
        img_src = page.evaluate('''() => {
            const imgs = Array.from(document.querySelectorAll('img'));
            // ابحث عن صورة في الـ main gallery أو تحتوي على كلمة product
            const productImg = imgs.find(i => 
                i.src.includes('.jpg') && 
                !i.src.includes('logo') && 
                (i.className.includes('main') || i.className.includes('detail'))
            );
            return productImg ? productImg.src : null;
        }''')
        if img_src: return clean_url(img_src)

    # أمازون والمواقع الأخرى
    meta = page.query_selector('meta[property="og:image"]')
    if meta:
        content = meta.get_attribute("content")
        if content: return clean_url(content)

    return None

# --- بداية التنفيذ ---
print("🔍 Starting Extraction Process...")
data = worksheet.get_all_values()
failed_links = []
failed_rows = []

with sync_playwright() as p:
    # استخدام متصفح بوضع "الرأس" (non-headless) أحياناً يساعد في تخطي حماية علي بابا
    browser = p.chromium.launch(headless=True) 
    context = browser.new_context(
        user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        viewport={'width': 1920, 'height': 1080}
    )
    page = context.new_page()
    
    for idx in range(1, len(data)):
        # التأكد من صحة الـ index (العمود G هو index 6 والعمود H هو index 7)
        img_val = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''
        
        if (not img_val or not img_val.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Processing {link[:60]}...")
            try:
                # محاولة الدخول وتخطي الـ بوب أب إن وجد
                page.goto(link, timeout=60000, wait_until="load")
                time.sleep(5) # انتظار قليل للتحميل
                
                img_url = smart_get_image_url(link, page)
                
                if img_url and "tps-297-40.png" not in img_url: # استثناء اللوجو صراحة
                    worksheet.update_cell(idx+1, 7, img_url)
                    print(f"✅ Success: {img_url}")
                else:
                    print(f"❌ Failed to find product image for row {idx+1}")
                    failed_links.append(link)
                    failed_rows.append(idx+1)
            except Exception as e:
                print(f"⚠️ Error: {e}")
                failed_links.append(link)
                failed_rows.append(idx+1)
    browser.close()

# --- جزء السيلينيوم للروابط الفاشلة (مع فلتر اللوجو) ---
if failed_links:
    print("\n🚨 Retrying with Selenium...")
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument("user-agent=Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36")
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    for i, link in enumerate(failed_links):
        row_num = failed_rows[i]
        try:
            driver.get(link)
            time.sleep(7)
            # جرب استخراج الـ og:image أولاً بالـ Selenium
            og = driver.find_element(By.XPATH, '//meta[@property="og:image"]')
            url = og.get_attribute("content")
            if url and ".jpg" in url.lower():
                worksheet.update_cell(row_num, 7, clean_url(url))
                print(f"✅ Selenium Fixed Row {row_num}")
                continue
            
            # محاولة أخيرة بالصور العادية بشرط الامتداد .jpg
            imgs = driver.find_elements(By.TAG_NAME, 'img')
            for img in imgs:
                src = img.get_attribute("src")
                if src and ".jpg" in src.lower() and "logo" not in src.lower():
                    worksheet.update_cell(row_num, 7, clean_url(src))
                    print(f"✅ Selenium Found JPG for Row {row_num}")
                    break
        except: pass
    driver.quit()

print("🎉 Task Completed.")
