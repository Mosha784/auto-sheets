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

# 2. دالة تنظيف الرابط للحصول على أعلى جودة
def clean_alibaba_url(url):
    if not url: return url
    # إزالة لاحقات الحجم وتحويل الرابط لبروتوكول كامل
    url = re.sub(r'_\d+x\d+.*$', '', url)
    if url.startswith('//'):
        url = "https:" + url
    return url

# 3. الدالة الذكية لاستخراج الصور (Playwright)
def smart_get_image_url(link, page):
    if not link: return None

    # روابط صور مباشرة أو جوجل درايف
    if "drive.google.com" in link:
        match = re.search(r"/d/([^/]+)", link)
        return f"https://drive.google.com/uc?export=download&id={match.group(1)}" if match else None
            
    if link.lower().endswith(('.jpg', '.jpeg', '.png', '.webp')):
        return link

    # منطق علي بابا و 1688 الصارم
    if "alibaba.com" in link or "1688.com" in link:
        try:
            # محاولة الاستخراج من سكريبتات الصفحة (أضمن طريقة لتجنب اللوجو)
            script_content = page.evaluate('''() => {
                const scripts = Array.from(document.querySelectorAll('script'));
                return scripts.map(s => s.innerText).join(' ');
            }''')
            
            # البحث عن رابط .jpg (صور المنتجات) وتجاهل الـ .png (اللوجو)
            img_match = re.search(r'(https:[^"]+?\.jpg)', script_content)
            if img_match:
                found_url = img_match.group(1).replace('\\u002F', '/')
                print(f"✅ Alibaba JSON Match: {found_url}")
                return clean_alibaba_url(found_url)

            # محاولة og:image بشرط ألا يكون png
            meta = page.query_selector('meta[property="og:image"]')
            if meta:
                content = meta.get_attribute("content")
                if content and ".jpg" in content.lower():
                    return clean_alibaba_url(content)
        except: pass

    # أمازون ونون والمواقع الأخرى
    if "amazon." in link:
        img = page.query_selector("#landingImage")
        if img: return img.get_attribute("src")

    meta = page.query_selector('meta[property="og:image"]')
    if meta:
        content = meta.get_attribute("content")
        if content and content.strip(): return content

    return None

# --- بداية التشغيل ---

print("🔍 Starting Extraction Process...")
data = worksheet.get_all_values()
failed_links = []
failed_rows = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(user_agent="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36")
    page = context.new_page()
    
    for idx in range(1, len(data)):
        img_val = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''
        
        if (not img_val or not img_val.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Processing {link[:50]}...")
            try:
                if "drive.google.com" in link or link.lower().endswith(('.jpg', '.jpeg')):
                    img_url = smart_get_image_url(link, None)
                else:
                    page.goto(link, timeout=60000, wait_until="domcontentloaded")
                    time.sleep(8) # وقت كافٍ لتحميل السكريبتات
                    img_url = smart_get_image_url(link, page)
                
                if img_url:
                    worksheet.update_cell(idx+1, 7, img_url)
                    print(f"✅ Success: {img_url}")
                else:
                    failed_links.append(link)
                    failed_rows.append(idx+1)
            except:
                failed_links.append(link)
                failed_rows.append(idx+1)
    browser.close()

# --- محاولة Selenium (للروابط التي فشلت) ---
if failed_links:
    print("\n🚨 Retrying failed links with Selenium...")
    options = Options()
    options.add_argument('--headless=new')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    for i, link in enumerate(failed_links):
        row_num = failed_rows[i]
        try:
            driver.get(link)
            time.sleep(10)
            # البحث عن أول صورة .jpg وليست لوجو
            images = driver.find_elements(By.TAG_NAME, 'img')
            for img in images:
                src = img.get_attribute("src")
                if src and ".jpg" in src.lower() and "logo" not in src.lower():
                    final_url = clean_alibaba_url(src)
                    worksheet.update_cell(row_num, 7, final_url)
                    print(f"✅ Selenium Fixed Row {row_num}")
                    break
        except: pass
    driver.quit()

print("🎉 Task Completed.")
