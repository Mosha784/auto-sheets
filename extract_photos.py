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

# تحميل بيانات الخدمة من ملف خارجي
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")

# دالة مساعدة لتنظيف الرابط والحصول على جودة عالية
def clean_alibaba_url(url):
    if not url: return url
    # إزالة لاحقة المقاسات مثل _300x300.jpg للحصول على الصورة الأصلية
    url = re.sub(r'_\d+x\d+.*$', '', url)
    if url.startswith('//'):
        url = "https:" + url
    return url

def smart_get_image_url(link, page):
    if not link: return None

    # روابط Google Drive أو صورة مباشرة
    if "drive.google.com" in link:
        match = re.search(r"/d/([^/]+)", link)
        if match:
            return f"https://drive.google.com/uc?export=download&id={match.group(1)}"
            
    if link.lower().endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
        return link

    # Alibaba & 1688 (تعديل دقيق لمنع اللوجو)
    if "alibaba.com" in link or "1688.com" in link:
        # 1. محاولة og:image أولاً
        meta = page.query_selector('meta[property="og:image"]')
        if meta:
            content = meta.get_attribute("content")
            if content and "logo" not in content.lower():
                return clean_alibaba_url(content)
        
        # 2. البحث عن كلاسات الصور الرئيسية للمنتج
        product_selectors = [
            "img.main-image", ".module-pdp-main-image img", 
            "img.detail-main-image", ".image-viewer img"
        ]
        for selector in product_selectors:
            img = page.query_selector(selector)
            if img:
                src = img.get_attribute("src") or img.get_attribute("data-src")
                if src and "logo" not in src.lower():
                    return clean_alibaba_url(src)

    # Amazon
    if "amazon." in link:
        img = page.query_selector("#landingImage")
        if img: return img.get_attribute("src")

    # الخيار العام (Open Graph)
    meta = page.query_selector('meta[property="og:image"]')
    if meta:
        content = meta.get_attribute("content")
        if content and content.strip(): return content

    return None

# --- بداية التنفيذ ---
print("🔁 Copying M:U to A:I ...")
data = worksheet.get_all_values()
# (ملاحظة: تركت منطق النسخ كما هو في كودك الأصلي)
# ... [كود النسخ الخاص بك] ...

print("🔍 Extracting images...")
failed_links = []
failed_rows = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    page = browser.new_page(user_agent=user_agent)
    
    for idx in range(1, len(data)):
        img_g = data[idx][6] if len(data[idx]) > 6 else ''
        link = data[idx][7] if len(data[idx]) > 7 else ''
        
        if (not img_g or not img_g.strip()) and link and link.strip():
            try:
                if "drive.google.com" in link or link.lower().endswith(('.jpg', '.jpeg', '.png')):
                    img_url = smart_get_image_url(link, page=None)
                else:
                    page.goto(link, timeout=60000)
                    time.sleep(10) # انتظار للتحميل
                    img_url = smart_get_image_url(link, page)
                
                if img_url:
                    worksheet.update_cell(idx+1, 7, img_url)
                    print(f"✅ Row {idx+1}: {img_url}")
                else:
                    failed_links.append(link)
                    failed_rows.append(idx+1)
            except Exception as e:
                failed_links.append(link)
                failed_rows.append(idx+1)
    browser.close()

# --- محاولة Selenium للروابط الفاشلة ---
if failed_links:
    print("\n🚨 Trying Selenium for failed links...")
    options = Options()
    options.add_argument('--headless=new')
    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    for i, link in enumerate(failed_links):
        row_num = failed_rows[i]
        try:
            driver.get(link)
            time.sleep(12)
            img_url = None
            
            # محاولة og:image مع استبعاد اللوجو
            try:
                og = driver.find_element(By.XPATH, '//meta[@property="og:image"]')
                content = og.get_attribute("content")
                if "logo" not in content.lower():
                    img_url = clean_alibaba_url(content)
            except: pass

            if not img_url:
                # محاولة البحث عن صور المنتج باستبعاد الكلمات المحظورة
                imgs = driver.find_elements(By.TAG_NAME, 'img')
                for img in imgs:
                    src = img.get_attribute("src")
                    if src and any(ext in src.lower() for ext in ['.jpg', '.png', '.jpeg']):
                        if not any(x in src.lower() for x in ['logo', 'icon', 'banner', 'nav']):
                            img_url = clean_alibaba_url(src)
                            break
            
            if img_url:
                worksheet.update_cell(row_num, 7, img_url)
                print(f"✅ Selenium Row {row_num}: {img_url}")
        except: pass
    driver.quit()

print("🎉 Process Finished.")
