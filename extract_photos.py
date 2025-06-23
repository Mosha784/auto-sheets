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

print("🔁 Copying M:U to A:I ...")
data = worksheet.get_all_values()
rows = [row for row in data[1:] if any(row[12:21])]
first_empty = next((i for i, row in enumerate(data) if not row[0].strip()), len(data))
for row in rows:
    values = row[12:21]
    if any(values):
        row_index = first_empty + 1
        worksheet.update(f"A{row_index}:I{row_index}", [values])
        first_empty += 1
print("✅ Done copying.")

# تحديث البيانات بعد النسخ
data = worksheet.get_all_values()
col_g = [row[6] if len(row) > 6 else '' for row in data]
col_h = [row[7] if len(row) > 7 else '' for row in data]

def smart_get_image_url(link, page):
    if not link: return None

    # روابط Google Drive أو صورة مباشرة
    if "drive.google.com" in link:
        match = re.search(r"/d/([^/]+)", link)
        if match:
            url = f"https://drive.google.com/uc?export=download&id={match.group(1)}"
            print(f"DEBUG: Google Drive image found: {url}")
            return url
    if link.lower().endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
        print(f"DEBUG: Direct image link: {link}")
        return link

    # Amazon (id="landingImage")
    if "amazon." in link:
        img = page.query_selector("#landingImage")
        if img:
            src = img.get_attribute("src")
            if src and src.strip():
                print(f"DEBUG: Amazon landingImage found: {src}")
                return src
        meta = page.query_selector('meta[property="og:image"]')
        if meta:
            content = meta.get_attribute("content")
            if content and content.strip():
                print(f"DEBUG: Amazon og:image found: {content}")
                return content

    # Noon (og:image فقط لو متاحة)
    if "noon.com" in link:
        meta = page.query_selector('meta[property="og:image"]')
        if meta:
            content = meta.get_attribute("content")
            if content and content.strip():
                print(f"DEBUG: Noon og:image found: {content}")
                return content

    # مواقع ووردبريس - og:image
    meta = page.query_selector('meta[property="og:image"]')
    if meta:
        content = meta.get_attribute("content")
        if content and content.strip():
            print(f"DEBUG: og:image found: {content}")
            return content

    # أول صورة كبيرة في الصفحة
    img = page.query_selector('img[src*=".jpg"], img[src*=".jpeg"], img[src*=".png"], img[src*=".webp"]')
    if img:
        src = img.get_attribute("src")
        if src and src.strip():
            print(f"DEBUG: First big image found: {src}")
            return src

    # Fallback قوي: اطبع كل الصور الكبيرة الموجودة
    imgs = page.query_selector_all('img')
    all_img_srcs = []
    for img in imgs:
        src = img.get_attribute('src')
        if src and any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png', '.webp']):
            all_img_srcs.append(src)
    if all_img_srcs:
        print("DEBUG: All found img srcs (fallback):", all_img_srcs)
        # رجع أول صورة كـ fallback (لو ده مناسب ليك)
        return all_img_srcs[0]

    print("DEBUG: No image found at all.")
    return None

print("🔍 Extracting images for all empty G with link in H ...")
failed_links = []
failed_rows = []
with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    page = browser.new_page(user_agent=user_agent)
    for idx in range(1, len(data)):
        img_g = col_g[idx] if idx < len(col_g) else ''
        link = col_h[idx] if idx < len(col_h) else ''
        if (not img_g or not img_g.strip()) and link and link.strip():
            print(f"🌐 Row {idx+1}: Fetching image from {link}")
            try:
                if "drive.google.com" in link or link.lower().endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
                    img_url = smart_get_image_url(link, page=None)
                else:
                    page.goto(link, timeout=60000)
                    time.sleep(15)  # زودت الـ sleep هنا!
                    img_url = smart_get_image_url(link, page)
                if img_url:
                    worksheet.update_cell(idx+1, 7, img_url)
                    print(f"✅ Row {idx+1} done. {img_url}")
                else:
                    print(f"❌ No image for row {idx+1}")
                    failed_links.append(link)
                    failed_rows.append(idx+1)
            except Exception as e:
                print(f"⚠️ Error row {idx+1}: {e}")
                failed_links.append(link)
                failed_rows.append(idx+1)
    browser.close()

# الآن نحاول على الروابط اللي فشلت باستخدام Selenium
if failed_links:
    print("\n🚨 Trying Selenium for failed links...")
    user_agent = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36"
    options = Options()
    options.add_argument('--headless=new')
    options.add_argument('--no-sandbox')
    options.add_argument('--disable-dev-shm-usage')
    options.add_argument(f"user-agent={user_agent}")
    options.add_argument("--window-size=1920,1080")

    driver = webdriver.Chrome(service=ChromeService(ChromeDriverManager().install()), options=options)

    for i, link in enumerate(failed_links):
        row_num = failed_rows[i]
        print(f"\n🔗 {link}")
        try:
            driver.get(link)
            time.sleep(15)
            img_url = None
            # جرب og:image الأول
            try:
                og = driver.find_element(By.XPATH, '//meta[@property="og:image"]')
                img_url = og.get_attribute("content")
                print("OG IMAGE:", img_url)
            except:
                pass
            # لو مفيش og:image أو فاضي جرب صور المنتجات في الصفحة
            if not img_url or ("noon" in link and "default" in (img_url or "")):
                try:
                    imgs = driver.find_elements(By.XPATH, '//img[contains(@src, ".jpg") or contains(@src, ".jpeg") or contains(@src, ".png")]')
                    all_img_srcs = []
                    for img in imgs:
                        src = img.get_attribute("src")
                        # Noon جرب أول صورة كبيرة
                        if src and "noon" in link and "product" in src and "default" not in src:
                            img_url = src
                            break
                        # Taobao أول صورة jpg
                        if src and "taobao" in link and ".jpg" in src:
                            img_url = src
                            break
                        if src and any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png']):
                            all_img_srcs.append(src)
                    if not img_url and all_img_srcs:
                        print("DEBUG (Selenium): All found img srcs (fallback):", all_img_srcs)
                        img_url = all_img_srcs[0]
                except:
                    pass
            if img_url:
                worksheet.update_cell(row_num, 7, img_url)
                print(f"✅ Row {row_num} done. {img_url}")
            else:
                print(f"❌ No image for row {row_num}")
        except Exception as e:
            print(f"⚠️ Error row {row_num}: {e}")
    driver.quit()

print("🎉 Done (Playwright + Selenium fallback)")
