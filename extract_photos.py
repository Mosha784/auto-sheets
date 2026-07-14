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

# تحميل بيانات حساب الخدمة (Service Account) من الملف
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")


def cell(row, i):
    """قراءة آمنة لخلية حتى لو الصف أقصر من المتوقع"""
    return row[i] if len(row) > i else ''


# ==========================================================
# 1) نسخ M:U إلى A:I — دفعة واحدة + مسح المصدر بعد النسخ
#    (المسح ضروري: بدونه نفس الصفوف كانت بتتنسخ تاني كل 30 دقيقة)
# ==========================================================
print("🔁 Copying M:U to A:I ...")
data = worksheet.get_all_values()

rows_to_copy = []
clear_ranges = []
for sheet_row, row in enumerate(data[1:], start=2):
    values = [cell(row, i) for i in range(12, 21)]  # M..U
    if any(v.strip() for v in values):
        rows_to_copy.append(values)
        clear_ranges.append(f"M{sheet_row}:U{sheet_row}")

if rows_to_copy:
    first_empty = next(
        (i + 1 for i, row in enumerate(data) if not cell(row, 0).strip()),
        len(data) + 1
    )
    end_row = first_empty + len(rows_to_copy) - 1
    # كتابة كل الصفوف في طلب واحد بدل طلب لكل صف (توفير كبير في حصة الـ API)
    worksheet.update(values=rows_to_copy, range_name=f"A{first_empty}:I{end_row}")
    # مسح المصدر حتى لا تتكرر نفس الصفوف في التشغيل القادم
    worksheet.batch_clear(clear_ranges)
    print(f"✅ Copied {len(rows_to_copy)} rows and cleared M:U source.")
else:
    print("ℹ️ Nothing new in M:U to copy.")

# ==========================================================
# 2) استخراج الصور — الكتابة تتجمع في دفعات بدل update_cell لكل صف
# ==========================================================
data = worksheet.get_all_values()
col_g = [cell(row, 6) for row in data]
col_h = [cell(row, 7) for row in data]

pending_updates = []  # [{'range': 'G5', 'values': [['url']]}, ...]


def queue_update(row_num, url):
    pending_updates.append({'range': f'G{row_num}', 'values': [[url]]})
    if len(pending_updates) >= 10:
        flush_updates()


def flush_updates():
    global pending_updates
    if pending_updates:
        worksheet.batch_update(pending_updates)
        pending_updates = []


def smart_get_image_url(link, page):
    if not link:
        return None

    # التعامل مع روابط Google Drive أو الصور المباشرة
    if "drive.google.com" in link:
        match = re.search(r"/d/([^/]+)", link)
        if match:
            url = f"https://drive.google.com/uc?export=download&id={match.group(1)}"
            print(f"DEBUG: Google Drive image found: {url}")
            return url
    if link.lower().endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
        print(f"DEBUG: Direct image link: {link}")
        return link

    # استخراج الصور من Amazon
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

    # og:image (يغطي Noon والمواقع العامة مثل ووردبريس)
    meta = page.query_selector('meta[property="og:image"]')
    if meta:
        content = meta.get_attribute("content")
        if content and content.strip():
            print(f"DEBUG: og:image found: {content}")
            return content

    # أول صورة كبيرة بصيغة معروفة في الصفحة
    img = page.query_selector('img[src*=".jpg"], img[src*=".jpeg"], img[src*=".png"], img[src*=".webp"]')
    if img:
        src = img.get_attribute("src")
        if src and src.strip():
            print(f"DEBUG: First big image found: {src}")
            return src

    # محاولة أخيرة: أول رابط صورة في الصفحة
    imgs = page.query_selector_all('img')
    for img in imgs:
        src = img.get_attribute('src')
        if src and any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png', '.webp']):
            print("DEBUG: Fallback first image:", src)
            return src

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
                    page.goto(link, timeout=60000, wait_until="domcontentloaded")
                    time.sleep(6)  # كان 15 ثانية لكل صف — كان بيخلي التشغيل يتجاوز الـ 30 دقيقة ويتداخل مع التشغيل التالي
                    img_url = smart_get_image_url(link, page)

                if img_url:
                    queue_update(idx + 1, img_url)
                    print(f"✅ Row {idx+1} done. {img_url}")
                else:
                    print(f"❌ No image for row {idx+1}")
                    failed_links.append(link)
                    failed_rows.append(idx + 1)
            except Exception as e:
                print(f"⚠️ Error row {idx+1}: {e}")
                failed_links.append(link)
                failed_rows.append(idx + 1)
    browser.close()

flush_updates()

# المحاولة الثانية باستخدام Selenium للروابط التي فشلت
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
            time.sleep(6)
            img_url = None

            try:
                og = driver.find_element(By.XPATH, '//meta[@property="og:image"]')
                img_url = og.get_attribute("content")
                print("OG IMAGE:", img_url)
            except Exception:
                pass

            if not img_url or ("noon" in link and "default" in (img_url or "")):
                try:
                    imgs = driver.find_elements(By.XPATH, '//img[contains(@src, ".jpg") or contains(@src, ".jpeg") or contains(@src, ".png")]')
                    all_img_srcs = []
                    for img in imgs:
                        src = img.get_attribute("src")
                        if src and "noon" in link and "product" in src and "default" not in src:
                            img_url = src
                            break
                        if src and "taobao" in link and ".jpg" in src:
                            img_url = src
                            break
                        if src and any(ext in src.lower() for ext in ['.jpg', '.jpeg', '.png']):
                            all_img_srcs.append(src)

                    if not img_url and all_img_srcs:
                        print("DEBUG (Selenium): Fallback to first image found.")
                        img_url = all_img_srcs[0]
                except Exception:
                    pass

            if img_url:
                queue_update(row_num, img_url)
                print(f"✅ Row {row_num} done (via Selenium). {img_url}")
            else:
                print(f"❌ Still no image for row {row_num}")
        except Exception as e:
            print(f"⚠️ Error row {row_num} in Selenium: {e}")
    driver.quit()
    flush_updates()

print("🎉 Process Finished (Playwright + Selenium fallback)")
