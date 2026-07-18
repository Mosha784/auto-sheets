# -*- coding: utf-8 -*-
"""
extract_photos.py — نسخة محسّنة لسحب صور المنتجات من اللينكات

النسخ من M:U إلى A:I بدون مسح المصدر (لأن M:U معادلات):
بدل المسح، بنقارن كل صف في M:U بالصفوف الموجودة فعلًا في A:I —
لو الصف موجود قبل كده يتم تخطّيه، فمفيش تكرار ومفيش لمس للمعادلات.

استخراج الصور:
1) لينكات مباشرة / Google Drive  → بدون أي طلب شبكة
2) requests (سريع جدًا)         → قراءة og:image من الـ HTML الخام
3) Playwright (fallback فقط)     → للمواقع اللي بتبني الصفحة بالجافاسكريبت
+ تنضيف رابط الصورة لأعلى جودة + التحقق أنه صورة فعلًا + فلترة اللوجوهات
+ الكتابة في الشيت بالدُفعات لتفادي حد 60 كتابة/دقيقة
"""

import json
import re
import time

import requests
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from playwright.sync_api import sync_playwright

# ================= إعداد Google Sheets =================
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")

HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
}

BAD_URL_WORDS = ('logo', 'icon', 'sprite', 'favicon', 'placeholder', 'default',
                 'loading', 'blank', 'avatar', 'badge', 'flag', 'payment')


def cell(row, i):
    return row[i] if len(row) > i else ''


def row_key(values):
    """مفتاح مقارنة موحّد للصف (نص متقصّص لكل خلية)"""
    return tuple((v or '').strip() for v in values)


# ================= 1) نسخ M:U إلى A:I (بدون مسح المصدر) =================
print("🔁 Copying new rows from M:U to A:I (no source clearing) ...")
data = worksheet.get_all_values()

# الصفوف الموجودة فعلًا في A:I — أي صف من M:U مطابق ليها يتم تخطّيه
existing = set()
for row in data[1:]:
    key = row_key([cell(row, i) for i in range(0, 9)])
    if any(key):
        existing.add(key)

rows_to_copy = []
for row in data[1:]:
    values = [cell(row, i) for i in range(12, 21)]  # M..U
    key = row_key(values)
    if any(key) and key not in existing:
        rows_to_copy.append(values)
        existing.add(key)  # حتى لا يتكرر نفس الصف مرتين داخل نفس التشغيل

if rows_to_copy:
    first_empty = next((i + 1 for i, row in enumerate(data) if not cell(row, 0).strip()), len(data) + 1)
    end_row = first_empty + len(rows_to_copy) - 1
    worksheet.update(values=rows_to_copy, range_name=f"A{first_empty}:I{end_row}")
    print(f"✅ Copied {len(rows_to_copy)} new rows (M:U formulas left untouched).")
else:
    print("ℹ️ No new rows in M:U — everything already exists in A:I.")


# ================= أدوات استخراج الصور =================
def normalize_url(url, base_link=""):
    """توحيد الرابط: بروتوكول + إزالة معاملات القياس لأعلى جودة"""
    if not url:
        return None
    url = url.strip()
    if url.startswith('data:'):
        return None
    if url.startswith('//'):
        url = 'https:' + url
    if not url.startswith('http'):
        return None

    # أمازون: إزالة توكن الحجم للحصول على الصورة الكاملة
    if 'media-amazon' in url or 'images-amazon' in url or 'amazon.' in (base_link or ''):
        url = re.sub(r'\._[A-Z0-9_,]+_\.', '.', url)

    # علي بابا / علي إكسبريس: إزالة لاحقة القياس _250x250xz.jpg
    if 'alicdn' in url or 'aliexpress' in url or 'alibaba' in url:
        url = re.sub(r'_\d+x\d+[a-z]*(\.\w+)?$', '', url)

    return url


def looks_like_product_image(url, base_link=""):
    """فلترة اللوجوهات والأيقونات وصور noon الافتراضية"""
    if not url:
        return False
    low = url.lower()
    if any(w in low for w in BAD_URL_WORDS):
        return False
    if low.endswith('.svg') or low.endswith('.gif'):
        return False
    if 'noon' in (base_link or '') and 'default' in low:
        return False
    return True


def verify_image(url):
    """التحقق أن الرابط يرجّع صورة فعلًا قبل كتابته في الشيت"""
    try:
        r = requests.get(url, headers=HEADERS, timeout=15, stream=True)
        ct = r.headers.get('Content-Type', '')
        r.close()
        return r.status_code == 200 and ct.startswith('image/')
    except requests.RequestException:
        return False


OG_PATTERNS = [
    r'<meta[^>]+property=["\']og:image(?::secure_url)?["\'][^>]+content=["\']([^"\']+)["\']',
    r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image(?::secure_url)?["\']',
    r'<meta[^>]+name=["\']twitter:image["\'][^>]+content=["\']([^"\']+)["\']',
    r'<link[^>]+rel=["\']image_src["\'][^>]+href=["\']([^"\']+)["\']',
]


def extract_from_html(html, base_link):
    """استخراج رابط الصورة من HTML خام"""
    for pat in OG_PATTERNS:
        m = re.search(pat, html, re.IGNORECASE)
        if m:
            url = normalize_url(m.group(1), base_link)
            if looks_like_product_image(url, base_link):
                return url, "og:image"

    if 'amazon.' in base_link:
        m = re.search(r'"hiRes"\s*:\s*"([^"]+)"', html) or re.search(r'"large"\s*:\s*"([^"]+)"', html)
        if m:
            url = normalize_url(m.group(1), base_link)
            if looks_like_product_image(url, base_link):
                return url, "amazon hiRes"

    for m in re.finditer(r'<img[^>]+src=["\']([^"\']+\.(?:jpe?g|png|webp)[^"\']*)["\']', html, re.IGNORECASE):
        url = normalize_url(m.group(1), base_link)
        if looks_like_product_image(url, base_link):
            return url, "first img"

    return None, None


def get_image_via_requests(link):
    """المحاولة السريعة: جلب الـ HTML الخام بدون متصفح"""
    try:
        r = requests.get(link, headers=HEADERS, timeout=25)
        if r.status_code != 200:
            print(f"DEBUG: requests status {r.status_code}")
            return None
        url, how = extract_from_html(r.text, link)
        if url:
            print(f"DEBUG: [requests/{how}] {url}")
        return url
    except requests.RequestException as e:
        print(f"DEBUG: requests failed: {e}")
        return None


def get_image_via_playwright(link, page):
    """الـ fallback: للمواقع اللي بتبني الصفحة بالجافاسكريبت (علي بابا/تاوباو...)"""
    page.goto(link, timeout=60000, wait_until="domcontentloaded")
    time.sleep(5)

    url = page.evaluate('''() => {
        const mainImg = document.querySelector('.main-image-thumb-item img') ||
                        document.querySelector('.module-pdp-main-image img') ||
                        document.querySelector('.image-viewer img') ||
                        document.querySelector('.detail-main-image') ||
                        document.querySelector('#landingImage');
        if (mainImg) {
            const src = mainImg.getAttribute('data-src') || mainImg.getAttribute('src');
            if (src) return src;
        }
        const og = document.querySelector('meta[property="og:image"]');
        if (og && og.content) return og.content;
        // أكبر صورة ظاهرة في الصفحة (بدل "أول صورة" اللي كانت بتجيب اللوجو)
        let best = null, bestArea = 0;
        for (const img of document.querySelectorAll('img')) {
            const src = img.currentSrc || img.src || '';
            if (!/\\.(jpe?g|png|webp)/i.test(src)) continue;
            const area = (img.naturalWidth || img.width) * (img.naturalHeight || img.height);
            if (area > bestArea && area > 250 * 250) { best = src; bestArea = area; }
        }
        return best;
    }''')

    url = normalize_url(url, link)
    if url and looks_like_product_image(url, link):
        print(f"DEBUG: [playwright] {url}")
        return url
    return None


def resolve_image(link, page):
    """المنطق الكامل: مباشر → requests → playwright، مع التحقق النهائي"""
    if "drive.google.com" in link:
        m = re.search(r"/d/([^/]+)", link)
        if m:
            return f"https://drive.google.com/uc?export=download&id={m.group(1)}"

    if link.lower().split('?')[0].endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
        return link

    url = get_image_via_requests(link)
    if url and verify_image(url):
        return url

    try:
        url = get_image_via_playwright(link, page)
    except Exception as e:
        print(f"DEBUG: playwright failed: {e}")
        url = None
    if url and verify_image(url):
        return url

    return None


# ================= 2) استخراج الصور =================
print("🔍 Extracting images for all empty G with link in H ...")
data = worksheet.get_all_values()

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


failed = []

with sync_playwright() as p:
    browser = p.chromium.launch(headless=True)
    context = browser.new_context(user_agent=HEADERS["User-Agent"], locale="en-US")
    page = context.new_page()

    for idx in range(1, len(data)):
        row = data[idx]
        img_g = cell(row, 6)
        link = cell(row, 7).strip()

        if img_g.strip() or not link:
            continue

        print(f"🌐 Row {idx+1}: {link[:60]}")
        try:
            img_url = resolve_image(link, page)
            if img_url:
                queue_update(idx + 1, img_url)
                print(f"✅ Row {idx+1}: {img_url}")
            else:
                failed.append(idx + 1)
                print(f"❌ Row {idx+1}: no valid image found")
        except Exception as e:
            failed.append(idx + 1)
            print(f"⚠️ Row {idx+1} error: {e}")

    browser.close()

flush_updates()

if failed:
    print(f"\n❗ Rows with no image: {failed}")
print("🎉 Process Finished.")
