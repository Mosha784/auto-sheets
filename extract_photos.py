# -*- coding: utf-8 -*-
"""
extract_photos.py — استخراج صور المنتجات حتى مع الحماية القوية

الفكرة: بدل محاولة كسر الحماية، نتجنبها بطبقات مرتبة من الأسرع للأثقل:

  الطبقة 0: لينكات مباشرة / Google Drive        → بدون شبكة
  الطبقة 1: حِيَل بدون سكرابينج                  → أمازون: ASIN → رابط صورة مباشر
                                                   من amazon-adsystem (لا تُفتح أي صفحة)
  الطبقة 2: requests بهيدرز متصفح كاملة          → محاولة عادية + محاولة بهوية موبايل
                                                   (صفحات الموبايل حمايتها أخف)
  الطبقة 3: Microlink API                        → خدمة معاينة لينكات تجلب الصفحة
                                                   من سيرفراتها هي (تنجح فيما يفشل من الرانر)
  الطبقة 4: Playwright stealth                   → متصفح مع إخفاء علامات الأتمتة

+ التحقق أن الناتج صورة فعلًا قبل الكتابة، وتنضيف الرابط لأعلى جودة،
+ الكتابة في الشيت بالدُفعات، والنسخ من M:U إلى A:I بدون مسح المعادلات.

ملاحظة: MICROLINK_KEY اختياري — النسخة المجانية محدودة الطلبات يوميًا.
لو عندك مفتاح حطه كـ secret باسم MICROLINK_KEY في GitHub وسيتم استخدامه تلقائيًا.
"""

import json
import os
import re
import time
from collections import Counter

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

MICROLINK_KEY = os.environ.get("MICROLINK_KEY", "").strip()

DESKTOP_HEADERS = {
    "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/124.0.0.0 Safari/537.36",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,image/avif,image/webp,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
    "Accept-Encoding": "gzip, deflate, br",
    "Connection": "keep-alive",
    "Upgrade-Insecure-Requests": "1",
    "Sec-Fetch-Dest": "document",
    "Sec-Fetch-Mode": "navigate",
    "Sec-Fetch-Site": "none",
    "Sec-Fetch-User": "?1",
    "sec-ch-ua": '"Chromium";v="124", "Google Chrome";v="124", "Not-A.Brand";v="99"',
    "sec-ch-ua-mobile": "?0",
    "sec-ch-ua-platform": '"Windows"',
}

MOBILE_HEADERS = {
    "User-Agent": "Mozilla/5.0 (iPhone; CPU iPhone OS 17_4 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/17.4 Mobile/15E148 Safari/604.1",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
    "Accept-Language": "en-US,en;q=0.9,ar;q=0.8",
}

BAD_URL_WORDS = ('logo', 'icon', 'sprite', 'favicon', 'placeholder', 'default',
                 'loading', 'blank', 'avatar', 'badge', 'flag', 'payment', 'captcha')


def cell(row, i):
    return row[i] if len(row) > i else ''


def row_key(values):
    return tuple((v or '').strip() for v in values)


# ================= 1) نسخ M:U إلى A:I (بدون مسح المعادلات) =================
print("🔁 Copying new rows from M:U to A:I (no source clearing) ...")
data = worksheet.get_all_values()

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
        existing.add(key)

if rows_to_copy:
    first_empty = next((i + 1 for i, row in enumerate(data) if not cell(row, 0).strip()), len(data) + 1)
    end_row = first_empty + len(rows_to_copy) - 1
    worksheet.update(values=rows_to_copy, range_name=f"A{first_empty}:I{end_row}")
    print(f"✅ Copied {len(rows_to_copy)} new rows (M:U formulas left untouched).")
else:
    print("ℹ️ No new rows in M:U.")


# ================= أدوات مشتركة =================
def normalize_url(url, base_link=""):
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
    # علي بابا / علي إكسبريس: إزالة لاحقة القياس
    if 'alicdn' in url or 'aliexpress' in url or 'alibaba' in url:
        url = re.sub(r'_\d+x\d+[a-z]*(\.\w+)?$', '', url)
    return url


def looks_like_product_image(url, base_link=""):
    if not url:
        return False
    low = url.lower()
    if any(w in low for w in BAD_URL_WORDS):
        return False
    if low.endswith('.svg'):
        return False
    if 'noon' in (base_link or '') and 'default' in low:
        return False
    # علي بابا: ملفات tps هي رسومات الموقع نفسه (زي رسمة الفيشة واليد) وليست منتجات
    if 'alicdn' in low and ('tps-' in low or '/tps/' in low):
        return False
    # صور منتجات علي بابا / علي إكسبريس الحقيقية تكون دائمًا تحت مسار /kf/
    if ('alibaba.com' in (base_link or '') or 'aliexpress' in (base_link or '')) \
            and 'alicdn' in low and '/kf/' not in low:
        return False
    return True


def verify_image(url, min_bytes=3000):
    """
    التحقق أن الرابط يرجّع صورة فعلًا وبحجم معقول
    (min_bytes يستبعد البكسلات الشفافة وصور الـ 1x1 اللي بترجع بدل صور المنتج)
    """
    try:
        r = requests.get(url, headers=DESKTOP_HEADERS, timeout=20, stream=True, allow_redirects=True)
        ct = r.headers.get('Content-Type', '')
        if r.status_code != 200 or not ct.startswith('image/'):
            r.close()
            return False
        cl = r.headers.get('Content-Length')
        if cl is not None:
            ok = int(cl) >= min_bytes
            r.close()
            return ok
        # لو مفيش Content-Length نقرأ أول جزء ونقيس
        chunk = next(r.iter_content(chunk_size=min_bytes + 1), b'')
        r.close()
        return len(chunk) >= min_bytes
    except (requests.RequestException, ValueError):
        return False


OG_PATTERNS = [
    r'<meta[^>]+property=["\']og:image(?::secure_url)?["\'][^>]+content=["\']([^"\']+)["\']',
    r'<meta[^>]+content=["\']([^"\']+)["\'][^>]+property=["\']og:image(?::secure_url)?["\']',
    r'<meta[^>]+name=["\']twitter:image["\'][^>]+content=["\']([^"\']+)["\']',
    r'<link[^>]+rel=["\']image_src["\'][^>]+href=["\']([^"\']+)["\']',
]


def extract_from_html(html, base_link):
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

    # بيانات JSON-LD المنظمة (كتير من المتاجر بتحط صورة المنتج فيها)
    for m in re.finditer(r'"image"\s*:\s*"(https?:[^"]+\.(?:jpe?g|png|webp)[^"]*)"', html, re.IGNORECASE):
        url = normalize_url(m.group(1), base_link)
        if looks_like_product_image(url, base_link):
            return url, "json-ld"

    for m in re.finditer(r'<img[^>]+src=["\']([^"\']+\.(?:jpe?g|png|webp)[^"\']*)["\']', html, re.IGNORECASE):
        url = normalize_url(m.group(1), base_link)
        if looks_like_product_image(url, base_link):
            return url, "first img"

    return None, None


# ================= الطبقة 1: حِيَل بدون سكرابينج =================
AMAZON_MARKETPLACES = {
    'amazon.com': ('ws-na', 'US'), 'amazon.ca': ('ws-na', 'CA'),
    'amazon.co.uk': ('ws-eu', 'GB'), 'amazon.de': ('ws-eu', 'DE'),
    'amazon.fr': ('ws-eu', 'FR'), 'amazon.it': ('ws-eu', 'IT'),
    'amazon.es': ('ws-eu', 'ES'), 'amazon.ae': ('ws-eu', 'AE'),
    'amazon.sa': ('ws-eu', 'SA'), 'amazon.eg': ('ws-eu', 'EG'),
    'amazon.in': ('ws-in', 'IN'),
}


def extract_asin(link):
    for pat in (r'/dp/([A-Z0-9]{10})', r'/gp/product/([A-Z0-9]{10})',
                r'/gp/aw/d/([A-Z0-9]{10})', r'/product/([A-Z0-9]{10})',
                r'[?&]asin=([A-Z0-9]{10})'):
        m = re.search(pat, link, re.IGNORECASE)
        if m:
            return m.group(1).upper()
    return None


def amazon_direct_image(link):
    """
    أمازون: بناء رابط الصورة مباشرة من الـ ASIN عبر خدمة amazon-adsystem
    (خدمة أمازون الرسمية لعرض صور المنتجات في الويدجتس) — لا تُفتح أي صفحة منتج
    ولا تتفعل أي حماية.
    """
    asin = extract_asin(link)
    if not asin:
        return None
    domain = next((d for d in AMAZON_MARKETPLACES if d in link), None)
    host, mp = AMAZON_MARKETPLACES.get(domain, ('ws-na', 'US'))
    url = (f"https://{host}.amazon-adsystem.com/widgets/q?"
           f"_encoding=UTF8&ASIN={asin}&Format=_SL800_&ID=AsinImage"
           f"&MarketPlace={mp}&ServiceVersion=20070822")
    print(f"DEBUG: [amazon-asin] trying {asin} ({mp})")
    return url


# ================= الطبقة 2: requests مقوّى =================
def get_image_via_requests(link):
    for attempt, headers in ((1, DESKTOP_HEADERS), (2, MOBILE_HEADERS)):
        try:
            r = requests.get(link, headers=headers, timeout=25, allow_redirects=True)
            if r.status_code == 200 and 'captcha' not in r.url.lower():
                url, how = extract_from_html(r.text, link)
                if url:
                    print(f"DEBUG: [requests#{attempt}/{how}] {url}")
                    return url
            else:
                print(f"DEBUG: requests#{attempt} status {r.status_code}")
        except requests.RequestException as e:
            print(f"DEBUG: requests#{attempt} failed: {e}")
        time.sleep(1.5)  # مهلة قصيرة قبل المحاولة الثانية
    return None


# ================= الطبقة 3: Microlink =================
def get_image_via_microlink(link):
    """
    خدمة معاينة لينكات (زي معاينات واتساب) — الجلب يتم من سيرفراتهم هم،
    فبتنجح مع مواقع بتحظر IPs الرانرز. النسخة المجانية محدودة الطلبات يوميًا.
    """
    try:
        params = {'url': link}
        headers = {}
        if MICROLINK_KEY:
            headers['x-api-key'] = MICROLINK_KEY
        r = requests.get('https://api.microlink.io/', params=params, headers=headers, timeout=40)
        if r.status_code != 200:
            print(f"DEBUG: microlink status {r.status_code}")
            return None
        j = r.json()
        img = (j.get('data') or {}).get('image') or {}
        url = normalize_url(img.get('url'), link)
        if url and looks_like_product_image(url, link):
            print(f"DEBUG: [microlink] {url}")
            return url
    except (requests.RequestException, ValueError) as e:
        print(f"DEBUG: microlink failed: {e}")
    return None


# ================= الطبقة 4: Playwright stealth =================
STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
window.chrome = window.chrome || { runtime: {} };
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en', 'ar'] });
Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
Object.defineProperty(navigator, 'platform', { get: () => 'Win32' });
"""


def get_image_via_playwright(link, page):
    page.goto(link, timeout=60000, wait_until="domcontentloaded")
    # سكرول بسيط يشغّل الـ lazy-loading ويبدو سلوكًا بشريًا
    try:
        page.mouse.wheel(0, 600)
    except Exception:
        pass
    time.sleep(5)

    if 'captcha' in page.url.lower():
        print("DEBUG: CAPTCHA page detected — skipping (لن نحاول حلها آليًا)")
        return None

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


# ================= المنطق الكامل =================
def resolve_image(link, page):
    # الطبقة 0: مباشر / Drive
    if "drive.google.com" in link:
        m = re.search(r"/d/([^/]+)", link)
        if m:
            return f"https://drive.google.com/uc?export=download&id={m.group(1)}"
    if link.lower().split('?')[0].endswith(('.jpg', '.jpeg', '.png', '.webp', '.gif')):
        return link

    # الطبقة 1: أمازون بدون فتح الصفحة أصلًا
    if 'amazon.' in link:
        url = amazon_direct_image(link)
        if url and verify_image(url):
            return url

    # الطبقة 2: requests (عادي ثم موبايل)
    url = get_image_via_requests(link)
    if url and verify_image(url):
        return url

    # الطبقة 3: Microlink (جلب من سيرفرات خارجية)
    url = get_image_via_microlink(link)
    if url and verify_image(url):
        return url

    # الطبقة 4: متصفح stealth
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


def is_known_placeholder(url):
    """الصور اللي اتكتبت غلط في تشغيلات سابقة (رسومات علي بابا tps)"""
    low = (url or '').lower()
    return 'alicdn' in low and ('tps-' in low or '/tps/' in low)


# تنضيف تلقائي: مسح خلايا G اللي فيها placeholder من تشغيلات سابقة
# حتى تتم إعادة سحب الصورة الصحيحة لها في نفس هذا التشغيل
cleanup_ranges = []
for idx in range(1, len(data)):
    g_val = cell(data[idx], 6).strip()
    if g_val and is_known_placeholder(g_val):
        cleanup_ranges.append(f"G{idx+1}")
        data[idx][6] = ''  # تعتبر فاضية محليًا فتدخل حلقة الاستخراج
if cleanup_ranges:
    worksheet.batch_clear(cleanup_ranges)
    print(f"🧹 Cleared {len(cleanup_ranges)} placeholder cells in G (alicdn tps) — will re-fetch them now.")

results = {}  # row_num -> img_url — الكتابة تتم في الآخر بعد فلترة المتكرر


def is_probable_url(link):
    """التحقق أن القيمة لينك فعلًا قبل أي معالجة — قيم زي كلمة 'string' تتخطى فورًا"""
    return bool(re.match(r'^https?://\S+\.\S+', link))


failed = []
skipped_not_url = 0

with sync_playwright() as p:
    browser = p.chromium.launch(
        headless=True,
        args=['--disable-blink-features=AutomationControlled', '--no-sandbox'],
    )
    context = browser.new_context(
        user_agent=DESKTOP_HEADERS["User-Agent"],
        locale="en-US",
        timezone_id="Asia/Dubai",
        viewport={"width": 1366, "height": 768},
        extra_http_headers={"Accept-Language": "en-US,en;q=0.9,ar;q=0.8"},
    )
    context.add_init_script(STEALTH_JS)
    page = context.new_page()

    for idx in range(1, len(data)):
        row = data[idx]
        img_g = cell(row, 6)
        link = cell(row, 7).strip()

        if img_g.strip() or not link:
            continue

        # إصلاح اللينكات المكتوبة بدون بروتوكول (www.site.com)
        if link.startswith('www.'):
            link = 'https://' + link

        # تخطي فوري لأي قيمة مش لينك (زي كلمة "string" الناتجة عن معادلة)
        # بدل إضاعة دقيقة كاملة على كل صف في تجربة الطبقات الأربعة
        if not is_probable_url(link):
            skipped_not_url += 1
            continue

        print(f"🌐 Row {idx+1}: {link[:60]}")
        try:
            img_url = resolve_image(link, page)
            if img_url:
                results[idx + 1] = img_url
                print(f"✅ Row {idx+1}: {img_url}")
            else:
                failed.append(idx + 1)
                print(f"❌ Row {idx+1}: no valid image found")
        except Exception as e:
            failed.append(idx + 1)
            print(f"⚠️ Row {idx+1} error: {e}")

        time.sleep(1)  # مهلة بين الصفوف — الطلبات المتلاحقة بسرعة أسرع طريق للحظر

    browser.close()

# ============ فلترة الـ placeholders قبل الكتابة ============
# لو نفس رابط الصورة رجع لأكتر من صفين مختلفين، فهو صورة صفحة حظر/خطأ
# وليس صورة منتج (منتجات مختلفة مستحيل يكون لها نفس الصورة بالظبط)
counts = Counter(results.values())
placeholder_urls = {u for u, c in counts.items() if c >= 3}
if placeholder_urls:
    print(f"\n🚫 استبعاد {len(placeholder_urls)} رابط صورة متكرر لعدة صفوف مختلفة (placeholder وليس منتج):")
    for u in placeholder_urls:
        print(f"   {u[:90]}")

final_updates = []
for row_num, url in sorted(results.items()):
    if url in placeholder_urls:
        failed.append(row_num)
    else:
        final_updates.append({'range': f'G{row_num}', 'values': [[url]]})

# الكتابة على دفعات من 50
for i in range(0, len(final_updates), 50):
    worksheet.batch_update(final_updates[i:i + 50])
print(f"\n💾 تم كتابة {len(final_updates)} صورة صحيحة في الشيت.")

if skipped_not_url:
    print(f"\n⏭️ تم تخطي {skipped_not_url} صف لأن عمود H فيها قيمة مش لينك (زي 'string') — "
          f"راجع المعادلة اللي بتملأ العمود ده وامسح الصفوف الزائدة من الشيت.")
if failed:
    print(f"\n❗ Rows with no image: {failed}")
print("🎉 Process Finished.")
