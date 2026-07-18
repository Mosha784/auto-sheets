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
SCRAPERAPI_KEY = os.environ.get("SCRAPERAPI_KEY", "").strip()
SCRAPINGBEE_KEY = os.environ.get("SCRAPINGBEE_KEY", "").strip()

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
            r = requests.get(link, headers=headers, timeout=15, allow_redirects=True)
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
        r = requests.get('https://api.microlink.io/', params=params, headers=headers, timeout=25)
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


# ================= الطبقة 3ج: ScraperAPI (يتفعّل فقط لو المفتاح موجود) =================
def get_image_via_scraperapi(link):
    """
    ScraperAPI: بروكسيات دوّارة + تخطي أنظمة الحماية — يرجّع HTML الصفحة الحقيقية.
    مجاني: 1000 كريدت/شهر (+5000 أول أسبوع). يتفعّل فقط عند وجود SCRAPERAPI_KEY.
    """
    if not SCRAPERAPI_KEY:
        return None
    try:
        r = requests.get(
            'https://api.scraperapi.com/',
            params={'api_key': SCRAPERAPI_KEY, 'url': link},
            timeout=70,  # الخدمة نفسها ممكن تاخد وقت في المواقع الصعبة
        )
        if r.status_code != 200:
            print(f"DEBUG: scraperapi status {r.status_code}")
            return None
        url, how = extract_from_html(r.text, link)
        if url:
            print(f"DEBUG: [scraperapi/{how}] {url}")
        return url
    except requests.RequestException as e:
        print(f"DEBUG: scraperapi failed: {e}")
        return None


# ================= الطبقة 3د: ScrapingBee (يتفعّل فقط لو المفتاح موجود) =================
def get_image_via_scrapingbee(link):
    """
    ScrapingBee: نفس فكرة ScraperAPI — بروكسيات + تخطي حماية.
    مجاني: 1000 كريدت تجريبية لمرة واحدة. يتفعّل فقط عند وجود SCRAPINGBEE_KEY.
    render_js=false لتوفير الكريدت (الرندر بيستهلك 5 أضعاف).
    """
    if not SCRAPINGBEE_KEY:
        return None
    try:
        r = requests.get(
            'https://app.scrapingbee.com/api/v1/',
            params={'api_key': SCRAPINGBEE_KEY, 'url': link, 'render_js': 'false'},
            timeout=70,
        )
        if r.status_code != 200:
            print(f"DEBUG: scrapingbee status {r.status_code}")
            return None
        url, how = extract_from_html(r.text, link)
        if url:
            print(f"DEBUG: [scrapingbee/{how}] {url}")
        return url
    except requests.RequestException as e:
        print(f"DEBUG: scrapingbee failed: {e}")
        return None


# ================= الطبقة 3ب: Jina Reader =================
# مواقع معروف أنها بتحظر IPs الرانرز — نوفر وقت الطبقات اللي عمرها ما هتنجح معاها
BLOCKED_DOMAINS = ('alibaba.com', 'aliexpress.', '1688.com', 'taobao.com', 'tmall.com')

IMG_URL_RE = re.compile(r'https?://[^\s\)\]"\'<>]+\.(?:jpe?g|png|webp)[^\s\)\]"\'<>]*', re.IGNORECASE)
ALICDN_KF_RE = re.compile(r'https?://[a-z0-9.]*alicdn\.com/kf/[^\s\)\]"\'<>]+\.(?:jpe?g|png|webp)[^\s\)\]"\'<>]*', re.IGNORECASE)


def get_image_via_jina(link):
    """
    Jina Reader (r.jina.ai): خدمة مجانية تجلب الصفحة من سيرفراتها وترجعها نصًا —
    زي Microlink لكن بحدود استخدام أعلى بكثير. ممتازة لعلي بابا والمواقع المحظورة.
    """
    try:
        r = requests.get('https://r.jina.ai/' + link, timeout=40,
                         headers={'User-Agent': DESKTOP_HEADERS['User-Agent']})
        if r.status_code != 200:
            print(f"DEBUG: jina status {r.status_code}")
            return None
        text = r.text

        # الأولوية: صور منتجات علي بابا الحقيقية (مسار /kf/)
        m = ALICDN_KF_RE.search(text)
        if m:
            url = normalize_url(m.group(0), link)
            if looks_like_product_image(url, link):
                print(f"DEBUG: [jina/kf] {url}")
                return url

        # أي صورة منتج صالحة في محتوى الصفحة
        for m in IMG_URL_RE.finditer(text):
            url = normalize_url(m.group(0), link)
            if looks_like_product_image(url, link):
                print(f"DEBUG: [jina] {url}")
                return url
    except requests.RequestException as e:
        print(f"DEBUG: jina failed: {e}")
    return None



STEALTH_JS = """
Object.defineProperty(navigator, 'webdriver', { get: () => undefined });
window.chrome = window.chrome || { runtime: {} };
Object.defineProperty(navigator, 'languages', { get: () => ['en-US', 'en', 'ar'] });
Object.defineProperty(navigator, 'plugins', { get: () => [1, 2, 3, 4, 5] });
Object.defineProperty(navigator, 'platform', { get: () => 'Win32' });
"""


def get_image_via_playwright(link, page):
    page.goto(link, timeout=25000, wait_until="domcontentloaded")
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

    blocked = any(d in link for d in BLOCKED_DOMAINS)

    # الطبقة 2: requests — نتخطاها للمواقع المحظورة (بترجع صفحة حظر دايمًا = وقت ضايع)
    if not blocked:
        url = get_image_via_requests(link)
        if url and verify_image(url):
            return url

    # الطبقة 3أ: Jina Reader (مجاني بحدود عالية — الأنسب لعلي بابا)
    url = get_image_via_jina(link)
    if url and verify_image(url):
        return url

    # الطبقة 3ب: Microlink (احتياطي — الحد المجاني ~50 طلب/يوم)
    url = get_image_via_microlink(link)
    if url and verify_image(url):
        return url

    # الطبقة 3ج: ScraperAPI (لو المفتاح موجود — بروكسيات تتخطى الحظر)
    url = get_image_via_scraperapi(link)
    if url and verify_image(url):
        return url

    # الطبقة 3د: ScrapingBee (لو المفتاح موجود — احتياطي أخير قبل المتصفح)
    url = get_image_via_scrapingbee(link)
    if url and verify_image(url):
        return url

    # الطبقة 4: متصفح stealth — نتخطاها للمواقع المحظورة (نفس الـ IP المحظور = فشل مضمون)
    if not blocked:
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

# ================= إعدادات التشغيل =================
MAX_ROWS_PER_RUN = 60        # حد أقصى للصفوف المعالجة في التشغيلة الواحدة —
                             # الورك فلو بيشتغل كل 30 دقيقة فبيلحّق الباقي تدريجيًا
TIME_BUDGET_SECONDS = 9 * 60  # السكربت يوقف نفسه ويحفظ قبل ما الـ timeout يقطعه ويضيّع الشغل
FAILED_MARKER = "NO_IMAGE"    # علامة للصفوف الفاشلة حتى لا تُعاد كل تشغيل للأبد
                             # (امسح العلامة من الخلية يدويًا لو عايز تعيد المحاولة لصف معين)

script_start = time.time()


def is_probable_url(link):
    """التحقق أن القيمة لينك فعلًا قبل أي معالجة — قيم زي كلمة 'string' تتخطى فورًا"""
    return bool(re.match(r'^https?://\S+\.\S+', link))


# ============ آلية الكتابة التدريجية مع كشف الـ placeholder ============
# الكتابة كل 10 نتائج (لو السكربت اتقطع، اللي خلص محفوظ)
# ولو نفس الرابط ظهر لـ 3 صفوف مختلفة يُعتبر placeholder:
# يتشال من المعلق، واللي اتكتب منه قبل كده يُمسح في النهاية
url_counts = Counter()
blacklist = set()
pending = {}   # row -> url مستنية الكتابة
written = {}   # row -> url اتكتبت فعلًا
failed = []
skipped_not_url = 0


def flush(force=False):
    global pending
    if pending and (force or len(pending) >= 10):
        updates = [{'range': f'G{r}', 'values': [[u]]} for r, u in sorted(pending.items())]
        worksheet.batch_update(updates)
        written.update(pending)
        pending = {}


def register_result(row_num, url):
    url_counts[url] += 1
    if url_counts[url] >= 3 and url not in blacklist:
        blacklist.add(url)
        print(f"🚫 رابط متكرر لعدة منتجات مختلفة — يُعتبر placeholder ويُرفض: {url[:80]}")
        for r, u in list(pending.items()):
            if u == url:
                pending.pop(r)
                failed.append(r)
        # اللي اتكتب فعلًا بنفس الرابط هيتمسح في نهاية التشغيل
    if url in blacklist:
        failed.append(row_num)
    else:
        pending[row_num] = url
        flush()


stopped_reason = None

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

    processed = 0
    # من تحت لفوق: الصفوف الجديدة (آخر الشيت) لها الأولوية وتخلص الأول
    for idx in range(len(data) - 1, 0, -1):
        # وقف رشيق قبل نفاد الوقت أو تجاوز حد الصفوف — الباقي يتكمّل في التشغيلة الجاية
        if processed >= MAX_ROWS_PER_RUN:
            stopped_reason = f"وصلنا حد {MAX_ROWS_PER_RUN} صف للتشغيلة الواحدة"
            break
        if time.time() - script_start > TIME_BUDGET_SECONDS:
            stopped_reason = "قربت ميزانية الوقت تخلص"
            break

        row = data[idx]
        img_g = cell(row, 6)
        link = cell(row, 7).strip()

        if img_g.strip() or not link:
            continue

        # إصلاح اللينكات المكتوبة بدون بروتوكول (www.site.com)
        if link.startswith('www.'):
            link = 'https://' + link

        # تخطي فوري لأي قيمة مش لينك (زي كلمة "string" الناتجة عن معادلة)
        if not is_probable_url(link):
            skipped_not_url += 1
            continue

        processed += 1
        print(f"🌐 Row {idx+1}: {link[:60]}")
        try:
            img_url = resolve_image(link, page)
            if img_url:
                register_result(idx + 1, img_url)
                print(f"✅ Row {idx+1}: {img_url}")
            else:
                failed.append(idx + 1)
                print(f"❌ Row {idx+1}: no valid image found")
        except Exception as e:
            failed.append(idx + 1)
            print(f"⚠️ Row {idx+1} error: {e}")

        time.sleep(1)  # مهلة بين الصفوف — الطلبات المتلاحقة بسرعة أسرع طريق للحظر

    browser.close()

# ============ الحفظ النهائي والتنضيف ============
flush(force=True)

# مسح أي صور placeholder اتكتبت قبل ما نكتشف إنها متكررة
bad_written = [f"G{r}" for r, u in written.items() if u in blacklist]
if bad_written:
    worksheet.batch_clear(bad_written)
    for r, u in list(written.items()):
        if u in blacklist:
            written.pop(r)
            failed.append(r)

# علامة NO_IMAGE للصفوف الفاشلة — حتى لا يعيد السكربت محاولتها كل 30 دقيقة للأبد
failed = sorted(set(failed))
if failed and FAILED_MARKER:
    marker_updates = [{'range': f'G{r}', 'values': [[FAILED_MARKER]]} for r in failed]
    for i in range(0, len(marker_updates), 50):
        worksheet.batch_update(marker_updates[i:i + 50])

good_count = len(written)
print(f"\n💾 تم كتابة {good_count} صورة صحيحة في الشيت.")
if failed:
    print(f"🏷️ تم وضع علامة {FAILED_MARKER} على {len(failed)} صف فاشل: {failed}")
    print("   (امسح العلامة من خلية G يدويًا لأي صف عايز تعيد محاولته)")
if skipped_not_url:
    print(f"⏭️ تم تخطي {skipped_not_url} صف لأن عمود H فيها قيمة مش لينك (زي 'string') — "
          f"راجع المعادلة اللي بتملأ العمود ده وامسح الصفوف الزائدة من الشيت.")
if stopped_reason:
    print(f"⏸️ توقف السكربت مبكرًا ({stopped_reason}) — الصفوف المتبقية ستُعالج في التشغيلة القادمة تلقائيًا.")
print("🎉 Process Finished.")
