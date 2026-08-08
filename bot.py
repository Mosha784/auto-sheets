# -*- coding: utf-8 -*-
import os, re, time, random, requests
import gspread
from playwright.sync_api import sync_playwright

UAH = {"User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/126.0 Safari/537.36"}

SHEET_URL = os.environ["SHEET_URL"]
open("sa.json", "w").write(os.environ["SA_JSON"])
USD_AED = float(os.environ.get("USD_AED", "3.67"))
CNY_AED = float(os.environ.get("CNY_AED", "0.51"))
PROXY = os.environ.get("PROXY_SERVER", "").strip()
SCRAPER_KEY = os.environ.get("SCRAPERAPI_KEY", "").strip()
BUDGET = int(os.environ.get("TIME_BUDGET_MIN", "30")) * 60

gc = gspread.service_account("sa.json")
sh = gc.open_by_url(SHEET_URL)
ws = sh.get_worksheet(0)
grid = ws.get_all_values()
headers = [str(h).strip().lower() for h in grid[0]]

def col(*parts):
    for i, h in enumerate(headers):
        if all(p in h for p in parts):
            return i
    return -1

def cl(n):
    s = ""
    n += 1
    while n:
        n, r = divmod(n - 1, 26)
        s = chr(65 + r) + s
    return s

def num(v):
    try:
        return float(str(v).replace(",", ""))
    except Exception:
        return None

def cellv(i, c):
    try:
        return grid[i + 1][c]
    except Exception:
        return ""

for name in ["Alibaba Landed (AED)", "16888 Landed (AED)"]:
    if not any("landed" in h for h in headers):
        ws.update_cell(1, len(headers) + 1, name)
        headers.append(name.lower())
        for r in grid:
            r.append("")

C = {"img": col("product link"), "ali_price": col("alibaba", "price"), "cn_price": col("1688", "price"), "L": col("length"), "W": col("width"), "H": col("height"), "WT": col("weight"), "items": col("items per carton"), "cbm": col("cbm"), "ali_link": col("alibaba", "link"), "ali_photo": col("alibaba", "photo"), "cn_link": col("1688", "link"), "cn_photo": col("1688", "product"), "land_ali": col("alibaba", "landed"), "land_cn": col("1688", "landed")}

try:
    meta = sh.worksheet("meta")
except gspread.WorksheetNotFound:
    meta = sh.add_worksheet("meta", rows=5, cols=2)

def progress():
    v = meta.acell("A1").value
    return int(v) if v and str(v).isdigit() else 0

def set_progress(n):
    meta.update_acell("A1", str(n))

def walk(node, out, depth=0):
    if depth > 8 or len(out) > 300:
        return
    if isinstance(node, dict):
        lk = {k.lower(): k for k in node}
        if any("price" in k for k in lk) and any(any(t in k for t in ("subject", "title", "name")) for k in lk):
            it = extract(node, lk)
            if it:
                out.append(it)
        for v in node.values():
            walk(v, out, depth + 1)
    elif isinstance(node, list):
        for v in node:
            walk(v, out, depth + 1)

def extract(d, lk):
    def find(subs):
        for k in lk:
            if any(s in k for s in subs):
                v = d[lk[k]]
                if isinstance(v, (str, int, float)):
                    return v
        return None
    title = find(("subject", "title", "name"))
    url = find(("detailurl", "producturl", "offerurl", "url", "link", "href"))
    price = find(("price",))
    if not title or not url:
        return None
    url = str(url)
    if url.startswith("//"):
        url = "https:" + url
    m = re.search(r"\d+(?:\.\d+)?", str(price or ""))
    return {"title": str(title)[:90], "url": url[:300], "price": float(m.group()) if m else None}

def make_hook(bucket):
    def f(resp):
        try:
            if resp.status != 200 or "json" not in (resp.headers.get("content-type") or ""):
                return
            walk(resp.json(), bucket)
        except Exception:
            pass
    return f

def dedupe(items):
    seen, out = set(), []
    for it in items:
        u = it["url"].split("#")[0].split("?")[0]
        if u in seen:
            continue
        seen.add(u)
        out.append(it)
    return out

def dom(page, hint):
    out = []
    try:
        els = page.eval_on_selector_all("a[href]", "els => els.map(e => ({h: e.href, t: (e.innerText||'').trim().replace(/\\s+/g,' ').slice(0,90)}))")
        seen = set()
        for a in els:
            if hint in a["h"] and len(a["t"]) > 6 and a["h"] not in seen:
                seen.add(a["h"])
                m = re.search(r"\d+(?:\.\d+)?", a["t"])
                out.append({"title": a["t"], "url": a["h"], "price": float(m.group()) if m else None})
    except Exception:
        pass
    return out

def upload(page, img_path, sels):
    try:
        inp = page.query_selector("input[type=file]")
        if inp:
            inp.set_input_files(img_path)
            return True
    except Exception:
        pass
    try:
        page.evaluate("() => { const els = Array.from(document.querySelectorAll('button, div, span, i, img, a, svg')); const el = els.find(e => { const c = (e.className && e.className.baseVal !== undefined) ? e.className.baseVal : (e.className || ''); const s = (c + ' ' + (e.id || '') + ' ' + (e.getAttribute('aria-label') || '') + ' ' + (e.getAttribute('title') || '')).toLowerCase(); return /camera|image[- ]?search|img[- ]?search|photo|picture/.test(s); }); if (el) el.click(); }")
        page.wait_for_timeout(2500)
        inp = page.query_selector("input[type=file]")
        if inp:
            inp.set_input_files(img_path)
            return True
    except Exception as e:
        print("generic click err:", str(e)[:80])
    try:
        handle = page.evaluate_handle("() => { const scope = document.querySelector('form') || document.querySelector('[class*=search]') || document.querySelector('header') || document.body; return Array.from(scope.querySelectorAll('button, [role=button], div, span, i, img, svg, a')).filter(e => { const r = e.getBoundingClientRect(); return r.width > 8 && r.width < 90 && r.height > 8 && r.height < 90; }).slice(0, 15); }")
        count = page.evaluate("els => els.length", handle)
        print("brute candidates:", count)
        for idx in range(count):
            try:
                with page.expect_file_chooser(timeout=2000) as fc:
                    page.evaluate("(els, i) => els[i].click()", [handle, idx])
                fc.value.set_files(img_path)
                print("brute force hit at", idx)
                return True
            except Exception:
                continue
    except Exception as e:
        print("brute err:", str(e)[:80])
    try:
        h = page.evaluate("() => { const f = document.querySelector('form'); return f ? f.outerHTML : document.body.innerHTML.slice(0, 2000); }")
        print("FORM HTML:", h.replace("\n", " ")[:2000])
    except Exception:
        pass
    return False

def download(url, path):
    for u in (url, "https://wsrv.nl/?url=" + requests.utils.quote(url, safe="")):
        try:
            r = requests.get(u, headers=UAH, timeout=30)
            if r.ok and len(r.content) > 1000:
                open(path, "wb").write(r.content)
                return True
        except Exception:
            pass
    return False

def blocked(title):
    return any(k in title.lower() for k in ("captcha", "verify", "verification", "access denied", "punish"))

def image_search(ctx, img_path, site):
    page = ctx.new_page()
    bucket = []
    page.on("response", make_hook(bucket))
    cands = []
    try:
        if site == "alibaba":
            page.goto("https://www.alibaba.com/", wait_until="domcontentloaded", timeout=60000)
            page.wait_for_timeout(3000)
            if blocked(page.title()):
                print("alibaba: blocked page")
            ok = upload(page, img_path, ["[class*='camera']", "[class*='image-search']", "[class*='img-search']", "form [class*='icon']"])
            print("alibaba upload ok:", ok)
            if ok:
                page.wait_for_timeout(9000)
            cands = dedupe(bucket) or dom(page, "product-detail")
        elif site == "1688":
            page.goto("https://s.1688.com/youyuan/index.htm?tab=imageSearch", wait_until="domcontentloaded", timeout=30000)
            page.wait_for_timeout(3000)
            if "login" in page.url or "passport" in page.url:
                print("1688: needs login - skipped")
                page.close()
                return []
            ok = upload(page, img_path, ["[class*='upload']", "[class*='camera']", "[class*='photo']"])
            print("1688 upload ok:", ok)
            if ok:
                page.wait_for_timeout(9000)
            cands = dedupe(bucket) or dom(page, "detail.1688.com")
    except Exception as e:
        print(site, "search err:", str(e)[:100])
    page.close()
    return cands[:5]

def scrape_alibaba(ctx, url):
    info = {}
    page = ctx.new_page()
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=60000)
        page.wait_for_timeout(5000)
        t = page.inner_text("body")
        m = re.search(r"Single package size:\s*([\d.]+)\s*[Xx×]\s*([\d.]+)\s*[Xx×]\s*([\d.]+)", t)
        if m:
            info["dims"] = [float(m.group(i)) for i in (1, 2, 3)]
        m = re.search(r"Single gross weight:?\s*([\d.,]+)\s*kg", t, re.I)
        if m:
            info["weight"] = float(m.group(1).replace(",", ""))
        prices = [float(p.replace(",", "")) for p in re.findall(r"(?:US\s*)?\$\s*([\d,]+(?:\.\d+)?)", t)]
        prices = [p for p in prices if 0.05 <= p <= 100000]
        if prices:
            info["price"] = min(prices)
    except Exception as e:
        print("alibaba page err:", str(e)[:100])
    page.close()
    return info

def scrape_1688(ctx, url):
    info = {}
    page = ctx.new_page()
    try:
        page.goto(url, wait_until="domcontentloaded", timeout=30000)
        page.wait_for_timeout(5000)
        t = page.inner_text("body")
        prices = [float(p) for p in re.findall(r"¥\s*(\d+(?:\.\d+)?)", t)]
        prices = [p for p in prices if 0.05 <= p <= 100000]
        if prices:
            info["price"] = min(prices)
        for kw in ["包装尺寸", "外箱尺寸", "单件尺寸", "尺寸"]:
            i = t.find(kw)
            if i >= 0:
                m = re.search(r"([\d.]+)\s*[Xx×*]\s*([\d.]+)\s*[Xx×*]\s*([\d.]+)", t[i:i+200])
                if m:
                    info["dims"] = [float(m.group(k)) for k in (1, 2, 3)]
                    break
    except Exception as e:
        print("1688 page err:", str(e)[:100])
    page.close()
    return info

def put(updates, r, c, v):
    if c >= 0:
        updates.append((r, c, v))

def flush(updates):
    if not updates:
        return
    try:
        data = [{"range": "'%s'!%s%d" % (ws.title, cl(c), r), "values": [[v]]} for r, c, v in updates]
        sh.values_batch_update({"data": data})
    except Exception as e:
        print("batch err:", str(e)[:100])
        for r, c, v in updates:
            try:
                ws.update_cell(r, c, v)
            except Exception:
                pass

def process(i, img, updates, ctx):
    r = i + 2
    path = "/tmp/img.jpg"
    if not download(img, path):
        print("image download failed")
        return
    dims = None
    ali_aed = None
    cn_aed = None
    ali = image_search(ctx, path, "alibaba")
    if ali:
        put(updates, r, C["ali_link"], ali[0]["url"])
        info = scrape_alibaba(ctx, ali[0]["url"])
        if info.get("price"):
            ali_aed = round(info["price"] * USD_AED, 2)
            put(updates, r, C["ali_price"], ali_aed)
        if info.get("dims"):
            dims = info["dims"]
            put(updates, r, C["L"], dims[0])
            put(updates, r, C["W"], dims[1])
            put(updates, r, C["H"], dims[2])
        if info.get("weight"):
            put(updates, r, C["WT"], info["weight"])
    cn = image_search(ctx, path, "1688")
    if cn:
        put(updates, r, C["cn_link"], cn[0]["url"])
        info = scrape_1688(ctx, cn[0]["url"])
        if info.get("price"):
            cn_aed = round(info["price"] * CNY_AED, 2)
            put(updates, r, C["cn_price"], cn_aed)
        if not dims and info.get("dims"):
            dims = info["dims"]
            put(updates, r, C["L"], dims[0])
            put(updates, r, C["W"], dims[1])
            put(updates, r, C["H"], dims[2])
    L = dims[0] if dims else num(cellv(i, C["L"]))
    W = dims[1] if dims else num(cellv(i, C["W"]))
    H = dims[2] if dims else num(cellv(i, C["H"]))
    items_per = num(cellv(i, C["items"])) or 1
    cbm_cost = num(cellv(i, C["cbm"])) or 0
    if L and W and H and cbm_cost:
        ship = (L * W * H / 1e6 / items_per) * cbm_cost
        if ali_aed:
            put(updates, r, C["land_ali"], round(ali_aed + ship, 2))
        if cn_aed:
            put(updates, r, C["land_cn"], round(cn_aed + ship, 2))

def main():
    t0 = time.time()
    total = len(grid) - 1
    if C["img"] < 0:
        print("no Product Link column!")
        return
    start = progress()
    if start >= total:
        start = 0
    updates = []
    i = start
    with sync_playwright() as p:
        browser = p.chromium.launch(headless=False, args=["--disable-blink-features=AutomationControlled"])
        ctx_kw = {"viewport": {"width": 1366, "height": 900}, "locale": "zh-CN", "user_agent": UAH["User-Agent"]}
        if PROXY:
            ctx_kw["proxy"] = {"server": PROXY}
        elif SCRAPER_KEY:
            ctx_kw["proxy"] = {"server": "http://proxy.scraperapi.com:8001", "username": "scraperapi", "password": SCRAPER_KEY}
            print("using ScraperAPI proxy")
        ctx = browser.new_context(**ctx_kw)
        while i < total and time.time() - t0 < BUDGET:
            img = str(cellv(i, C["img"])).strip()
            print("=== row", i + 2, "===")
            if img:
                try:
                    process(i, img, updates, ctx)
                except Exception as e:
                    print("row err:", str(e)[:100])
            i += 1
            set_progress(i)
            time.sleep(random.uniform(2, 4))
        ctx.close()
        browser.close()
    flush(updates)
    print("done row", i, "of", total)

main()
