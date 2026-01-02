import json
import gspread
import requests
from bs4 import BeautifulSoup
from oauth2client.service_account import ServiceAccountCredentials
import time
import re

# 1. إعداد الوصول لجوجل شيت
with open('service_account.json') as f:
    service_account_info = json.load(f)

scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_dict(service_account_info, scope)
client = gspread.authorize(creds)
sheet = client.open_by_url("https://docs.google.com/spreadsheets/d/1YFdOAR04ORhSbs38KfZPEdJQouX-bcH6exWjI06zvec/edit")
worksheet = sheet.worksheet("Missing In Form")

def clean_image_url(url):
    """تنظيف الرابط للحصول على الصورة الأصلية عالية الجودة"""
    if not url: return None
    # إزالة لاحقات الحجم مثل _300x300.jpg
    url = re.sub(r'_\d+x\d+.*$', '', url)
    if url.startswith('//'): url = "https:" + url
    return url

def get_image_statically(link):
    """استخراج الصورة بدون متصفح (أسرع وأدق لتجنب البلوك)"""
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36",
        "Accept-Language": "en-US,en;q=0.9"
    }
    try:
        response = requests.get(link, headers=headers, timeout=15)
        if response.status_code != 200: return None
        
        soup = BeautifulSoup(response.text, 'html.parser')
        
        # 1. محاولة og:image (الأكثر دقة)
        meta_og = soup.find("meta", property="og:image")
        if meta_og and meta_og.get("content"):
            content = meta_og["content"]
            # استبعاد اللوجو (اللوجو غالباً PNG وصورة المنتج JPG)
            if ".jpg" in content.lower() and "tps-" not in content:
                return clean_image_url(content)
        
        # 2. البحث في سكريبتات الصفحة عن روابط الصور
        scripts = soup.find_all("script")
        for script in scripts:
            if script.string and "imageConfig" in script.string:
                # استخراج أول رابط صورة منتج ينتهي بـ .jpg
                match = re.search(r'(https:[^"]+\.jpg)', script.string)
                if match:
                    return clean_image_url(match.group(1).replace('\\u002F', '/'))
                    
    except Exception as e:
        print(f"Error fetching {link}: {e}")
    return None

# --- التشغيل الأساسي ---
print("🚀 Starting High-Speed Extraction Process...")
data = worksheet.get_all_values()

# تحديد الأعمدة (G هو 7 و H هو 8)
COL_IMAGE = 7
COL_LINK = 8

for idx in range(1, len(data)):
    row = data[idx]
    # التأكد من وجود بيانات كافية في الصف
    img_val = row[COL_IMAGE-1] if len(row) >= COL_IMAGE else ''
    link = row[COL_LINK-1] if len(row) >= COL_LINK else ''
    
    # إذا كانت خلية الصورة فارغة وهناك رابط منتج
    if (not img_val or not img_val.strip()) and link and link.strip():
        print(f"🌐 Row {idx+1}: Processing {link[:50]}...")
        
        # محاولة الاستخراج السريع
        img_url = get_image_statically(link)
        
        if img_url:
            worksheet.update_cell(idx+1, COL_IMAGE, img_url)
            print(f"✅ Success: {img_url}")
            # تأخير بسيط لتجنب الحظر من جوجل شيت
            time.sleep(1)
        else:
            print(f"❌ Failed to find product image for row {idx+1}")

print("🎉 Task Completed Successfully.")
