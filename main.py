import pandas as pd
import requests
import os
import re
from io import BytesIO
from PIL import Image
import unicodedata

# 1. Hàm SEO: Chuyển tiếng Việt chuẩn & Gọt độ dài (30-50 ký tự)
def slugify_short(text, max_chars=50):
    if not text or str(text).lower() == 'nan' or str(text).strip() == '':
        return ""
    text = str(text).lower()
    
    replace_dict = {
        'đ': 'd', 'Đ': 'd', 'á': 'a', 'à': 'a', 'ả': 'a', 'ã': 'a', 'ạ': 'a',
        'ă': 'a', 'ắ': 'a', 'ằ': 'a', 'ẳ': 'a', 'ẵ': 'a', 'ặ': 'a',
        'â': 'a', 'ấ': 'a', 'ầ': 'a', 'ẩ': 'a', 'ẫ': 'a', 'ậ': 'a',
        'é': 'e', 'è': 'e', 'ẻ': 'e', 'ẽ': 'e', 'ẹ': 'e',
        'ê': 'e', 'ế': 'e', 'ề': 'e', 'ể': 'e', 'ễ': 'e', 'ệ': 'e',
        'í': 'i', 'ì': 'i', 'ỉ': 'i', 'ĩ': 'i', 'ị': 'i',
        'ó': 'o', 'ò': 'o', 'ỏ': 'o', 'õ': 'o', 'ọ': 'o',
        'ô': 'o', 'ố': 'o', 'ồ': 'o', 'ổ': 'o', 'ỗ': 'o', 'ộ': 'o',
        'ơ': 'o', 'ớ': 'o', 'ờ': 'o', 'ở': 'o', 'ỡ': 'o', 'ợ': 'o',
        'ú': 'u', 'ù': 'u', 'ủ': 'u', 'ũ': 'u', 'ụ': 'u',
        'ư': 'u', 'ứ': 'u', 'ừ': 'u', 'ử': 'u', 'ữ': 'u', 'ự': 'u',
        'ý': 'y', 'ỳ': 'y', 'ỷ': 'y', 'ỹ': 'y', 'ỵ': 'y'
    }
    for char, rep in replace_dict.items():
        text = text.replace(char, rep)
    
    text = unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('utf-8')
    text = re.sub(r'[^a-z0-9]+', '-', text)
    text = re.sub(r'-+', '-', text).strip('-')
    
    if len(text) > max_chars:
        text = text[:max_chars].strip('-')
        
    return text

# --- CẤU HÌNH ---
file_excel = 'data.xlsx'
brand_tag = "epharmacy" 
headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) Chrome/91.0.4472.124 Safari/537.36'}

try:
    df = pd.read_excel(file_excel)
    print(f"--- Đã kết nối file {file_excel} ---")
except Exception as e:
    print(f"❌ Lỗi: {e}")
    exit()

for index, row in df.iterrows():
    val_id = str(row.iloc[0]).strip()
    val_name = str(row.iloc[1]).strip()
    
    if (not val_id or val_id.lower() in ['nan', 'sku', 'id']) and (not val_name or val_name.lower() in ['nan', 'tên sản phẩm']):
        continue

    clean_id = re.sub(r'[^a-zA-Z0-9]', '-', val_id).strip('-') if val_id.lower() != 'nan' else ""
    clean_name_short = slugify_short(val_name, max_chars=50) 
    
    folder_full_name = f"{clean_id}-{clean_name_short}" if clean_id else clean_name_short
    file_base_name = f"{clean_name_short if clean_name_short else clean_id}-{brand_tag}"
    
    if not os.path.exists(folder_full_name):
        os.makedirs(folder_full_name)
    
    print(f"\n📂 Thư mục: {folder_full_name}")

    img_count = 0
    for i in range(2, len(df.columns)):
        url = row.iloc[i]
        if pd.notna(url) and str(url).startswith('http'):
            current_count = img_count + 1
            save_name = f"{file_base_name}-{current_count}.jpg"
            save_path = os.path.join(folder_full_name, save_name)
            
            try:
                response = requests.get(url, headers=headers, timeout=15)
                if response.status_code == 200:
                    img = Image.open(BytesIO(response.content))
                    if img.mode in ("RGBA", "P"):
                        img = img.convert("RGB")
                    img.save(save_path, "JPEG", quality=90, optimize=True)
                    print(f"   ✅ OK: {save_name}")
                    img_count += 1
            except Exception as e:
                print(f"   ❌ Lỗi link hình cột {i+1}")

    if img_count == 0 and os.path.exists(folder_full_name):
        os.rmdir(folder_full_name)

print("\n" + "="*50)
print(f"🎯 XONG! Tool đã nạp code chuẩn SEO ePharmacy.")
print("="*50)
