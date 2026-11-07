"""
Muratpaşa İlçe MEM İSG Bilgi Yönetim Portalı
Akıllı Dosya Depolama, Optimizasyon ve Arşivleme Sistemi (isg_storage.py)
"""

import os
import io
import re
import zipfile
import hashlib
from datetime import datetime
from typing import Optional, Tuple, Dict, Any
from PIL import Image, ImageOps

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOADS_ROOT = os.path.join(BASE_DIR, "static", "uploads")
SCHOOLS_STORAGE_DIR = os.path.join(UPLOADS_ROOT, "okullar")
ARCHIVE_STORAGE_DIR = os.path.join(UPLOADS_ROOT, "arsiv")

def slugify(text: str) -> str:
    """Türkçe karakterleri ve boşlukları güvenli klasör/dosya ismine çevirir."""
    if not text:
        return "kurum"
    tr_map = {
        'ı': 'i', 'İ': 'i', 'ğ': 'g', 'Ğ': 'g',
        'ü': 'u', 'Ü': 'u', 'ş': 's', 'Ş': 's',
        'ö': 'o', 'Ö': 'o', 'ç': 'c', 'Ç': 'c'
    }
    for tr_char, en_char in tr_map.items():
        text = text.replace(tr_char, en_char)
    text = re.sub(r'[^a-zA-Z0-9_\-\.]', '_', text.strip().lower())
    text = re.sub(r'_+', '_', text)
    return text.strip('_')

def get_category_folder(belge_turu: str) -> str:
    """Belge türüne göre alt klasör belirler."""
    b_lower = (belge_turu or "").lower()
    if any(k in b_lower for k in ["periyodik", "muayene", "kazan", "yangin", "tup", "asansor", "topraklama"]):
        return "periyodik_kontroller"
    elif any(k in b_lower for k in ["okulum_temiz", "hijyen", "enfeksiyon"]):
        return "okulum_temiz"
    elif any(k in b_lower for k in ["adep", "tatbikat", "tahliye", "yangin_tatbikati"]):
        return "adep_tatbikat"
    elif any(k in b_lower for k in ["mebbis", "rapor", "icmal", "excel"]):
        return "mebbis_raporlari"
    elif any(k in b_lower for k in ["foto", "gorsel", "resim", "etiket", "pano"]):
        return "saha_fotograflari"
    else:
        return "genel"

def optimize_image(image_bytes: bytes, max_dim: int = 1600, quality: int = 82) -> bytes:
    """
    Telefonlardan çekilen 8-15 MB'lık fotoğrafları:
    1. EXIF yönelimini (dikey çekim) düzeltir.
    2. Maksimum 1600px genişlik/yükseklikle orantılı olarak yeniden boyutlandırır.
    3. Kalite kaybı olmadan ~150-250 KB seviyesine sıkıştırır (%90+ tasarruf).
    """
    try:
        img = Image.open(io.BytesIO(image_bytes))
        # EXIF rotasyonunu düzelt
        try:
            img = ImageOps.exif_transpose(img)
        except Exception:
            pass

        # Yeniden boyutlandırma (Orantılı)
        if img.width > max_dim or img.height > max_dim:
            img.thumbnail((max_dim, max_dim), Image.Resampling.LANCZOS)

        # RGB dönüşümü (RGBA veya CMYK ise JPEG için RGB yap)
        if img.mode in ("RGBA", "P"):
            img = img.convert("RGB")
        elif img.mode != "RGB":
            img = img.convert("RGB")

        out_buf = io.BytesIO()
        img.save(out_buf, format="JPEG", quality=quality, optimize=True)
        return out_buf.getvalue()
    except Exception as e:
        # Hata olursa orijinal baytları döndür
        return image_bytes

def save_school_document(
    file_bytes: bytes,
    original_filename: str,
    school_info: Dict[str, Any],
    belge_turu: str = "genel"
) -> Dict[str, Any]:
    """
    Belgeyi hiyerarşik klasör yapısına optimize ederek kaydeder:
    static/uploads/okullar/{kurum_kodu}_{okul_slug}/{yil}/{kategori}/{timestamp}_{dosya_adi}
    """
    orig_size = len(file_bytes)
    orig_kb = round(orig_size / 1024, 1)

    # Okul klasör adı
    kurum_kodu = str(school_info.get("kurum_kodu") or "000000")
    kurum_adi = school_info.get("kurum_adi") or "okul"
    school_slug = f"{kurum_kodu}_{slugify(kurum_adi)[:30]}"

    # Yıl ve kategori
    current_year = datetime.now().strftime("%Y")
    category = get_category_folder(belge_turu)

    target_dir = os.path.join(SCHOOLS_STORAGE_DIR, school_slug, current_year, category)
    os.makedirs(target_dir, exist_ok=True)

    # Dosya uzantısı kontrolü
    base_name, ext = os.path.splitext(original_filename)
    ext_lower = ext.lower()

    is_image = ext_lower in [".jpg", ".jpeg", ".png", ".webp", ".heic"]
    final_bytes = file_bytes
    final_ext = ext_lower

    if is_image:
        final_bytes = optimize_image(file_bytes, max_dim=1600, quality=82)
        final_ext = ".jpg"

    saved_size = len(final_bytes)
    saved_kb = round(saved_size / 1024, 1)
    savings_pct = int(round((1 - (saved_size / orig_size)) * 100)) if orig_size > saved_size else 0

    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    clean_base = slugify(base_name)[:40]
    final_filename = f"{timestamp}_{clean_base}{final_ext}"
    final_abs_path = os.path.join(target_dir, final_filename)

    with open(final_abs_path, "wb") as f:
        f.write(final_bytes)

    # SHA256 Hash
    sha256_hash = hashlib.sha256(final_bytes).hexdigest()

    # Web URL yolu
    relative_url = f"/static/uploads/okullar/{school_slug}/{current_year}/{category}/{final_filename}"

    return {
        "relative_url": relative_url,
        "absolute_path": final_abs_path,
        "filename": final_filename,
        "original_filename": original_filename,
        "original_kb": orig_kb,
        "saved_kb": saved_kb,
        "savings_pct": savings_pct,
        "is_image": is_image,
        "sha256": sha256_hash,
        "category": category,
        "year": current_year
    }

def get_storage_statistics() -> Dict[str, Any]:
    """Depolama alanının genel durumunu ve tasarruf analizini hesaplar."""
    total_bytes = 0
    total_files = 0
    schools_set = set()
    years_set = set()

    if os.path.exists(SCHOOLS_STORAGE_DIR):
        for root, dirs, files in os.walk(SCHOOLS_STORAGE_DIR):
            for f in files:
                f_path = os.path.join(root, f)
                try:
                    f_size = os.path.getsize(f_path)
                    total_bytes += f_size
                    total_files += 1

                    parts = os.path.relpath(f_path, SCHOOLS_STORAGE_DIR).split(os.sep)
                    if len(parts) >= 1:
                        schools_set.add(parts[0])
                    if len(parts) >= 2 and parts[1].isdigit():
                        years_set.add(parts[1])
                except Exception:
                    pass

    total_mb = round(total_bytes / (1024 * 1024), 2)

    return {
        "total_files": total_files,
        "total_mb": total_mb,
        "schools_count": len(schools_set),
        "available_years": sorted(list(years_set), reverse=True)
    }

def export_annual_archive_zip(year: int) -> Tuple[str, str, float]:
    """
    Belirli bir yılın tüm okul evraklarını tek bir zip dosyasında paketler:
    static/uploads/arsiv/muratpasa_isg_arsiv_{year}.zip
    Döndürür: (relative_url, absolute_path, zip_boyut_mb)
    """
    os.makedirs(ARCHIVE_STORAGE_DIR, exist_ok=True)
    zip_filename = f"muratpasa_isg_arsiv_{year}.zip"
    zip_abs_path = os.path.join(ARCHIVE_STORAGE_DIR, zip_filename)

    year_str = str(year)
    files_added = 0

    with zipfile.ZipFile(zip_abs_path, "w", zipfile.ZIP_DEFLATED) as zipf:
        if os.path.exists(SCHOOLS_STORAGE_DIR):
            for root, dirs, files in os.walk(SCHOOLS_STORAGE_DIR):
                for f in files:
                    full_p = os.path.join(root, f)
                    rel_to_schools = os.path.relpath(full_p, SCHOOLS_STORAGE_DIR)
                    parts = rel_to_schools.split(os.sep)
                    # parts[0]: okul, parts[1]: yil, parts[2]: kategori
                    if len(parts) >= 2 and parts[1] == year_str:
                        zipf.write(full_p, arcname=rel_to_schools)
                        files_added += 1

    zip_size_mb = round(os.path.getsize(zip_abs_path) / (1024 * 1024), 2)
    relative_url = f"/static/uploads/arsiv/{zip_filename}"
    return relative_url, zip_abs_path, zip_size_mb
