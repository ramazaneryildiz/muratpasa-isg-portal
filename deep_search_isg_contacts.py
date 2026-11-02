import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

ISG_DIR = r"C:\iş_güvenliği"

def search_text_in_files():
    print("=" * 65)
    print(" C:\\iş_güvenliği İÇİNDEKİ TÜM DOSYALAR TARANIYOR...")
    print("=" * 65)
    
    matches = []
    for root, dirs, files in os.walk(ISG_DIR):
        # gradle ve .gradle gibi derleme klasörlerini atla
        if any(x in root for x in [".gradle", "gradle", "app\\build", ".kotlin"]):
            continue
        for f in files:
            ext = os.path.splitext(f)[1].lower()
            fp = os.path.join(root, f)
            
            # Excel dosyaları
            if ext in [".xlsx", ".xlsm"]:
                try:
                    wb = openpyxl.load_workbook(fp, data_only=True, read_only=True)
                    for sname in wb.sheetnames:
                        ws = wb[sname]
                        for row in ws.iter_rows(values_only=True):
                            row_str = " ".join([str(v) for v in row if v is not None])
                            if any(k in row_str.lower() for k in ["0505", "053", "054", "055", "telefon", "cep tel", "müdür yrd"]):
                                matches.append((f, sname, row_str[:120]))
                                break
                except Exception:
                    pass

            # Text / js / py / csv
            elif ext in [".txt", ".csv", ".js", ".py", ".json", ".md"]:
                try:
                    with open(fp, "r", encoding="utf-8", errors="ignore") as tf:
                        for line in tf:
                            if any(k in line.lower() for k in ["0505", "053", "054", "055", "telefon", "cep tel"]):
                                matches.append((f, "text", line.strip()[:120]))
                                break
                except Exception:
                    pass

    print(f"Toplam tespit edilen telefon / iletişim eşleşmesi: {len(matches)}")
    for fn, src, sample in matches:
        print(f"  * [{fn}] ({src}): {sample}")

if __name__ == "__main__":
    search_text_in_files()
