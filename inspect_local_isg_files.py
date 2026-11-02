import os
import sys
import openpyxl

sys.stdout.reconfigure(encoding='utf-8')

ISG_DIR = r"C:\iş_güvenliği"

def inspect_excel(path):
    print("=" * 60)
    print(f"DOSYA: {path}")
    print("=" * 60)
    try:
        wb = openpyxl.load_workbook(path, data_only=True)
        print(f"Sayfalar: {wb.sheetnames}")
        for sheetname in wb.sheetnames:
            ws = wb[sheetname]
            print(f"\n--- Sayfa: {sheetname} (Satır: {ws.max_row}, Sütun: {ws.max_column}) ---")
            for r in range(1, min(10, ws.max_row + 1)):
                row_vals = [ws.cell(r, c).value for c in range(1, min(15, ws.max_column + 1))]
                print(f"  [{r}] {row_vals}")
    except Exception as e:
        print(f"Hata: {e}")

def main():
    excel_files = [
        os.path.join(ISG_DIR, "okul listesi.xlsx"),
        os.path.join(ISG_DIR, "ANTALYA MEM ÇİZELGELER TABLO 2-3-4(muratpaşa).xlsx"),
        os.path.join(ISG_DIR, "ÖNCELİKLİ OKULLAR LİSTESİ ANTALYA İL GENELİ2026.xlsx"),
        os.path.join(ISG_DIR, "ÇİZELGELER İLÇELER.xlsx")
    ]
    for ef in excel_files:
        if os.path.exists(ef):
            inspect_excel(ef)
        else:
            print(f"Bulunamadı: {ef}")

if __name__ == "__main__":
    main()
