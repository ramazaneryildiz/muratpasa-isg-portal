import os
import sys
import xlrd

sys.stdout.reconfigure(encoding='utf-8')

RR_DIR = r"C:\iş_güvenliği\risk_raporları"

def main():
    print("=" * 65)
    print(" 6 RİSK RAPORU DOSYASI PARSE EDİLİYOR")
    print("=" * 65)
    for f in os.listdir(RR_DIR):
        if not f.endswith(".xls"):
            continue
        fp = os.path.join(RR_DIR, f)
        try:
            book = xlrd.open_workbook(fp)
            sheet = book.sheet_by_index(0)
            print(f"\nDosya: {f} (Satır: {sheet.nrows}, Sütun: {sheet.ncols})")
            # İlk 5 satırı yazdır
            kurum_bilgisi = ""
            for r in range(min(5, sheet.nrows)):
                row_vals = [str(sheet.cell_value(r, c)) for c in range(min(10, sheet.ncols))]
                row_text = " | ".join(row_vals)
                if "KURUM" in row_text.upper() or "KODU" in row_text.upper() or "OKUL" in row_text.upper():
                    kurum_bilgisi = row_text
                print(f"  [{r+1}] {row_text[:120]}")
            if kurum_bilgisi:
                print(f"  -> Kurum Bilgisi: {kurum_bilgisi}")
        except Exception as e:
            print(f"Hata {f}: {e}")

if __name__ == "__main__":
    main()
