"""
PDF to Excel Converter
======================
Bu Python aracı, belirli bir klasördeki PDF dosyalarındaki tabloları otomatik olarak tespit edip
tek bir Excel dosyasına (her PDF için ayrı sayfa olacak şekilde) aktarır.

Özellikler:
-----------
- Klasördeki tüm PDF'leri tarar
- Tüm sayfalardaki tabloları algılar
- Her PDF dosyasını Excel’de ayrı bir sayfaya (sheet) yazar
- Tablosu olmayan PDF'leri atlar
- Hata durumlarını kullanıcıya bildirir

Kullanım:
---------
python main.py
"""

import tabula
import pandas as pd
import glob
import os


def pdfs_to_excel(pdf_folder: str = "PDF", excel_output: str = "excel.xlsx"):
    """Belirtilen klasördeki PDF dosyalarındaki tabloları Excel'e aktarır."""
    pdf_files = glob.glob(os.path.join(pdf_folder, "*.pdf"))

    if not pdf_files:
        print(f"⚠️ '{pdf_folder}' klasöründe PDF bulunamadı.")
        return

    with pd.ExcelWriter(excel_output, engine="openpyxl") as writer:
        for pdf_file in pdf_files:
            sheet_name = os.path.splitext(os.path.basename(pdf_file))[0]

            try:
                # Tüm sayfalardaki tabloları oku
                tables = tabula.read_pdf(pdf_file, pages="all", multiple_tables=True)

                # Hiç tablo yoksa atla
                if not tables:
                    print(f"⚠️ {pdf_file} içinde tablo bulunamadı, atlandı.")
                    continue

                # Tüm tabloları birleştir
                combined = pd.concat(tables, ignore_index=True)

                # Excel sayfası ismini 31 karakterle sınırla (Excel limiti)
                combined.to_excel(writer, sheet_name=sheet_name[:31], index=False)

                print(f"✅ {pdf_file} işlendi.")
            except Exception as e:
                print(f"❌ {pdf_file} okunamadı: {e}")

    print(f"\n🎉 Tüm PDF'ler '{excel_output}' dosyasına aktarıldı.")


if __name__ == "__main__":
    pdfs_to_excel()
