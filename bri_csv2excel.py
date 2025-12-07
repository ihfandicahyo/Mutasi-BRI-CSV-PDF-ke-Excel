import pandas as pd
import glob
import os
from openpyxl.styles import NamedStyle

def format_jam(x):
    """Mengubah angka 124141 menjadi 12:41:41"""
    s = str(x).split('.')[0].zfill(6) 
    return f"{s[:2]}:{s[2:4]}:{s[4:]}"

def format_tanggal_indo(dt):
    """Mengubah datetime menjadi string format Indonesia (07 Oktober 2025)"""
    if pd.isnull(dt):
        return ""
    
    bulan_indo = {
        1: 'Januari', 2: 'Februari', 3: 'Maret', 4: 'April', 5: 'Mei', 6: 'Juni',
        7: 'Juli', 8: 'Agustus', 9: 'September', 10: 'Oktober', 11: 'November', 12: 'Desember'
    }
    
    # Format: dd NamaBulan yyyy (contoh: 07 Oktober 2025)
    return f"{dt.day:02d} {bulan_indo[dt.month]} {dt.year}"

def auto_adjust_excel_width(worksheet, df):
    """Mengatur lebar kolom otomatis"""
    from openpyxl.utils import get_column_letter
    for i, col in enumerate(df.columns):
        # Cari panjang maksimum data (header vs isi)
        max_len = max(
            df[col].astype(str).map(len).max(),
            len(str(col))
        )
        # Set lebar kolom
        worksheet.column_dimensions[get_column_letter(i + 1)].width = max_len + 2

# --- MULAI PROSES ---
csv_files = glob.glob('*.csv')

if not csv_files:
    print("Tidak ditemukan file CSV.")
else:
    print(f"Ditemukan {len(csv_files)} file CSV. Memproses...")

    for file_csv in csv_files:
        try:
            print(f"Mengolah: {file_csv} ...")
            
            # 1. Baca CSV (Pastikan NOREK dibaca sebagai string)
            df = pd.read_csv(file_csv, dtype={'NOREK': str})
            
            # --- TRANSFORMASI DATA ---
            
            # A. Perbaiki NOREK (Kolom B) -> Pastikan String
            df['NOREK'] = df['NOREK'].astype(str)
            
            # B. Perbaiki Tanggal (Kolom D) -> Format Indonesia
            # Convert ke datetime dulu
            df['TGL_EFEKTIF'] = pd.to_datetime(df['TGL_EFEKTIF'])
            # Apply fungsi format indonesia
            df['TGL_EFEKTIF'] = df['TGL_EFEKTIF'].apply(format_tanggal_indo)
            
            # C. Perbaiki Jam (Kolom E)
            df['JAM_TRAN'] = df['JAM_TRAN'].apply(format_jam)
            
            # --- SIMPAN KE EXCEL ---
            file_xlsx = os.path.splitext(file_csv)[0] + '.xlsx'
            
            with pd.ExcelWriter(file_xlsx, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Sheet1')
                
                workbook = writer.book
                worksheet = writer.sheets['Sheet1']
                
                # D. Format Kolom Keuangan (H, I, J, K) -> Index 7, 8, 9, 10
                # Kolom Excel: H, I, J, K
                finance_cols = ['H', 'I', 'J', 'K']
                
                for col_letter in finance_cols:
                    for cell in worksheet[col_letter]:
                        if cell.row > 1: # Lewati header
                            cell.number_format = '#,##0.00'
                            
                # Pastikan Kolom NOREK (B) formatnya Text di Excel agar tidak ada error number stored as text
                for cell in worksheet['B']:
                    if cell.row > 1:
                        cell.number_format = '@'

                # E. Auto-fit Lebar Kolom
                auto_adjust_excel_width(worksheet, df)
                
            print(f"--> Selesai! Disimpan: {file_xlsx}")
            
        except Exception as e:
            print(f"--> Gagal pada {file_csv}. Error: {e}")

    print("\nSemua proses selesai.")