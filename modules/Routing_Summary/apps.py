# modules/Routing_Summary/apps.py (KODE BARU)

import traceback
import tkinter as tk
from tkinter import filedialog, messagebox
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
from ..shared_utils import load_config, load_master_data, get_save_path, open_file_externally


# ==============================================================================
# FUNGSI-FUNGSI UTAMA (HELPER FUNCTIONS)
# ==============================================================================

# Fungsi yang duplikat (simpan_file_excel, buka_file, get_base_path, load_config)
# TELAH DIHAPUS DARI SINI.

# Fungsi spesifik untuk modul ini bisa tetap di sini.
def pilih_file_excel(prompt="Pilih file Excel"):
    """Membuka dialog untuk memilih satu file Excel."""
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Upload File", prompt)
    file_path = filedialog.askopenfilename(
        title=prompt,
        filetypes=[("Excel files", "*.xlsx *.xls")]
    )
    return file_path

def contains_capacity_constraint(file_path):
    """Mengecek apakah baris-baris awal file Excel mengandung 'capacity constraint'."""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=20, values_only=True):
            if any("capacity constraint" in str(cell).lower() for cell in row if cell):
                return True
    except Exception:
        return False
    return False


# ==============================================================================
# FUNGSI-FUNGSI PEMROSESAN INTI
# ==============================================================================

def buat_mapping_driver(lokasi_value):
    """
    Membuat mapping email ke nama driver dari master data berdasarkan lokasi.
    Kini menggunakan shared_utils.
    """
    # 2. Menggunakan fungsi terpusat untuk memuat data master
    master_df = load_master_data(lokasi_value)
    if master_df is None or master_df.empty:
        return {}
    
    # Membuat dictionary dari DataFrame
    mapping = pd.Series(master_df.Driver.values, index=master_df.Email).to_dict()
    return mapping

def proses_truck_detail(workbook, source_df, lokasi):
    """
    Memproses detail truk dari DataFrame sumber, dengan mengakumulasi data
    untuk entri duplikat (Vehicle Name + Assignee).
    """
    email_to_name = buat_mapping_driver(lokasi)
    df = source_df.copy()

    required_cols = {
        "Vehicle Name": "", "Assignee": "", "Weight Percentage": "",
        "Volume Percentage": "", "Total Distance (m)": 0, "Total Visits": "",
        "Total Spent Time (mins)": 0
    }

    for col, default in required_cols.items():
        if col not in df.columns:
            df[col] = default

    # --- AWAL PERUBAHAN ---
    # Fungsi bantuan untuk mengubah kolom menjadi numerik sebelum agregasi
    def parse_to_numeric(value, is_percentage=False):
        try:
            val_str = str(value).strip()
            if is_percentage:
                val_str = val_str.replace('%', '')
            # Menghapus koma sebagai pemisah ribuan
            return float(val_str.replace(',', ''))
        except (ValueError, TypeError):
            return 0.0

    # Konversi kolom yang akan diakumulasi menjadi tipe data numerik
    df['Weight Percentage'] = df['Weight Percentage'].apply(lambda x: parse_to_numeric(x, is_percentage=True))
    df['Volume Percentage'] = df['Volume Percentage'].apply(lambda x: parse_to_numeric(x, is_percentage=True))
    df['Total Distance (m)'] = df['Total Distance (m)'].apply(parse_to_numeric)
    df['Total Spent Time (mins)'] = df['Total Spent Time (mins)'].apply(parse_to_numeric)
    
    # Tentukan aturan agregasi
    agg_rules = {
        'Weight Percentage': 'sum',
        'Volume Percentage': 'sum',
        'Total Distance (m)': 'sum',
        'Total Spent Time (mins)': 'sum'
    }
    
    # Kelompokkan berdasarkan 'Vehicle Name' dan 'Assignee', lalu agregasi
    # Hanya baris di mana kedua kolom ini tidak kosong yang akan dikelompokkan
    df_grouped = df.dropna(subset=['Vehicle Name', 'Assignee'])
    df_agg = df_grouped.groupby(['Vehicle Name', 'Assignee']).agg(agg_rules).reset_index()
    
    # Gabungkan kembali dengan baris yang tidak memiliki 'Vehicle Name' atau 'Assignee'
    df_others = df[df['Vehicle Name'].isna() | df['Assignee'].isna()]
    df = pd.concat([df_agg, df_others], ignore_index=True)
    # --- AKHIR PERUBAHAN ---

    def to_h_mm(minutes):
        try:
            minutes = float(str(minutes).replace(",", "").strip())
            hours = int(minutes // 60)
            mins = int(round(minutes % 60))
            return f"'{hours}:{mins:02d}"
        except (ValueError, TypeError): return ""

    def format_percentage(value):
        try:
            # Nilai sudah berupa numerik hasil penjumlahan, tinggal format
            val_float = float(value)
            return f"{val_float:.1f}%"
        except (ValueError, TypeError): return ""

    df['Ship Duration'] = df['Total Spent Time (mins)'].apply(to_h_mm)
    df['Weight Percentage'] = df['Weight Percentage'].apply(format_percentage)
    df['Volume Percentage'] = df['Volume Percentage'].apply(format_percentage)
    
    df['Total Distance (m)'] = df['Total Distance (m)'].astype(int)

    # Mengosongkan kolom 'Total Visits' dan 'Total Delivered' sesuai permintaan
    df["Total Visits"] = ""
    if "Total Delivered" not in df.columns:
        df["Total Delivered"] = ""
    else:
        df["Total Delivered"] = ""
    
    df["Assignee"] = df["Assignee"].str.lower().map(email_to_name).fillna(df["Assignee"])
    
    existing_drivers = set(df["Assignee"].dropna())
    all_drivers = set(email_to_name.values())
    missing_drivers = sorted(list(all_drivers - existing_drivers))
    if missing_drivers:
        missing_df = pd.DataFrame(missing_drivers, columns=["Assignee"])
        df = pd.concat([df, missing_df], ignore_index=True)

    df.sort_values(by="Assignee", inplace=True, na_position='last')

    final_cols = ["Vehicle Name", "Assignee", "Weight Percentage", "Volume Percentage", "Total Distance (m)", "Total Visits", "Total Delivered", "Ship Duration"]
    sheet_detail = workbook.create_sheet(title="Truck Detail")
    sheet_detail.append(final_cols)

    # Pastikan semua kolom final ada di DataFrame sebelum diakses
    for col in final_cols:
        if col not in df.columns:
            df[col] = ""

    for r in df[final_cols].to_records(index=False):
        sheet_detail.append(list(r))
            
    center_align_cols = ["Weight Percentage", "Volume Percentage", "Total Distance (m)", "Total Visits", "Total Delivered", "Ship Duration"]
    for col_idx, col_title in enumerate(final_cols, 1):
        max_length = len(str(col_title))
        for row_idx in range(2, sheet_detail.max_row + 1):
            cell_value = sheet_detail.cell(row=row_idx, column=col_idx).value
            if cell_value: max_length = max(max_length, len(str(cell_value)))
        col_letter = get_column_letter(col_idx)
        sheet_detail.column_dimensions[col_letter].width = max_length + 2
        alignment = Alignment(horizontal='center') if col_title in center_align_cols else Alignment(horizontal='left')
        for cell in sheet_detail[col_letter]: cell.alignment = alignment

    df['Type'] = df['Assignee'].str.extract(r'(DRY|FRZ)', expand=False, flags=re.IGNORECASE).fillna('OTHER')
    distance_summary = df.groupby('Type')['Total Distance (m)'].sum() / 1000
    sheet_dist = workbook.create_sheet(title="Total Distance Summary")
    sheet_dist["A1"] = "DRY"; sheet_dist["B1"] = "FRZ"
    sheet_dist["A2"] = round(distance_summary.get("DRY", 0), 2)
    sheet_dist["B2"] = round(distance_summary.get("FRZ", 0), 2)


def proses_truck_usage(workbook, source_df):
    """
    Memproses ringkasan penggunaan truk dari DataFrame sumber.
    Setiap truk (kombinasi Vehicle Name + Assignee) dihitung sekali.
    """
    master_df = load_master_data()
    upload_df = source_df.copy()

    # --- AWAL PERUBAHAN ---
    # Hapus duplikat berdasarkan 'Vehicle Name' dan 'Assignee' untuk memastikan
    # setiap truk dihitung hanya satu kali.
    upload_df.drop_duplicates(subset=['Vehicle Name', 'Assignee'], inplace=True, ignore_index=True)
    # --- AKHIR PERUBAHAN ---

    plat_type_map = dict(zip(master_df["Plat"].astype(str), master_df["Type"]))
    def find_vehicle_tag(vehicle_name):
        if pd.isna(vehicle_name): return ""
        for plat, v_type in plat_type_map.items():
            if plat in str(vehicle_name): return v_type
        return ""
    upload_df["Vehicle Tags"] = upload_df["Vehicle Name"].apply(find_vehicle_tag)
    
    upload_df.loc[upload_df["Vehicle Tags"].str.contains("HAVI", na=False, case=False), "Vehicle Tags"] = "Fuso-DRY"
    upload_df.loc[upload_df["Vehicle Tags"].str.contains("KFC", na=False, case=False), "Vehicle Tags"] = "CDD-Long-FROZEN"
    
    upload_df["Vehicle Tags"] = upload_df["Vehicle Tags"].str.upper()
    
    dry_df = upload_df[upload_df["Vehicle Tags"].str.contains("DRY", na=False)]
    frozen_df = upload_df[upload_df["Vehicle Tags"].str.contains("FROZEN", na=False)]
    
    display_order = ["L300", "CDE", "CDE-LONG", "CDD", "CDD-LONG", "FUSO"]
    counting_order = ["CDE-LONG", "CDD-LONG", "L300", "CDE", "CDD", "FUSO"]

    def count_types(df_tags):
        counts = {v_type: 0 for v_type in display_order}
        tags_list = df_tags.tolist()
        
        for v_type in counting_order:
            found_count = 0
            remaining_tags = []
            for tag in tags_list:
                if pd.notna(tag) and v_type in tag:
                    found_count += 1
                else:
                    remaining_tags.append(tag)
            counts[v_type] = found_count
            tags_list = remaining_tags
        return counts
        
    dry_counts = count_types(dry_df["Vehicle Tags"].astype(str))
    frozen_counts = count_types(frozen_df["Vehicle Tags"].astype(str))

    sheet_usage = workbook.create_sheet(title="Truck Usage")
    sheet_usage["A1"] = "Tipe Kendaraan"
    sheet_usage["B1"] = "Jumlah (DRY)"; sheet_usage["C1"] = "Jumlah (FROZEN)"
    row = 2
    for v_type in display_order:
        sheet_usage[f"A{row}"] = v_type
        dry_count = dry_counts.get(v_type, 0)
        frozen_count = frozen_counts.get(v_type, 0)
        sheet_usage[f"B{row}"] = dry_count if dry_count != 0 else "-"
        sheet_usage[f"C{row}"] = frozen_count if frozen_count != 0 else "-"
        row += 1
    for col_letter in ["A", "B", "C"]: sheet_usage.column_dimensions[col_letter].width = 25


# ==============================================================================
# FUNGSI UTAMA (MAIN CONTROLLER)
# ==============================================================================

def main():
    """Fungsi controller yang menjalankan semua proses secara otomatis."""
    try:
        # 4. Pemanggilan fungsi diperbarui
        config = load_config()
            
        if config and "lokasi" in config:
            lokasi = config["lokasi"]
        else:
            messagebox.showwarning("Dibatalkan", "Pilih lokasi cabang!")
            return
        
        all_data = []
        index = 1
        while True:
            prompt = f"Pilih File Excel ke-{index}"
            path = pilih_file_excel(prompt)
            if not path:
                if index == 1: 
                    messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")
                    return
                else: 
                    break 

            skip_rows = 10 if contains_capacity_constraint(path) else 0
            df = pd.read_excel(path, skiprows=skip_rows)
            df.columns = df.columns.str.strip()
            all_data.append(df)
            
            index += 1
            lanjut = messagebox.askyesno("Konfirmasi", "Apakah ada file lain yang ingin diproses?")
            if not lanjut:
                break
        
        if not all_data:
            return

        combined_df = pd.concat(all_data, ignore_index=True)
        
        required_columns = [
            "Vehicle Name", "Assignee", "Weight Percentage", "Volume Percentage",
            "Total Distance (m)", "Total Visits", "Total Spent Time (mins)"
        ]
        missing_columns = [col for col in required_columns if col not in combined_df.columns]
        if missing_columns:
            messagebox.showerror(
                "Proses Gagal", "File tidak valid!\n\nUpload file Export Routing dengan benar!"
            )
            return
        
        email_prefixes = combined_df["Assignee"].dropna().str.extract(r'kendaraan\.([^.@]+)', expand=False)
        email_prefixes = email_prefixes.dropna().str.lower().unique()

        if not any(lokasi.lower() in prefix for prefix in email_prefixes):
            messagebox.showerror(
                "Proses Gagal", "Lokasi cabang tidak valid!\n\nLokasi cabang tidak sesuai dengan file Routing!"
            )
            return
        
        # Pengecekan path manual sudah tidak diperlukan lagi, karena ditangani oleh load_master_data
        
        output_wb = openpyxl.Workbook()
        output_wb.remove(output_wb.active)
    
        proses_truck_detail(output_wb, combined_df, lokasi)
        proses_truck_usage(output_wb, combined_df)

        if not output_wb.sheetnames:
            messagebox.showinfo("Selesai", "Tidak ada data yang diproses atau dihasilkan.")
            return
        
        # 5. Logika penyimpanan dan pembukaan file diubah
        save_path = get_save_path("Routing Summary")
        if save_path:
            output_wb.save(save_path)
            open_file_externally(save_path)
        else:
            messagebox.showwarning("Proses Gagal", "Penyimpanan file dibatalkan.")

    except Exception as e:
        error_message = traceback.format_exc()
        messagebox.showerror("Terjadi Kesalahan", f"Error: {e}\n\n{error_message}")


if __name__ == "__main__":
    main()