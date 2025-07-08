import os
import sys
import json
import subprocess
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import re
import pandas as pd
import openpyxl
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

CONFIG_PATH = "config.json"
#==============================================================================
# FUNGSI-FUNGSI UTAMA (HELPER FUNCTIONS)
#==============================================================================

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

def simpan_file_excel(wb, base_filename="Hasil_Proses"):
    """Membuka dialog untuk memilih lokasi penyimpanan file Excel."""
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Pilih lokasi untuk menyimpan file")
    if not folder_path:
        return None

    filename = f"{base_filename}.xlsx"
    full_path = os.path.join(folder_path, filename)
    counter = 1
    while os.path.exists(full_path):
        filename = f"{base_filename}_{counter}.xlsx"
        full_path = os.path.join(folder_path, filename)
        counter += 1

    wb.save(full_path)
    return full_path

def buka_file(path):
    """Membuka file dengan aplikasi default sistem operasi."""
    if sys.platform.startswith('darwin'): # macOS
        subprocess.call(('open', path))
    elif os.name == 'nt': # Windows
        os.startfile(path)
    elif os.name == 'posix': # Linux
        subprocess.call(('xdg-open', path))

def get_base_path():
    """Mendapatkan path dasar (base path) baik untuk script maupun executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)

def contains_capacity_constraint(file_path):
    """Mengecek apakah 10 baris pertama file Excel mengandung 'capacity constraint'."""
    try:
        wb = openpyxl.load_workbook(file_path, read_only=True)
        ws = wb.active
        for row in ws.iter_rows(min_row=1, max_row=20, values_only=True):
            if any("capacity constraint" in str(cell).lower() for cell in row if cell):
                return True
    except Exception:
        return False
    return False

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)
    return None

#==============================================================================
# FUNGSI-FUNGSI PEMROSESAN INTI
#==============================================================================

def buat_mapping_driver(master_path, lokasi_value):
    """Membuat mapping email ke nama driver dari file master berdasarkan lokasi."""
    wb = openpyxl.load_workbook(master_path)
    ws = wb.active
    header = [str(cell.value).strip() for cell in ws[1]]
    email_idx = header.index("Email")
    driver_idx = header.index("Driver")
    mapping = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        email, driver = row[email_idx], row[driver_idx]
        if email and driver and lokasi_value.lower() in str(email).lower():
            mapping[str(email).strip().lower()] = str(driver).strip()
    return mapping

def proses_truck_detail(workbook, source_df, master_path, lokasi):
    """Memproses detail truk dari DataFrame sumber."""
    email_to_name = buat_mapping_driver(master_path, lokasi)
    df = source_df.copy()

    required_cols = {
        "Vehicle Name": "", "Assignee": "", "Weight Percentage": "",
        "Volume Percentage": "", "Total Distance (m)": 0, "Total Visits": "",
        "Total Spent Time (mins)": 0
    }

    for col, default in required_cols.items():
        if col not in df.columns:
            df[col] = default
            
    def to_h_mm(minutes):
        try:
            minutes = float(str(minutes).replace(",", "").strip())
            hours = int(minutes // 60)
            mins = int(round(minutes % 60))
            return f"'{hours}:{mins:02d}"
        except (ValueError, TypeError): return ""

    def format_percentage(value):
        try:
            val_str = str(value).strip()
            if not val_str: return ""
            val_float = float(val_str.replace('%', '').replace(',', '').strip())
            if '%' not in val_str and val_float <= 1.0: val_float *= 100
            return f"{val_float:.1f}%"
        except (ValueError, TypeError): return ""

    df['Ship Duration'] = df['Total Spent Time (mins)'].apply(to_h_mm)
    df['Weight Percentage'] = df['Weight Percentage'].apply(format_percentage)
    df['Volume Percentage'] = df['Volume Percentage'].apply(format_percentage)
    
    df['Total Distance (m)'] = df['Total Distance (m)'].astype(str).str.replace(r'[^\d.]', '', regex=True)
    df['Total Distance (m)'] = pd.to_numeric(df['Total Distance (m)'], errors='coerce').fillna(0).astype(int)
    
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

def proses_truck_usage(workbook, source_df, master_path):
    """Memproses ringkasan penggunaan truk dari DataFrame sumber."""
    master_df = pd.read_excel(master_path)
    upload_df = source_df.copy()

    plat_type_map = dict(zip(master_df["Plat"].astype(str), master_df["Type"]))
    def find_vehicle_tag(vehicle_name):
        if pd.isna(vehicle_name): return ""
        for plat, v_type in plat_type_map.items():
            if plat in str(vehicle_name): return v_type
        return ""
    upload_df["Vehicle Tags"] = upload_df["Vehicle Name"].apply(find_vehicle_tag)
    
    dry_df = upload_df[upload_df["Vehicle Tags"].str.contains("DRY", na=False)]
    frozen_df = upload_df[upload_df["Vehicle Tags"].str.contains("FROZEN", na=False)]
    
    vehicle_types = ["L300", "CDE-Long", "CDE", "CDD-Long", "CDD", "Fuso"]
    def count_types(df_tags):
        counts = {v_type: 0 for v_type in vehicle_types}
        tags_list = df_tags.tolist()
        for v_type in vehicle_types:
            found_count = 0; remaining_tags = []
            for tag in tags_list:
                if pd.notna(tag) and v_type in tag: found_count += 1
                else: remaining_tags.append(tag)
            counts[v_type] = found_count
            tags_list = remaining_tags
        return counts
    dry_counts = count_types(dry_df["Vehicle Tags"].astype(str))
    frozen_counts = count_types(frozen_df["Vehicle Tags"].astype(str))

    sheet_usage = workbook.create_sheet(title="Truck Usage")
    sheet_usage["A1"] = "Tipe Kendaraan"
    sheet_usage["B1"] = "Jumlah (DRY)"; sheet_usage["C1"] = "Jumlah (FROZEN)"
    row = 2
    for v_type in vehicle_types:
        sheet_usage[f"A{row}"] = v_type
        dry_count = dry_counts.get(v_type, 0)
        frozen_count = frozen_counts.get(v_type, 0)
        sheet_usage[f"B{row}"] = dry_count if dry_count != 0 else "-"
        sheet_usage[f"C{row}"] = frozen_count if frozen_count != 0 else "-"
        row += 1
    for col_letter in ["A", "B", "C"]: sheet_usage.column_dimensions[col_letter].width = 25

#==============================================================================
# FUNGSI UTAMA (MAIN CONTROLLER)
#==============================================================================

def main():
    """Fungsi controller yang menjalankan semua proses secara otomatis."""
    try:
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
        # Cek kolom wajib
        required_columns = [
            "Vehicle Name", "Assignee", "Weight Percentage", "Volume Percentage",
            "Total Distance (m)", "Total Visits", "Total Spent Time (mins)"
        ]
        missing_columns = [col for col in required_columns if col not in combined_df.columns]
        if missing_columns:
            messagebox.showerror(
                "Proses Gagal",
                f"File tidak valid!\n" +
                "\n Upload kembali file Routing yang benar"
            )
            return
        # TAHAP 3: PROSES DATA GABUNGAN
        base_dir = get_base_path()
        master_path = os.path.join(base_dir, "Master_Driver.xlsx")
        if not os.path.exists(master_path):
            raise FileNotFoundError(f"File 'Master_Driver.xlsx' tidak ditemukan di folder:\n{base_dir}")

        output_wb = openpyxl.Workbook()
        output_wb.remove(output_wb.active)
    
        proses_truck_detail(output_wb, combined_df, master_path, lokasi)
        proses_truck_usage(output_wb, combined_df, master_path)

        # TAHAP 4: SIMPAN DAN BUKA HASIL
        if not output_wb.sheetnames:
            messagebox.showinfo("Selesai", "Tidak ada data yang diproses atau dihasilkan.")
            return
        
        save_path = simpan_file_excel(output_wb, "Hasil Truck Summary")
        if save_path:
            buka_file(save_path)
        else:
            messagebox.showwarning("Proses Gagal", "Penyimpanan file dibatalkan.")

    except Exception as e:
        error_message = traceback.format_exc()
        messagebox.showerror("Terjadi Kesalahan", f"Error: {e}\n\n{error_message}")


if __name__ == "__main__":
    main()