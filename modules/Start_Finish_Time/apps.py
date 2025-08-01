# modules/Start_Finish_Time/apps.py (KODE AKHIR)

from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from tkinter import ttk, messagebox
import openpyxl
import pandas as pd
import requests
import tkinter as tk

# Impor fungsi bantuan dari shared_utils dan gui_utils
from ..shared_utils import (
    load_config,
    load_constants,
    load_master_data,
    get_save_path,
    open_file_externally
)
from ..gui_utils import create_date_picker_window

# =============================================================================
# BAGIAN 1: FUNGSI-FUNGSI BANTU (HELPER FUNCTIONS)
# =============================================================================

def extract_email_from_id(_id):
    """Mengekstrak email dari ID MileApp."""
    parts = _id.split('_')
    return parts[1] if len(parts) > 1 else _id

def convert_to_jam(menit):
    """Mengonversi durasi dalam menit ke format jam:menit."""
    jam = menit // 60
    sisa_menit = menit % 60
    return f"{jam}:{sisa_menit:02}"

def tambah_7_jam(waktu_str):
    """Menambahkan 7 jam ke waktu string untuk penyesuaian zona waktu."""
    waktu = datetime.strptime(waktu_str, "%Y-%m-%d %H:%M:%S")
    return (waktu + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")

def simpan_file_excel(dataframe, lokasi_name, tanggal_str):
    """
    Menyimpan DataFrame ke file Excel dengan penamaan yang dinamis.
    
    Args:
        dataframe (pd.DataFrame): DataFrame yang akan disimpan.
        lokasi_name (str): Nama lokasi cabang.
        tanggal_str (str): Tanggal dalam format DD.MM.YYYY.
    """
    # Bagian pemrosesan data DataFrame TIDAK BERUBAH
    kolom_buang = ['finish.lat', 'finish.lon', 'finish.notes']
    dataframe = dataframe.drop(columns=[col for col in kolom_buang if col in dataframe.columns])

    kolom_baru = {
        'Driver': 'Driver', 'startTime': 'Start Trip', 'finish.finishTime': 'Finish Trip',
        'trackedTime': 'Tracked Time', 'finish.totalDuration': 'Total Duration',
        'finish.totalDistance': 'Total Distance',
    }
    dataframe = dataframe.rename(columns=kolom_baru)

    for kolom in ['Start Trip', 'Finish Trip']:
        if kolom in dataframe.columns:
            dataframe[kolom] = pd.to_datetime(dataframe[kolom], errors='coerce')

    if 'Start Trip' in dataframe.columns:
        dataframe.insert(dataframe.columns.get_loc('Start Trip'), 'Start Date', dataframe['Start Trip'].dt.strftime('%d-%m-%Y'))
        dataframe['Start Trip'] = dataframe['Start Trip'].dt.strftime('%H:%M')

    if 'Finish Trip' in dataframe.columns:
        dataframe.insert(dataframe.columns.get_loc('Finish Trip'), 'Finish Date', dataframe['Finish Trip'].dt.strftime('%d-%m-%Y'))
        dataframe['Finish Trip'] = dataframe['Finish Trip'].dt.strftime('%H:%M')

    dataframe = dataframe.rename(columns={'Start Trip': 'Start Time', 'Finish Trip': 'Finish Time'})

    if 'Tracked Time' in dataframe.columns:
        dataframe['Tracked Time'] = dataframe['Tracked Time'].astype(str).apply(lambda x: f"'{x}" if pd.notna(x) and x != 'None' and x != '' else x)
    if 'Total Duration' in dataframe.columns:
        dataframe['Total Duration'] = dataframe['Total Duration'].astype(str).apply(lambda x: f"'{x}" if pd.notna(x) and x != 'None' and x != '' else x)

    if 'Total Distance' in dataframe.columns:
        dataframe['Total Distance'] = pd.to_numeric(dataframe['Total Distance'], errors='coerce')
        dataframe['Total Distance'] = dataframe['Total Distance'].round(2)
        dataframe['Total Distance'] = dataframe['Total Distance'].astype(object)
        dataframe.loc[dataframe['Total Distance'] == 0, 'Total Distance'] = ''

    urutan_kolom = [
        'Plat', 'Driver', 'Start Date', 'Start Time', 'Finish Date',
        'Finish Time', 'Tracked Time', 'Total Duration', 'Total Distance',
    ]
    urutan_kolom_final = [col for col in urutan_kolom if col in dataframe.columns]
    dataframe = dataframe[urutan_kolom_final]

    # --- MODIFIKASI UNTUK NAMA FILE DINAMIS ---
    file_basename = f"Time Summary {lokasi_name} - {tanggal_str}"
    filename = get_save_path(file_basename)
    if not filename:
        messagebox.showwarning("Dibatalkan", "Penyimpanan file dibatalkan.")
        return

    dataframe.to_excel(filename, index=False)

    # Bagian styling Excel TIDAK BERUBAH
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    kolom_rata_tengah = ['Plat', 'Start Date', 'Start Time', 'Finish Date', 'Finish Time', 'Tracked Time', 'Total Duration', 'Total Distance']
    col_idx_map = {col_name: idx + 1 for idx, col_name in enumerate(dataframe.columns)}
    center_alignment = Alignment(horizontal='center', vertical='center')

    for col_name in kolom_rata_tengah:
        if col_name in col_idx_map:
            col_number = col_idx_map[col_name]
            for row_idx in range(1, ws.max_row + 1):
                ws.cell(row=row_idx, column=col_number).alignment = center_alignment

    for column_idx, column_name in enumerate(dataframe.columns):
        column_letter = get_column_letter(column_idx + 1)
        max_length = 0
        try:
            max_length = max(len(str(cell.value)) for cell in ws[column_letter] if cell.value is not None)
            adjusted_width = (max_length + 2)
            if adjusted_width > 0:
                ws.column_dimensions[column_letter].width = adjusted_width
        except ValueError:
            pass
            
    merah_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    start_date_col_idx = col_idx_map.get('Start Date')
    finish_date_col_idx = col_idx_map.get('Finish Date')
    if start_date_col_idx and finish_date_col_idx:
        for row in range(2, ws.max_row + 1):
            start_date_cell = ws.cell(row=row, column=start_date_col_idx)
            finish_date_cell = ws.cell(row=row, column=finish_date_col_idx)
            if start_date_cell.value and finish_date_cell.value and start_date_cell.value != finish_date_cell.value:
                start_date_cell.fill = merah_fill
                finish_date_cell.fill = merah_fill

    wb.save(filename)
    
    open_file_externally(filename)


# =============================================================================
# BAGIAN 2: FUNGSI-FUNGSI PEMROSESAN UTAMA
# =============================================================================

def ambil_data(dates, app_instance=None):
    """
    Fungsi utama untuk mengambil data start-finish time dari API.
    
    Args:
        dates (dict): Dictionary yang berisi format tanggal, misal {"dmy": "25-07-2025"}.
        app_instance (object): Instance GUI untuk update status.
    """
    tanggal_str = dates["dmy"]
    
    if app_instance:
        app_instance.update_status("Mengambil data dari API...")

    constants = load_constants()
    config = load_config()
    if not all([constants, config]):
        return
    
    api_token = constants.get("token")
    lokasi_code = config.get("lokasi")
    
    if not api_token:
        messagebox.showerror("Data Tidak Ditemukan", "Token API tidak ditemukan di constant.json.")
        return
    if not lokasi_code:
        messagebox.showerror("Konfigurasi Salah", "Lokasi cabang tidak ditemukan di config.json.")
        return
        
    lokasi_mapping = constants.get('lokasi_mapping', {})
    lokasi_name = next((name for name, code in lokasi_mapping.items() if code == lokasi_code), lokasi_code)
    
    tanggal_obj = datetime.strptime(tanggal_str, "%d-%m-%Y")
    tanggal_input = tanggal_obj.strftime("%Y-%m-%d")
    tanggal_from = (tanggal_obj - timedelta(days=1)).strftime("%Y-%m-%d")

    url = "https://apiweb.mile.app/api/v3/location-histories"
    params = {
        "limit": 150, "startFinish": "true", "fields": "finish,startTime",
        "timeTo": f"{tanggal_input} 23:59:59", "timeFrom": f"{tanggal_from} 00:00:00",
        "timeBy": "createdTime"
    }
    headers = {"Authorization": f"Bearer {api_token}"}
    
    try:
        response = requests.get(url, params=params, headers=headers, timeout=30)
        response.raise_for_status()
    except requests.exceptions.RequestException as e:
        messagebox.showerror("API Error", f"Gagal terhubung ke API:\n{e}")
        return

    if not response.json().get("tasks", {}).get("data"):
        tk.messagebox.showerror("Error", "Data tidak ditemukan dari API untuk tanggal yang dipilih.")
        return

    result = response.json()
    items = result.get("tasks", {}).get("data", [])

    filtered_items = []
    for item in items:
        if item.get("trackedTime") not in [None, 0] and item.get("trackedTime", 0) >= 10 and item.get("finish", {}).get("totalDistance", float('inf')) > 5:
            item_copy = item.copy()
            item_copy["_id"] = extract_email_from_id(item_copy["_id"])
            item_copy["startTime"] = tambah_7_jam(item_copy["startTime"])
            if item_copy.get("finish"):
                item_copy["finish"]["finishTime"] = tambah_7_jam(item_copy["finish"]["finishTime"])
                if "totalDuration" in item_copy["finish"]:
                    item_copy["finish"]["totalDuration"] = convert_to_jam(item_copy["finish"]["totalDuration"])
            if "trackedTime" in item_copy:
                item_copy["trackedTime"] = convert_to_jam(item_copy["trackedTime"])
            filtered_items.append(item_copy)

    df_api_data = pd.json_normalize(filtered_items)
    if df_api_data.empty:
        messagebox.showinfo("Informasi", "Tidak ada data trip yang memenuhi kriteria untuk diproses.")
        return
        
    df_api_data.rename(columns={"_id": "Email"}, inplace=True)

    if 'startTime' in df_api_data.columns:
        df_api_data['startTime'] = pd.to_datetime(df_api_data['startTime'])
        df_api_data = df_api_data[df_api_data['startTime'].dt.strftime("%Y-%m-%d") == tanggal_input]

    try:
        if 'Email' in df_api_data.columns:
            df_api_data = df_api_data[df_api_data['Email'].str.contains(lokasi_code, na=False, case=False)]

        if df_api_data.empty:
            tk.messagebox.showerror("Error", "Data tidak ditemukan untuk lokasi yang dipilih.")
            return

        mapping_df = load_master_data(lokasi_cabang=lokasi_code)
        if mapping_df is None: return

        required_master_cols = {'Email', 'Driver', 'Plat'}
        if not required_master_cols.issubset(mapping_df.columns):
            raise ValueError(f"Master data driver harus ada kolom: {', '.join(required_master_cols)}")

        mapping_df_filtered = mapping_df[mapping_df['Email'].str.contains(lokasi_code, na=False, case=False)]
        df_merged = df_api_data.merge(mapping_df_filtered[['Email', 'Driver', 'Plat']], on='Email', how='left')
        
        df_merged['Driver'] = df_merged['Driver'].fillna(df_merged['Email'])
        df_merged['Plat'] = df_merged['Plat'].fillna('')

        df_merged.drop(columns=['Email'], inplace=True)
        df_merged = df_merged.sort_values(by='Driver', ascending=True)

    except Exception as e:
        messagebox.showerror("Error", f"Gagal memproses master data driver:\n{e}")
        return

    existing_drivers = df_merged['Driver'].unique().tolist()
    missing_rows_df = mapping_df_filtered[~mapping_df_filtered['Driver'].isin(existing_drivers)].copy()
    
    if not missing_rows_df.empty:
        missing_drivers_df_processed = missing_rows_df[['Driver', 'Plat']].copy()
        for col in df_merged.columns:
            if col not in missing_drivers_df_processed.columns:
                missing_drivers_df_processed[col] = ""
        missing_drivers_df_processed = missing_drivers_df_processed[df_merged.columns]
        final_df = pd.concat([df_merged, missing_drivers_df_processed], ignore_index=True)
    else:
        final_df = df_merged.copy()

    def durasi_ke_menit(durasi_str):
        try:
            if isinstance(durasi_str, str) and ':' in durasi_str:
                jam, menit = map(int, durasi_str.split(':'))
                return jam * 60 + menit
        except (ValueError, AttributeError):
            pass
        return 0

    final_df['finish.totalDuration_menit'] = final_df['finish.totalDuration'].apply(durasi_ke_menit)
    final_df['finish.totalDistance'] = pd.to_numeric(final_df['finish.totalDistance'], errors='coerce').fillna(0)
    final_df = final_df.sort_values(by='Driver')

    def filter_duplikat(grup):
        idx_to_drop = set()
        for i in range(len(grup)):
            for j in range(len(grup)):
                if i != j:
                    row_i = grup.iloc[i]
                    row_j = grup.iloc[j]
                    if (row_i['finish.totalDuration_menit'] < row_j['finish.totalDuration_menit'] and
                        row_i['finish.totalDistance'] < row_j['finish.totalDistance']):
                        idx_to_drop.add(grup.index[i])
        return grup.drop(list(idx_to_drop))

    non_empty_df = final_df[final_df['finish.totalDuration_menit'] > 0]
    groups = [filter_duplikat(group) for _, group in non_empty_df.groupby('Driver')]
    filtered_df_final = pd.concat(groups, ignore_index=True) if groups else pd.DataFrame()

    kosong_df = final_df[final_df['finish.totalDuration_menit'] == 0]
    final_df = pd.concat([filtered_df_final, kosong_df], ignore_index=True)
    final_df.drop(columns=['finish.totalDuration_menit'], inplace=True)
    final_df = final_df.sort_values(by='Driver', ascending=True)
    
    tanggal_format_titik = tanggal_str.replace('-', '.')
    simpan_file_excel(final_df, lokasi_name, tanggal_format_titik)

    if app_instance:
        app_instance.update_status("Proses selesai.")

# =============================================================================
# BAGIAN 3: FUNGSI GUI DAN EKSEKUSI
# =============================================================================

def main():
    """Fungsi utama untuk modul Start Finish Time."""
    config = load_config()
    if not config or not config.get("lokasi"):
        messagebox.showinfo("Setup Awal", "Lokasi cabang belum diatur. Silakan atur melalui menu Pengaturan > Ganti Lokasi Cabang.")
        return

    def process_wrapper(dates, app_instance):
        ambil_data(dates, app_instance)

    create_date_picker_window("Start-Finish Time", process_wrapper)

if __name__ == "__main__":
    main()