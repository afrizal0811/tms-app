# modules/Start_Finish_Time/apps.py (KODE BARU)

from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter
from tkcalendar import DateEntry
from tkinter import ttk, messagebox, filedialog
import openpyxl
import pandas as pd
import requests
import tkinter as tk

# 1. Impor fungsi bantuan dari shared_utils
from ..shared_utils import (
    load_config,
    load_constants,
    load_master_data,
    get_save_path,
    open_file_externally,
    CONFIG_PATH  # Impor juga path konstanta jika diperlukan
)
from ..gui_utils import create_date_picker_window

# =============================================================================
# BAGIAN 1: FUNGSI-FUNGSI BANTU (HELPER FUNCTIONS)
# =============================================================================

# Fungsi-fungsi duplikat (resource_path, get_base_path, simpan_config, load_config,
# get_constant_file_path) TELAH DIHAPUS.

# Fungsi spesifik untuk modul ini tetap dipertahankan.
def extract_email_from_id(_id):
    parts = _id.split('_')
    return parts[1] if len(parts) > 1 else _id

def convert_to_jam(menit):
    jam = menit // 60
    sisa_menit = menit % 60
    return f"{jam}:{sisa_menit:02}"

def tambah_7_jam(waktu_str):
    waktu = datetime.strptime(waktu_str, "%Y-%m-%d %H:%M:%S")
    return (waktu + timedelta(hours=7)).strftime("%Y-%m-%d %H:%M:%S")

# 2. Fungsi simpan_file_excel diubah untuk menggunakan shared_utils
def simpan_file_excel(dataframe):
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

    # Bagian penyimpanan file diubah menggunakan shared_utils
    filename = get_save_path("Hasil Start Finish")
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
    
    # Buka file menggunakan shared_utils
    open_file_externally(filename)


# =============================================================================
# BAGIAN 2: FUNGSI-FUNGSI PEMROSESAN UTAMA
# =============================================================================

# 3. Fungsi ambil_data diubah untuk menggunakan shared_utils
def ambil_data(dates, app_instance=None): # app_instance opsional
    """
    Fungsi utama untuk mengambil data start-finish time dari API.
    """
    # Gunakan format tanggal yang sesuai
    tanggal_str = dates["dmy"] 
    
    # Update status jika app_instance tersedia
    if app_instance:
        app_instance.update_status("Mengambil data dari API...")

    constants = load_constants()
    if not constants:
        return # Pesan error sudah ditangani di dalam shared_utils
    
    api_token = constants.get("token")
    if not api_token:
        messagebox.showerror("Data Tidak Ditemukan", "Token API tidak ditemukan di constant.json.\n\nHubungi Admin.")
        return
    
    # --- Proses tanggal & API request (TIDAK BERUBAH) ---
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
        response.raise_for_status() # Cek jika status code bukan 2xx
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
        config = load_config()
        nama_lokasi = config.get("lokasi", "").lower()

        if 'Email' in df_api_data.columns:
            df_api_data = df_api_data[df_api_data['Email'].str.contains(nama_lokasi, na=False, case=False)]

        if df_api_data.empty:
            tk.messagebox.showerror("Error", "Data tidak ditemukan untuk lokasi yang dipilih.")
            return

        # Memuat master data menggunakan shared_utils
        mapping_df = load_master_data()
        if mapping_df is None: return

        required_master_cols = {'Email', 'Driver', 'Plat'}
        if not required_master_cols.issubset(mapping_df.columns):
            raise ValueError(f"Master data driver harus ada kolom: {', '.join(required_master_cols)}")

        mapping_df_filtered = mapping_df[mapping_df['Email'].str.contains(nama_lokasi, na=False, case=False)]
        df_merged = df_api_data.merge(mapping_df_filtered[['Email', 'Driver', 'Plat']], on='Email', how='left')
        
        df_merged['Driver'] = df_merged['Driver'].fillna(df_merged['Email'])
        df_merged['Plat'] = df_merged['Plat'].fillna('')

        df_merged.drop(columns=['Email'], inplace=True)
        df_merged = df_merged.sort_values(by='Driver', ascending=True)

    except Exception as e:
        messagebox.showerror("Error", f"Gagal memproses master data driver:\n{e}")
        return

    # Sisa logika pemrosesan (TIDAK BERUBAH)
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

    simpan_file_excel(final_df)
    # Beri tahu user jika proses selesai
    if app_instance:
        app_instance.update_status("Proses selesai.")

# =============================================================================
# BAGIAN 3: FUNGSI GUI DAN EKSEKUSI
# =============================================================================

def pilih_cabang_gui():
    """GUI untuk memilih cabang jika belum ada di config.json."""
    # Fungsi ini spesifik, namun kita gunakan CONFIG_PATH dari shared_utils
    # agar konsisten.
    lokasi_dict = {
        "01. Sidoarjo": "plsda", "02. Jakarta": "pljkt", "03. Bandung": "plbdg",
        "04. Semarang": "plsmg", "05. Yogyakarta": "plygy", "06. Malang": "plmlg",
        "07. Denpasar": "pldps", "08. Makasar": "plmks", "09. Jember": "pljbr"
    }
    
    dialog = tk.Toplevel()
    dialog.title("Pilih Lokasi Cabang")
    
    def on_select():
        selected_display = combo.get()
        if selected_display in lokasi_dict:
            kode_lokasi = lokasi_dict[selected_display]
            try:
                # Simpan langsung ke config.json
                with open(CONFIG_PATH, "w") as f:
                    import json
                    json.dump({"lokasi": kode_lokasi}, f, indent=2)
                dialog.destroy()
            except IOError as e:
                messagebox.showerror("Error", f"Gagal menyimpan konfigurasi:\n{e}")

    # Logika GUI tetap sama...
    lebar, tinggi = 350, 180
    x = (dialog.winfo_screenwidth() - lebar) // 2
    y = (dialog.winfo_screenheight() - tinggi) // 2
    dialog.geometry(f"{lebar}x{tinggi}+{x}+{y}")
    tk.Label(dialog, text="Pilih Lokasi Cabang:", font=("Arial", 14)).pack(pady=10)
    combo = ttk.Combobox(dialog, values=list(lokasi_dict.keys()), font=("Arial", 12), state="readonly")
    combo.pack(pady=10)
    combo.current(0)
    tk.Button(dialog, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)
    
    dialog.transient()
    dialog.grab_set()
    dialog.wait_window()


def buka_tanggal_gui():
    """GUI utama untuk memilih tanggal."""
    config = load_config()
    if not config or "lokasi" not in config:
        pilih_cabang_gui()
        # Cek lagi setelah pemilihan
        config = load_config()
        if not config or "lokasi" not in config:
            messagebox.showwarning("Dibatalkan", "Pemilihan lokasi dibatalkan. Program akan berhenti.")
            return

    root = tk.Tk()
    root.title("Pilih Tanggal")

    lebar = 300
    tinggi = 150
    x = (root.winfo_screenwidth() - lebar) // 2
    y = (root.winfo_screenheight() - tinggi) // 2
    root.geometry(f"{lebar}x{tinggi}+{x}+{y}")

    ttk.Label(root, text="Pilih Tanggal", font=("Helvetica", 14)).pack(pady=10)
    cal = DateEntry(root, date_pattern='dd-MM-yyyy', font=("Helvetica", 14), style='TButton')
    cal.pack(pady=10)

    def proses():
        tanggal = cal.get()
        root.destroy()  # Tutup GUI dulu sebelum proses panjang
        ambil_data(tanggal)

    ttk.Button(root, text="Proses", command=proses).pack(pady=10)
    root.mainloop()

def main():
    """Fungsi utama untuk modul Start Finish Time."""
    # Pengecekan config awal bisa dilakukan di sini sebelum memanggil GUI
    config = load_config()
    if not config:
        messagebox.showinfo("Setup Awal", "Lokasi cabang belum diatur. Silakan atur melalui menu Pengaturan > Ganti Lokasi Cabang.")
        # panggil fungsi pilih lokasi jika masih mau ada popup dari sini
        # atau return agar user mengaturnya dari menu utama
        return

    # Definisikan fungsi wrapper yang akan dijalankan oleh GUI
    def process_wrapper(dates, app_instance):
        ambil_data(dates, app_instance)

    create_date_picker_window("Start-Finish Time", process_wrapper)

if __name__ == "__main__":
    main()