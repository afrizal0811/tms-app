from datetime import datetime, timedelta
from openpyxl.styles import PatternFill, Alignment 
from openpyxl.utils import get_column_letter 
from tkcalendar import DateEntry
from tkinter import ttk, messagebox, filedialog
import json
import openpyxl
import os
import pandas as pd
import requests 
import subprocess
import tkinter as tk
import sys
from path_manager import MASTER_JSON_PATH

CONFIG_PATH = "config.json"
CONSTANTS = {}

def resource_path(relative_path):
    """Mendapatkan path absolut ke resource dalam bundle atau development."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_base_path():
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)
    
def simpan_config(data):
    with open(CONFIG_PATH, 'w') as f:
        json.dump(data, f)

def load_config():
    if os.path.exists(CONFIG_PATH):
        with open(CONFIG_PATH, 'r') as f:
            return json.load(f)
    return None

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

def simpan_file_excel(dataframe):
    # Hapus kolom yang gak dibutuhin
    kolom_buang = ['finish.lat', 'finish.lon', 'finish.notes']
    dataframe = dataframe.drop(columns=[col for col in kolom_buang if col in dataframe.columns])

    # Rename kolom awal
    kolom_baru = {
        'Driver': 'Driver', 
        'startTime': 'Start Trip',
        'finish.finishTime': 'Finish Trip',
        'trackedTime': 'Tracked Time',
        'finish.totalDuration': 'Total Duration',
        'finish.totalDistance': 'Total Distance',
    }
    dataframe = dataframe.rename(columns=kolom_baru)

    # Convert ke datetime
    for kolom in ['Start Trip', 'Finish Trip']:
        if kolom in dataframe.columns:
            dataframe[kolom] = pd.to_datetime(dataframe[kolom], errors='coerce')

    # Tambah kolom tanggal & ubah format waktu
    if 'Start Trip' in dataframe.columns:
        dataframe.insert(
            dataframe.columns.get_loc('Start Trip'),
            'Start Date',
            dataframe['Start Trip'].dt.strftime('%d-%m-%Y')
        )
        dataframe['Start Trip'] = dataframe['Start Trip'].dt.strftime('%H:%M')

    if 'Finish Trip' in dataframe.columns:
        dataframe.insert(
            dataframe.columns.get_loc('Finish Trip'),
            'Finish Date',
            dataframe['Finish Trip'].dt.strftime('%d-%m-%Y')
        )
        dataframe['Finish Trip'] = dataframe['Finish Trip'].dt.strftime('%H:%M')

    # Rename kolom waktu
    dataframe = dataframe.rename(columns={
        'Start Trip': 'Start Time',
        'Finish Trip': 'Finish Time'
    })

    # --- START: Tambahkan tanda petik satu (') di depan Tracked Time dan Total Duration ---
    if 'Tracked Time' in dataframe.columns:
        # Pastikan kolom bertipe string sebelum menambahkan tanda petik
        dataframe['Tracked Time'] = dataframe['Tracked Time'].astype(str)
        dataframe['Tracked Time'] = dataframe['Tracked Time'].apply(lambda x: f"'{x}" if pd.notna(x) and x != 'None' and x != '' else x)
    if 'Total Duration' in dataframe.columns:
        # Pastikan kolom bertipe string sebelum menambahkan tanda petik
        dataframe['Total Duration'] = dataframe['Total Duration'].astype(str)
        dataframe['Total Duration'] = dataframe['Total Duration'].apply(lambda x: f"'{x}" if pd.notna(x) and x != 'None' and x != '' else x)
    # --- END: Tambahkan tanda petik satu (') ---

    # Bikin angka jarak jadi 2 angka di belakang koma
    if 'Total Distance' in dataframe.columns:
        dataframe['Total Distance'] = pd.to_numeric(dataframe['Total Distance'], errors='coerce')
        dataframe['Total Distance'] = dataframe['Total Distance'].round(2)
        # Ubah tipe data kolom ke 'object' agar bisa menampung string kosong
        dataframe['Total Distance'] = dataframe['Total Distance'].astype(object) 
        # KOSONGKAN NILAI JIKA 0
        dataframe.loc[dataframe['Total Distance'] == 0, 'Total Distance'] = '' 

    # Urutin kolom, tambahkan 'Plat' di awal
    urutan_kolom = [
        'Plat', 
        'Driver',
        'Start Date',
        'Start Time',
        'Finish Date',
        'Finish Time',
        'Tracked Time',
        'Total Duration',
        'Total Distance',
    ]
    # Filter kolom yang ada di dataframe
    urutan_kolom_final = [col for col in urutan_kolom if col in dataframe.columns]
    dataframe = dataframe[urutan_kolom_final]

    # Simpan file
    folder = filedialog.askdirectory(title="Pilih folder untuk menyimpan file")
    if not folder:
        return

    base_filename = "Hasil Start Finish"
    filename = os.path.join(folder, f"{base_filename}.xlsx")
    counter = 1
    while os.path.exists(filename):
        filename = os.path.join(folder, f"{base_filename} - {counter}.xlsx")
        counter += 1

    dataframe.to_excel(filename, index=False)

    # --- Start: Perubahan untuk format Excel ---
    wb = openpyxl.load_workbook(filename)
    ws = wb.active

    # 1. Rata tengah (Center) untuk kolom yang ditentukan
    kolom_rata_tengah = ['Plat', 'Start Date', 'Start Time', 'Finish Date', 'Finish Time', 'Tracked Time', 'Total Duration', 'Total Distance']
    
    # Buat dictionary untuk memetakan nama kolom ke indeks kolom (numerik 1, 2, 3...)
    col_idx_map = {col_name: idx + 1 for idx, col_name in enumerate(dataframe.columns)}


    center_alignment = Alignment(horizontal='center', vertical='center')

    for col_name in kolom_rata_tengah:
        if col_name in col_idx_map: # Pastikan kolom ada di dataframe
            col_number = col_idx_map[col_name]
            for row_idx in range(1, ws.max_row + 1): # Iterasi dari baris 1 (header) hingga akhir
                cell = ws.cell(row=row_idx, column=col_number)
                cell.alignment = center_alignment

    # 2. Atur lebar kolom otomatis
    for column_idx, column_name in enumerate(dataframe.columns):
        # Dapatkan huruf kolom (misal 'A', 'B') dari indeks numerik
        column_letter = get_column_letter(column_idx + 1)
        max_length = 0

        for row_idx in range(1, ws.max_row + 1):
            cell_value = ws.cell(row=row_idx, column=column_idx + 1).value
            try:
                # Konversi ke string untuk menghitung panjang, dan handle None
                cell_str = str(cell_value) if cell_value is not None else ""
                if len(cell_str) > max_length:
                    max_length = len(cell_str)
            except:
                pass
        
        # Tambahkan sedikit padding (misal 2-3 karakter)
        adjusted_width = (max_length + 2) 
        if adjusted_width > 0: # Pastikan lebarnya positif
            ws.column_dimensions[column_letter].width = adjusted_width

    # Warna merah untuk tanggal start/finish yang berbeda
    merah_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    # Dapatkan indeks kolom numerik untuk 'Start Date' dan 'Finish Date'
    start_date_col_idx = col_idx_map.get('Start Date')
    finish_date_col_idx = col_idx_map.get('Finish Date')

    if start_date_col_idx and finish_date_col_idx: # Pastikan kolom ditemukan
        for row in range(2, ws.max_row + 1):  # mulai dari baris 2 (hindari header)
            start_date_cell = ws.cell(row=row, column=start_date_col_idx)
            finish_date_cell = ws.cell(row=row, column=finish_date_col_idx)

            if start_date_cell.value and finish_date_cell.value and start_date_cell.value != finish_date_cell.value:
                start_date_cell.fill = merah_fill
                finish_date_cell.fill = merah_fill

    wb.save(filename)
    # --- End: Perubahan untuk format Excel ---

    # Buka file
    try:
        os.startfile(filename)
    except AttributeError:
        subprocess.run(['open', filename], check=True)

def get_constant_file_path(base_path):
    # 1. Cek di bundle (PyInstaller)
    bundle_path = resource_path("modules/constant.json")
    if os.path.exists(bundle_path):
        return bundle_path

    # 2. Cek di root project (development)
    dev_path = os.path.join(base_path, "constant.json")
    if os.path.exists(dev_path):
        return dev_path

    raise FileNotFoundError("File constant.json tidak ditemukan di bundle maupun root project.")

def ambil_data(tanggal_str, base_path):
    # --- Load token dari constant.json ---
    try:
        constants_path = get_constant_file_path(base_path)
        with open(constants_path, "r", encoding="utf-8") as f:
            constants = json.load(f)
    except FileNotFoundError:
        messagebox.showerror("File Tidak Ditemukan", "constant.json tidak ditemukan di bundle atau folder project. \n\nHubungi Admin.")
        return
    except json.JSONDecodeError:
        messagebox.showerror("File Tidak Ditemukan", "constant.json tidak valid. \n\nHubungi Admin.")
        return
    
    api_token = constants.get("token")
    if not api_token:
        messagebox.showerror("Data Tidak Ditemukan", "Token tidak ditemukan. \n\nHubungi Admin.")
        return
    
    # --- Proses tanggal & API request seperti sebelumnya ---
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
    
    response = requests.get(url, params=params, headers=headers)
    if response.status_code == 200:
        if not response.json().get("tasks", {}).get("data"):
            tk.messagebox.showerror("Error", "Data tidak ditemukan dari API.") 
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

            # --- PERUBAHAN UTAMA DI SINI ---
            try:
                # 1. Menggunakan path terpusat untuk membaca master.json
                mapping_df = pd.read_json(MASTER_JSON_PATH)
            except FileNotFoundError:
                # 2. Memperbarui pesan error
                messagebox.showerror("File Tidak Ditemukan", "Master data driver tidak dapat ditemukan.")
                return
            # --- AKHIR PERUBAHAN ---
            
            required_master_cols = {'Email', 'Driver', 'Plat'}
            if not required_master_cols.issubset(mapping_df.columns):
                # 3. Memperbarui pesan error
                raise ValueError(f"Master data driver harus ada key: {', '.join(required_master_cols)}")

            mapping_df_filtered = mapping_df[mapping_df['Email'].str.contains(nama_lokasi, na=False, case=False)]
            df_merged = df_api_data.merge(mapping_df_filtered[['Email', 'Driver', 'Plat']], on='Email', how='left')
            
            df_merged['Driver'] = df_merged['Driver'].fillna(df_merged['Email']) 
            df_merged['Plat'] = df_merged['Plat'].fillna('') 

            df_merged.drop(columns=['Email'], inplace=True) 
            df_merged = df_merged.sort_values(by='Driver', ascending=True)

        except Exception as e:
            messagebox.showerror("Error", f"Gagal memproses master data driver:\n{e}")
            return

        # Sisa kode fungsi untuk menggabungkan driver yang hilang dan filter duplikat tetap sama
        # ... (kode dari "existing_drivers = ..." hingga akhir) ...
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
                jam, menit = map(int, durasi_str.split(':'))
                return jam * 60 + menit
            except (ValueError, AttributeError):
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

    else:
        result = response.json()
        tk.messagebox.showerror(
            "API Error",
            f"Status Code: {response.status_code}\n{result['message']}"
        )

def pilih_cabang_gui():
    lokasi_dict = {
        "01. Sidoarjo": "plsda", "02. Jakarta": "pljkt", "03. Bandung": "plbdg",
        "04. Semarang": "plsmg", "05. Yogyakarta": "plygy", "06. Malang": "plmlg",
        "07. Denpasar": "pldps", "08. Makasar": "plmks", "09. Jember": "pljbr"
    }
    config_path = os.path.join(get_base_path(), "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                return json.load(f).get("lokasi")
        except Exception:
            pass

    selected_value = None
    def on_select():
        nonlocal selected_value
        selected = combo.get()
        if selected in lokasi_dict:
            selected_value = lokasi_dict[selected]
            with open(config_path, "w") as f:
                json.dump({"lokasi": selected_value}, f)
            root.destroy()

    root = tk.Tk()
    root.title("Pilih Lokasi Cabang")
    lebar = 350
    tinggi = 180
    x = (root.winfo_screenwidth() - lebar) // 2
    y = (root.winfo_screenheight() - tinggi) // 2
    root.geometry(f"{lebar}x{tinggi}+{x}+{y}")

    tk.Label(root, text="Pilih Lokasi Cabang:", font=("Arial", 14)).pack(pady=10)
    combo = ttk.Combobox(root, values=list(lokasi_dict.keys()), font=("Arial", 12))
    combo.pack(pady=10)
    combo.current(0)
    tk.Button(root, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)
    root.mainloop()
    return selected_value

def buka_tanggal_gui(base_path):
    config = load_config()
    if not config or "lokasi" not in config:
        pilih_cabang_gui()
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
        ambil_data(tanggal, base_path)
        root.destroy()

    ttk.Button(root, text="Proses", command=proses).pack(pady=10)
    root.mainloop()

def main():
    base_path = os.path.dirname(sys.executable) if getattr(sys, 'frozen', False) else os.path.dirname(__file__)
    config = load_config()
    if config and "lokasi" in config:
        buka_tanggal_gui(base_path)
    else:
        messagebox.showwarning("Dibatalkan", "Pilih lokasi cabang!")

if __name__ == "__main__":
    main()