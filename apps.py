import requests 
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import pandas as pd
from datetime import datetime, timedelta
import os
import subprocess
import json
from config import API_TOKEN, CABANG_OPTIONS
import openpyxl
from openpyxl.styles import Font, PatternFill

CONFIG_PATH = "config.json"

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
    import os
    import subprocess
    import openpyxl
    from openpyxl.styles import Font
    from tkinter import filedialog
    import pandas as pd

    # Hapus kolom yang gak dibutuhin
    kolom_buang = ['finish.lat', 'finish.lon', 'finish.notes']
    dataframe = dataframe.drop(columns=[col for col in kolom_buang if col in dataframe.columns])

    # Rename kolom awal
    kolom_baru = {
        'Email': 'Driver',
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

    # Bikin angka jarak jadi 2 angka di belakang koma
    if 'Total Distance' in dataframe.columns:
        dataframe['Total Distance'] = pd.to_numeric(dataframe['Total Distance'], errors='coerce')
        dataframe['Total Distance'] = dataframe['Total Distance'].round(2)


    # Urutin kolom
    urutan_kolom = [
        'Driver',
        'Start Date',
        'Start Time',
        'Finish Date',
        'Finish Time',
        'Tracked Time',
        'Total Duration',
        'Total Distance',
    ]
    dataframe = dataframe[urutan_kolom]

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

    # Bold 3 baris terakhir kolom pertama
    wb = openpyxl.load_workbook(filename)
    ws = wb.active
    max_row = ws.max_row

    merah_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

    for row in range(2, max_row + 1):  # mulai dari baris 2 (hindari header)
        start_date = ws.cell(row=row, column=2).value
        finish_date = ws.cell(row=row, column=4).value
        if start_date and finish_date and start_date != finish_date:
            ws.cell(row=row, column=2).fill = merah_fill
            ws.cell(row=row, column=4).fill = merah_fill

    # for row in range(max_row - 2, max_row + 1):
    #     ws.cell(row=row, column=1).font = Font(bold=True)

    wb.save(filename)

    # Buka file
    try:
        os.startfile(filename)
    except AttributeError:
        subprocess.run(['open', filename], check=True)

def ambil_data(tanggal_str):
    tanggal_obj = datetime.strptime(tanggal_str, "%d-%m-%Y")
    tanggal_input = tanggal_obj.strftime("%Y-%m-%d")
    tanggal_from = (tanggal_obj - timedelta(days=1)).strftime("%Y-%m-%d")

    url = "https://apiweb.mile.app/api/v3/location-histories"
    params = {
        "limit": 150,
        "startFinish": "true",
        "fields": "finish,startTime",
        "timeTo": f"{tanggal_input} 23:59:59",
        "timeFrom": f"{tanggal_from} 00:00:00",
        "timeBy": "createdTime"
    }
    headers = {
        "Authorization": API_TOKEN
    }

    response = requests.get(url, params=params, headers=headers)
    
    if response.json()["tasks"]["data"] == []:
        tk.messagebox.showerror("Error", "Data tidak ditemukan")
        return

    if response.status_code == 200:
        result = response.json()
        items = result.get("tasks", {}).get("data", [])

        items = [
            item for item in items
            if item.get("trackedTime") not in [None, 0]
            and item.get("trackedTime", 0) >= 10
            and item.get("finish", {}).get("totalDistance", float('inf')) > 5
        ]


        for item in items:
            item["_id"] = extract_email_from_id(item["_id"])
            item["startTime"] = tambah_7_jam(item["startTime"])
            if "finish" in item and item["finish"]:
                item["finish"]["finishTime"] = tambah_7_jam(item["finish"]["finishTime"])
                if "totalDuration" in item["finish"]:
                    item["finish"]["totalDuration"] = convert_to_jam(item["finish"]["totalDuration"])
            if "trackedTime" in item:
                item["trackedTime"] = convert_to_jam(item["trackedTime"])

        # Buat dataframe
        df = pd.json_normalize(items)
        df.rename(columns={"_id": "Email"}, inplace=True)

        # Filter startTime sesuai tanggal input (setelah ditambah 7 jam)
        df['startTime'] = pd.to_datetime(df['startTime'])
        df = df[df['startTime'].dt.strftime("%Y-%m-%d") == tanggal_input]

        # --- START FILTER & MAPPING DRIVER ---
        try:
            # Filter email berdasarkan nama cabang dari config.json
            config = load_config()
            if config and "cabang" in config:
                nama_cabang = config["cabang"].lower()
                df = df[df['Email'].str.contains(nama_cabang)]

            mapping_df = pd.read_excel("Master_Driver.xlsx")
            
            # Cek kolom yang dibutuhkan ada gak
            if not {'Email', 'Driver'}.issubset(mapping_df.columns):
                raise ValueError("File Master_Driver.xlsx harus ada kolom 'Email' dan 'Driver'")
            
            # Merge untuk mapping Driver berdasarkan Email
            df = df.merge(mapping_df[['Email', 'Driver']], on='Email', how='left')
            df['Email'] = df['Driver']
            df.drop(columns=['Driver'], inplace=True)
            df = df.sort_values(by='Email', ascending=True)

        except Exception as e:
            messagebox.showerror("Error", f"Gagal load atau mapping file Master_Driver.xlsx:\n{e}")

        # Tambah kode filter duplikat di sini
        def durasi_ke_menit(durasi_str):
            try:
                jam, menit = map(int, durasi_str.split(':'))
                return jam * 60 + menit
            except:
                return 0

        df['finish.totalDuration_menit'] = df['finish.totalDuration'].apply(durasi_ke_menit)
        df['finish.totalDistance'] = pd.to_numeric(df['finish.totalDistance'], errors='coerce').fillna(0)

        df = df.sort_values(by='Email')

        def filter_duplikat(grup):
            idx_to_drop = []
            for i in range(len(grup)):
                for j in range(len(grup)):
                    if i != j:
                        row_i = grup.iloc[i]
                        row_j = grup.iloc[j]
                        if (row_i['finish.totalDuration_menit'] < row_j['finish.totalDuration_menit'] and
                            row_i['finish.totalDistance'] < row_j['finish.totalDistance']):
                            idx_to_drop.append(grup.index[i])
                            break
            return grup.drop(idx_to_drop)

        groups = [filter_duplikat(group) for _, group in df.groupby('Email')]

        if groups == []:
            tk.messagebox.showerror("Error", "Data tidak ditemukan")
            return
        
        df = pd.concat(groups, ignore_index=True)


        df.drop(columns=['finish.totalDuration_menit'], inplace=True)

        # Tambahin 2 baris kosong
        # empty_rows = pd.DataFrame([[""] * len(df.columns)] * 2, columns=df.columns)

        # # Bikin baris note-nya
        # notes = pd.DataFrame({
        #     df.columns[0]: [
        #         "*Data diambil berdasarkan tanggal Start Trip.",
        #         "*Jika tidak muncul, berarti Driver belum Start Trip di tanggal tersebut atau belum Finish Trip dari hari sebelumnya.",
        #         "*Cek tanggal sebelumnya jika ada data yang hilang."
        #     ]
        # })

        # # Gabungin: data asli + 2 baris kosong + 3 baris note
        # df_final = pd.concat([df, empty_rows, notes], ignore_index=True)

        simpan_file_excel(df)

    else:
        print(f"Gagal request: {response.status_code}")
        print(response.text)

def pilih_cabang_gui():
    root = tk.Tk()
    root.title("Pilih Cabang")

    lebar = 350
    tinggi = 180
    x = (root.winfo_screenwidth() - lebar) // 2
    y = (root.winfo_screenheight() - tinggi) // 2
    root.geometry(f"{lebar}x{tinggi}+{x}+{y}")

    ttk.Label(root, text="Pilih Lokasi Cabang", font=("Helvetica", 12)).pack(pady=10)
    cabang_var = tk.StringVar()
    dropdown = ttk.Combobox(root, textvariable=cabang_var, values=list(CABANG_OPTIONS.keys()), font=("Helvetica", 12), state="readonly")
    dropdown.current(0)
    dropdown.pack(pady=10)

    def simpan_dan_lanjut():
        cabang_kode = CABANG_OPTIONS[cabang_var.get()]
        simpan_config({"cabang": cabang_kode})
        root.destroy()
        buka_tanggal_gui()  # lanjut ke GUI tanggal

    ttk.Button(root, text="Lanjut", command=simpan_dan_lanjut).pack(pady=10)
    root.mainloop()

def buka_tanggal_gui():
    config = load_config()
    if not config or "cabang" not in config:
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
    cal = DateEntry(root, date_pattern='dd-MM-yyyy', font=("Helvetica", 14))
    cal.pack(pady=10)

    def proses():
        tanggal = cal.get()
        ambil_data(tanggal)
        root.destroy()

    ttk.Button(root, text="Proses", command=proses).pack(pady=10)
    root.mainloop()

if __name__ == "__main__":
    config = load_config()
    if config and "cabang" in config:
        buka_tanggal_gui()
    else:
        pilih_cabang_gui()
