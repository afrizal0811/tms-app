import requests
import json
import pandas as pd
import os
import re
from openpyxl.styles import Alignment, PatternFill
from datetime import datetime
import tkinter as tk
from tkinter import ttk, messagebox
from tkcalendar import DateEntry
import threading

def load_json_file(filename):
    """Membaca file JSON dan mengembalikan datanya."""
    if not os.path.exists(filename):
        messagebox.showerror("File Tidak Ditemukan", f"KESALAHAN: File '{filename}' tidak ditemukan.")
        return None
    try:
        with open(filename, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, Exception) as e:
        messagebox.showerror("Gagal Membaca File", f"KESALAHAN saat membaca '{filename}': {e}")
        return None

def get_unique_filename(base_name, extension):
    """Membuat nama file unik dengan menambahkan angka jika sudah ada."""
    filename = f"{base_name}.{extension}"
    counter = 1
    while os.path.exists(filename):
        filename = f"{base_name} - {counter}.{extension}"
        counter += 1
    return filename

def process_task_data(task, master_map, real_sequence_map):
    """
    Memproses satu data 'task' dan mengekstrak semua informasi yang dibutuhkan.
    Ini untuk menghindari ekstraksi data yang berulang.
    """
    vehicle_assignee_email = (task.get('assignedVehicle') or {}).get('assignee')
    if not vehicle_assignee_email:
        return None

    master_record = master_map.get(vehicle_assignee_email, {})
    driver_name = master_record.get('Driver', vehicle_assignee_email)
    customer_name = task.get('customerName', '')

    # Time processing
    t_arrival_utc = pd.to_datetime(task.get('klikJikaAndaSudahSampai'), errors='coerce')
    t_departure_utc = pd.to_datetime(task.get('doneTime'), errors='coerce')
    
    t_arrival_local = t_arrival_utc.tz_convert('Asia/Jakarta') if pd.notna(t_arrival_utc) else pd.NaT
    t_departure_local = t_departure_utc.tz_convert('Asia/Jakarta') if pd.notna(t_departure_utc) else pd.NaT

    actual_visit_time = pd.NA
    if pd.notna(t_arrival_local) and pd.notna(t_departure_local):
        t_arrival_minute = t_arrival_local.replace(second=0, microsecond=0)
        t_departure_minute = t_departure_local.replace(second=0, microsecond=0)
        delta_minutes = (t_departure_minute - t_arrival_minute).total_seconds() / 60
        actual_visit_time = int(delta_minutes)

    et_sequence = task.get('routePlannedOrder', 0)
    real_sequence = real_sequence_map.get(task['_id'], 0)
    
    return {
        'task_id': task['_id'],
        'license_plat': (task.get('assignedVehicle') or {}).get('name', 'N/A'),
        'driver_name': driver_name,
        'assignee_email': vehicle_assignee_email,
        'customer_name': customer_name,
        'status_delivery': ', '.join(task.get('statusDelivery', [])),
        'open_time': task.get('openTime', ''),
        'close_time': task.get('closeTime', ''),
        'eta': (task.get('eta') or '')[:5],
        'etd': (task.get('etd') or '')[:5],
        'actual_arrival': t_arrival_local.strftime('%H:%M') if pd.notna(t_arrival_local) else '',
        'actual_departure': t_departure_local.strftime('%H:%M') if pd.notna(t_departure_local) else '',
        'visit_time': task.get('visitTime', ''),
        'actual_visit_time': actual_visit_time,
        'et_sequence': et_sequence,
        'real_sequence': real_sequence,
        'is_same_sequence': "SAMA" if et_sequence == real_sequence else "TIDAK SAMA",
        'labels': task.get('label', []),
        'alasan_batal': task.get('alasanBatal', ''),
        'alasan_tolakan': task.get('alasanTolakan', '')
    }

def format_excel_sheet(writer, df, sheet_name, centered_cols, colored_cols=None):
    """Menulis DataFrame ke sheet dan menerapkan format."""
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    worksheet = writer.sheets[sheet_name]
    center_align = Alignment(horizontal='center', vertical='center')

    for idx, col_name in enumerate(df.columns):
        col_letter = chr(65 + idx)
        try:
            max_len = max(df[col_name].astype(str).map(len).max(), len(col_name)) + 2
            worksheet.column_dimensions[col_letter].width = max_len
        except (ValueError, TypeError):
             worksheet.column_dimensions[col_letter].width = len(col_name) + 2

        if col_name in centered_cols:
            for cell in worksheet[col_letter][1:]:
                cell.alignment = center_align
        
        if colored_cols and col_name in colored_cols:
            fill = PatternFill(start_color=colored_cols[col_name], end_color=colored_cols[col_name], fill_type="solid")
            for cell in worksheet[col_letter]:
                cell.fill = fill

def panggil_api_dan_simpan(selected_date, app_instance):
    """
    Fungsi utama untuk memanggil API, memproses data, dan menyimpan ke Excel.
    """
    # --- PENGATURAN ---
    configs = {
        'constants': load_json_file('constant.json'),
        'config': load_json_file('config.json'),
        'master_data': load_json_file('master.json')
    }
    if any(v is None for v in configs.values()): return False

    API_TOKEN = configs['constants'].get('token') 
    LOKASI_FILTER = configs['config'].get('lokasi')
    HUB_ID = configs['constants'].get('hub_ids', {}).get(LOKASI_FILTER) 

    if not API_TOKEN or not LOKASI_FILTER or not HUB_ID:
        messagebox.showerror("Konfigurasi Salah", "KESALAHAN: 'token', 'lokasi', atau hubId tidak ditemukan di file konfigurasi.")
        return False

    base_filename = "Delivery Summary"
    NAMA_FILE_OUTPUT = get_unique_filename(base_filename, "xlsx")
    master_map = {item['Email']: item for item in configs['master_data']}
    
    # --- API Call ---
    api_url = "https://apiweb.mile.app/api/v3/tasks"
    params = {
        "status": "DONE",
        "hubId": HUB_ID,
        "timeFrom": f"{selected_date} 00:00:00",
        "timeTo": f"{selected_date} 23:59:59",
        "timeBy": "doneTime",
        "limit": 1000
    }
    headers = {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

    app_instance.update_status("🚀 Memulai pemanggilan API...")
    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=60)
        response.raise_for_status()
        tasks_data = response.json().get('tasks', {}).get('data')
        if not tasks_data:
            messagebox.showwarning("Data Kosong", f"Tidak ada data tugas yang ditemukan untuk tanggal {selected_date}.")
            return False
        app_instance.update_status(f"✅ Ditemukan total {len(tasks_data)} data tugas.")
    except requests.exceptions.HTTPError as errh:
        # PERBAIKAN: Penanganan error server dan otentikasi yang lebih baik
        status_code = errh.response.status_code
        if status_code == 401:
            messagebox.showerror("Akses Ditolak (401)", "KESALAHAN: Unauthorized. Token API mungkin salah atau sudah kedaluwarsa.")
        elif status_code >= 500: # Error 500, 502, 503, dll.
            messagebox.showerror("Masalah Server API", f"KESALAHAN: Terjadi masalah pada server API (Status Code: {status_code}). Coba lagi nanti.")
        else:
            messagebox.showerror("Kesalahan HTTP", f"KESALAHAN HTTP: {errh}")
        return False
    except requests.exceptions.ConnectionError:
        messagebox.showerror("Koneksi Gagal", "KESALAHAN: Tidak dapat terhubung ke server. Periksa koneksi internet Anda.")
        return False
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Kesalahan API", f"KESALAHAN REQUEST API: {e}")
        return False
    
    # --- Pre-processing Real Sequence ---
    tasks_by_assignee = {}
    for task in tasks_data:
        assignee_email = (task.get('assignedVehicle') or {}).get('assignee')
        if assignee_email and LOKASI_FILTER in assignee_email:
            tasks_by_assignee.setdefault(assignee_email, []).append(task)
    
    real_sequence_map = {}
    for assignee, tasks in tasks_by_assignee.items():
        sorted_tasks = sorted(tasks, key=lambda x: x.get('doneTime') or '9999-12-31T23:59:59Z')
        for i, task in enumerate(sorted_tasks):
            real_sequence_map[task['_id']] = i + 1

    # --- Data Aggregation ---
    app_instance.update_status("\n📊 Memulai agregasi data untuk laporan Excel...")
    summary_data = {email: {'License Plat': record.get('Plat', 'N/A'), 'Driver': record.get('Driver', email), 'Total Visit': pd.NA, 'Total Delivered': pd.NA} for email, record in master_map.items() if LOKASI_FILTER in email}
    pending_so_data, ro_vs_real_data = [], []
    undelivered_labels = ["PENDING", "BATAL", "TERIMA SEBAGIAN"]

    for task in tasks_data:
        processed = process_task_data(task, master_map, real_sequence_map)
        if not processed or LOKASI_FILTER not in processed['assignee_email']:
            continue
        
        # Sheet 1 Data
        if processed['assignee_email'] in summary_data:
            if pd.isna(summary_data[processed['assignee_email']]['Total Visit']):
                summary_data[processed['assignee_email']]['Total Visit'] = 0
                summary_data[processed['assignee_email']]['Total Delivered'] = 0
            summary_data[processed['assignee_email']]['Total Visit'] += 1
            if not any(label in undelivered_labels for label in processed['labels']):
                summary_data[processed['assignee_email']]['Total Delivered'] += 1

        # Sheet 2 Data
        if any(label in undelivered_labels for label in processed['labels']):
            match = re.search(r'(C0[0-9]+)', processed['customer_name'])
            reason = ''
            if "BATAL" in processed['labels']: reason = processed['alasan_batal']
            elif "TERIMA SEBAGIAN" in processed['labels']: reason = processed['alasan_tolakan']
            elif "PENDING" in processed['labels']: reason = processed['alasan_batal']
            
            pending_so_data.append({
                'License Plat': processed['license_plat'], 'Driver': processed['driver_name'],
                'Faktur Batal/ Tolakan SO': processed['customer_name'] if "BATAL" in processed['labels'] else '',
                'Terkirim Sebagian': processed['customer_name'] if "TERIMA SEBAGIAN" in processed['labels'] else '',
                'Pending': processed['customer_name'] if "PENDING" in processed['labels'] else '', 'Reason': reason,
                'Open Time': processed['open_time'], 'Close Time': processed['close_time'], 'ETA': processed['eta'], 'ETD': processed['etd'],
                'Actual Arrival': processed['actual_arrival'], 'Actual Departure': processed['actual_departure'],
                'Visit Time': processed['visit_time'], 'Actual Visit Time': processed['actual_visit_time'],
                'Customer ID': match.group(1) if match else 'N/A', 'ET Sequence': processed['et_sequence'],
                'Real Sequence': processed['real_sequence'], 'Temperature': 'DRY' if processed['driver_name'].startswith("'DRY'") else ('FRZ' if processed['driver_name'].startswith("'FRZ'") else 'N/A')
            })

        # Sheet 3 Data
        ro_vs_real_data.append({
            'License Plat': processed['license_plat'], 'Driver': processed['driver_name'], 'Customer': processed['customer_name'],
            'Status Delivery': processed['status_delivery'], 'Open Time': processed['open_time'], 'Close Time': processed['close_time'],
            'Actual Arrival': processed['actual_arrival'], 'Actual Departure': processed['actual_departure'],
            'Visit Time': processed['visit_time'], 'Actual Visit Time': processed['actual_visit_time'],
            'ET Sequence': processed['et_sequence'], 'Real Sequence': processed['real_sequence'], 'Is Same Sequence': processed['is_same_sequence']
        })

    # --- DataFrame Creation and Formatting ---
    df_delivered = pd.DataFrame(list(summary_data.values())).sort_values(by='Driver', ascending=True)
    
    df_pending = pd.DataFrame(pending_so_data)
    if not df_pending.empty:
        df_pending.insert(df_pending.columns.get_loc('Reason') + 1, ' ', '')
        df_pending = df_pending.sort_values(by='Driver', ascending=True)

    df_ro_vs_real = pd.DataFrame(ro_vs_real_data)
    if not df_ro_vs_real.empty:
        df_ro_vs_real = df_ro_vs_real.sort_values(by=['Driver', 'Real Sequence'], ascending=[True, True])
        final_ro_rows = []
        last_driver = None
        for _, row in df_ro_vs_real.iterrows():
            if last_driver is not None and row['Driver'] != last_driver:
                final_ro_rows.append({col: '' for col in df_ro_vs_real.columns})
            final_ro_rows.append(row.to_dict())
            last_driver = row['Driver']
        df_ro_vs_real = pd.DataFrame(final_ro_rows)

    # --- Excel Writing ---
    app_instance.update_status("💾 Menyimpan data ke file Excel...")
    try:
        with pd.ExcelWriter(NAMA_FILE_OUTPUT, engine='openpyxl') as writer:
            format_excel_sheet(writer, df_delivered, 'Total Delivered', centered_cols=['Total Visit', 'Total Delivered'])
            format_excel_sheet(writer, df_pending, 'Hasil Pending SO', 
                               centered_cols=['Open Time', 'Close Time', 'ETA', 'ETD', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'Customer ID', 'ET Sequence', 'Real Sequence', 'Temperature'],
                               colored_cols={' ': "FFC0CB"})
            format_excel_sheet(writer, df_ro_vs_real, 'Hasil RO vs Real', 
                               centered_cols=['Status Delivery', 'Open Time', 'Close Time', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'ET Sequence', 'Real Sequence', 'Is Same Sequence'])
        
        app_instance.update_status(f"✅ Laporan berhasil disimpan sebagai '{NAMA_FILE_OUTPUT}'")
        os.startfile(os.path.realpath(NAMA_FILE_OUTPUT))
        return True
    except Exception as e:
        messagebox.showerror("Gagal Menyimpan", f"GAGAL MENYIMPAN FILE EXCEL: {e}")
        return False

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Pilih Tanggal")
        
        window_width = 350
        window_height = 220
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        center_x = int(screen_width/2 - window_width / 2)
        center_y = int(screen_height/2 - window_height / 2)
        self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
        self.config(bg='SystemButtonFace')

        style = ttk.Style(self)
        style.theme_use('clam')
        style.configure("TButton", font=("Helvetica", 12), padding=5)
        style.configure("TLabel", background='SystemButtonFace', font=("Helvetica", 16, "bold"))
        style.configure("TProgressbar", thickness=20)
        
        main_frame = tk.Frame(self, bg='SystemButtonFace')
        main_frame.pack(expand=True, pady=20)

        label = ttk.Label(main_frame, text="Pilih Tanggal")
        label.pack(pady=(0, 10))

        # PERBAIKAN: Menggunakan style 'TEntry' agar tidak ada dropdown
        self.cal = DateEntry(main_frame, date_pattern='dd-MM-yyyy', font=("Helvetica", 16), 
                             width=12, justify='center', borderwidth=2, relief="solid",
                             style='TEntry') # Mengubah style
        self.cal.pack(pady=10, ipady=5)
        # PERBAIKAN: Bind event klik ke seluruh widget DateEntry
        self.cal.bind("<Button-1>", self._on_date_click)

        self.run_button = ttk.Button(main_frame, text="Proses", command=self.run_report_thread, style="TButton")
        self.run_button.pack(pady=10)

        self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="indeterminate")

        self.status_label = ttk.Label(main_frame, text="", foreground="blue", font=("Helvetica", 10))
        self.status_label.pack(pady=5)

    def _on_date_click(self, event):
        """Secara programatis membuka kalender dropdown saat widget diklik."""
        # DateEntry menyimpan kalender di _top_cal, kita panggil metode drop_down
        self.cal.drop_down()

    def run_report_thread(self):
        """Menjalankan proses utama di thread terpisah agar GUI tidak freeze."""
        self.run_button.pack_forget()
        self.progress.pack(pady=10)
        self.progress.start(10)
        
        selected_date_obj = self.cal.get_date()
        selected_date_str = selected_date_obj.strftime('%Y-%m-%d')
        
        self.process_thread = threading.Thread(
            target=panggil_api_dan_simpan,
            args=(selected_date_str, self)
        )
        self.process_thread.start()
        self.after(100, self.check_thread)

    def check_thread(self):
        """Memeriksa apakah thread sudah selesai."""
        if self.process_thread.is_alive():
            self.after(100, self.check_thread)
        else:
            self.progress.stop()
            self.progress.pack_forget()
            self.status_label.config(text="")
            self.run_button.pack(pady=10)

    def update_status(self, message):
        """Fungsi untuk update status label dari thread lain."""
        self.status_label.config(text=message)

def main():
    """Initializes and runs the Tkinter application."""
    app = App()
    app.mainloop()

if __name__ == "__main__":
    main()