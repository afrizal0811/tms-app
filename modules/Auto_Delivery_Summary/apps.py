import requests
import pandas as pd
import re
from openpyxl.styles import Alignment, PatternFill
from tkinter import messagebox

# Impor fungsi terpusat dari shared_utils
from ..shared_utils import (
    load_config,
    load_constants,
    load_master_data,
    open_file_externally,
    get_save_path,
    load_secret
)
from ..gui_utils import create_date_picker_window

def process_task_data(task, master_map, real_sequence_map):
    """
    Memproses satu data 'task' dan mengekstrak semua informasi yang dibutuhkan.
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

def panggil_api_dan_simpan(dates, app_instance):
    """
    Fungsi utama untuk memanggil API, memproses data, dan menyimpan ke Excel.
    """
    selected_date = dates["ymd"] 
    # --- PENGATURAN MENGGUNAKAN SHARED UTILS ---
    constants = load_constants()
    config = load_config()
    master_df = load_master_data()
    secrets = load_secret()

    if any(v is None for v in [constants, config, master_df, secrets]):
        return False
        
    master_data_list = master_df.to_dict('records')

    API_TOKEN = secrets.get('token') # Ambil token dari secrets
    LOKASI_FILTER = config.get('lokasi')
    HUB_ID = constants.get('hub_ids', {}).get(LOKASI_FILTER)
    LOKASI_MAPPING = constants.get('lokasi_mapping', {})

    if not API_TOKEN or API_TOKEN == "PASTE_YOUR_MILEAPP_TOKEN_HERE":
        messagebox.showerror("Error Token API", "Token API belum diatur di secret.json.")
        return False
    if not LOKASI_FILTER or not HUB_ID:
        messagebox.showerror("Konfigurasi Salah", "KESALAHAN: 'lokasi' atau hubId tidak ditemukan di file konfigurasi.")
        return False

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
        status_code = errh.response.status_code
        if status_code == 401:
            messagebox.showerror("Akses Ditolak (401)", "KESALAHAN: Unauthorized. Token API mungkin salah atau sudah kedaluwarsa.")
        elif status_code >= 500:
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
    
    # --- Data Processing (No changes in this part) ---
    master_map = {item['Email']: item for item in master_data_list}
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

    app_instance.update_status("\n📊 Memulai agregasi data untuk laporan Excel...")
    summary_data = {email: {'License Plat': record.get('Plat', 'N/A'), 'Driver': record.get('Driver', email), 'Total Visit': pd.NA, 'Total Delivered': pd.NA} for email, record in master_map.items() if LOKASI_FILTER in email}
    pending_so_data, ro_vs_real_data = [], []
    undelivered_labels = ["PENDING", "BATAL", "TERIMA SEBAGIAN"]

    for task in tasks_data:
        processed = process_task_data(task, master_map, real_sequence_map)
        if not processed or LOKASI_FILTER not in processed['assignee_email']:
            continue
        
        if processed['assignee_email'] in summary_data:
            if pd.isna(summary_data[processed['assignee_email']]['Total Visit']):
                summary_data[processed['assignee_email']]['Total Visit'] = 0
                summary_data[processed['assignee_email']]['Total Delivered'] = 0
            summary_data[processed['assignee_email']]['Total Visit'] += 1
            if not any(label in undelivered_labels for label in processed['labels']):
                summary_data[processed['assignee_email']]['Total Delivered'] += 1

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

        ro_vs_real_data.append({
            'License Plat': processed['license_plat'], 'Driver': processed['driver_name'], 'Customer': processed['customer_name'],
            'Status Delivery': processed['status_delivery'], 'Open Time': processed['open_time'], 'Close Time': processed['close_time'],
            'Actual Arrival': processed['actual_arrival'], 'Actual Departure': processed['actual_departure'],
            'Visit Time': processed['visit_time'], 'Actual Visit Time': processed['actual_visit_time'],
            'ET Sequence': processed['et_sequence'], 'Real Sequence': processed['real_sequence'], 'Is Same Sequence': processed['is_same_sequence']
        })

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
    app_instance.update_status("💾 Meminta lokasi penyimpanan file...")
    
    # Mendapatkan nama lokasi dari mapping
    lokasi_name = next((name for name, code in LOKASI_MAPPING.items() if code == LOKASI_FILTER), LOKASI_FILTER)
    selected_date_for_filename = dates["dmy"].replace("-", ".")
    base_name = f"Delivery Summary {lokasi_name} - {selected_date_for_filename}"
    
    NAMA_FILE_OUTPUT = get_save_path(base_name)
    
    if not NAMA_FILE_OUTPUT:
        messagebox.showwarning("Dibatalkan", "Proses penyimpanan dibatalkan oleh pengguna.")
        return False
        
    try:
        with pd.ExcelWriter(NAMA_FILE_OUTPUT, engine='openpyxl') as writer:
            format_excel_sheet(writer, df_delivered, 'Total Delivered', centered_cols=['Total Visit', 'Total Delivered'])
            format_excel_sheet(writer, df_pending, 'Hasil Pending SO', 
                                centered_cols=['Open Time', 'Close Time', 'ETA', 'ETD', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'Customer ID', 'ET Sequence', 'Real Sequence', 'Temperature'],
                                colored_cols={' ': "FFC0CB"})
            format_excel_sheet(writer, df_ro_vs_real, 'Hasil RO vs Real', 
                                centered_cols=['Status Delivery', 'Open Time', 'Close Time', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'ET Sequence', 'Real Sequence', 'Is Same Sequence'])
        
        open_file_externally(NAMA_FILE_OUTPUT)
        return True
    except Exception as e:
        messagebox.showerror("Gagal Menyimpan", f"GAGAL MENYIMPAN FILE EXCEL: {e}")
        return False

def main():
    """Fungsi utama untuk modul Auto Delivery Summary."""
    def process_wrapper(dates, app_instance):
        
        # Buat fungsi untuk menutup GUI dengan aman
        def safe_close():
            # Pastikan window masih ada sebelum melakukan apapun
            if app_instance and app_instance.winfo_exists():
                # 1. Hentikan animasi progress bar terlebih dahulu
                app_instance.progress.stop()
                # 2. Baru hancurkan window
                app_instance.destroy()

        try:
            # Panggil fungsi pemrosesan utama seperti biasa
            panggil_api_dan_simpan(dates, app_instance)
        finally:
            # Apapun hasilnya, jadwalkan fungsi penutupan aman 
            # untuk berjalan di thread utama GUI.
            if app_instance and app_instance.winfo_exists():
                app_instance.after(100, safe_close) # Diberi jeda sedikit (100ms)
                
    create_date_picker_window("Auto Delivery Summary", process_wrapper)

if __name__ == "__main__":
    main()