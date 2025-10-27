from openpyxl.styles import Alignment, PatternFill, Font
import pandas as pd
import numpy as np # Import numpy for sorting
import re
import requests
import traceback
import math
from openpyxl.comments import Comment
from utils.function import (
    get_save_path,
    load_config,
    load_constants,
    load_master_data,
    load_secret,
    open_file_externally
)
from utils.gui import create_date_picker_window
from utils.function import show_error_message, show_info_message
from utils.messages import ERROR_MESSAGES, INFO_MESSAGES
from utils.api_handler import handle_requests_error

def haversine_distance(coord1_str, coord2_str):
    """
    Menghitung jarak Haversine antara dua koordinat (lat, long) dalam meter.
    Koordinat harus dalam format string "lat,long".
    Mengembalikan 0 jika salah satu koordinat tidak valid atau kosong.
    """
    if not coord1_str or not coord2_str:
        return 0

    try:
        lat1, lon1 = map(float, coord1_str.split(','))
        lat2, lon2 = map(float, coord2_str.split(','))
    except ValueError:
        return 0

    R = 6371000
    lat1_rad, lon1_rad, lat2_rad, lon2_rad = map(math.radians, [lat1, lon1, lat2, lon2])
    dlon = lon2_rad - lon1_rad
    dlat = lat2_rad - lat1_rad
    a = math.sin(dlat / 2)**2 + math.cos(lat1_rad) * math.cos(lat2_rad) * math.sin(dlon / 2)**2
    c = 2 * math.atan2(math.sqrt(a), math.sqrt(1 - a))
    distance = R * c
    return round(distance)

# =============================================================================
# ▼▼▼ process_task_data DARI KODE 1 (UNTUK Total Delivered, Pending SO) ▼▼▼
# =============================================================================
# =============================================================================
# ▼▼▼ process_task_data DARI KODE 1 (UNTUK Total Delivered, Pending SO) ▼▼▼
# =============================================================================
def process_task_data_code1(task, master_map, real_sequence_map):
    """
    Memproses satu data 'task' (Logic from Code 1).
    Returns None if essential assignee info is missing.
    """
    
    # --- PERBAIKAN: Gunakan assignedTo SEBAGAI FALLBACK ---
    vehicle_assignee_email = (task.get('assignedVehicle') or {}).get('assignee')
    assigned_to_data = task.get('assignedTo') or {}
    assigned_to_email = assigned_to_data.get('email')
    assigned_to_name = assigned_to_data.get('name') # Nama driver dari assignedTo

    # Tentukan email utama (prioritas vehicle, fallback ke assignedTo)
    final_assignee_email = vehicle_assignee_email or assigned_to_email
    
    # Filter Awal: Jika tidak ada email assignee SAMA SEKALI, skip task ini
    if not final_assignee_email:
        return None

    master_record = master_map.get(final_assignee_email, {})
    
    # Tentukan nama driver (prioritas assignedTo.name, fallback ke master, fallback ke email)
    driver_name = assigned_to_name or master_record.get('Driver', final_assignee_email)
    
    # Tentukan plat (prioritas assignedVehicle.name, fallback ke master)
    license_plat = (task.get('assignedVehicle') or {}).get('name')
    if not license_plat or license_plat == 'N/A':
        license_plat = master_record.get('Plat', 'N/A')
    # --- AKHIR PERBAIKAN ---

    customer_name = task.get('customerName', '')

    t_arrival_utc = pd.to_datetime(task.get('klikJikaSudahSampai'), errors='coerce')
    # Gunakan 'doneTime' sebagai waktu departure untuk konsistensi
    t_departure_utc = pd.to_datetime(task.get('doneTime'), errors='coerce')
    t_arrival_local = t_arrival_utc.tz_convert('Asia/Jakarta') if pd.notna(t_arrival_utc) else pd.NaT
    t_departure_local = t_departure_utc.tz_convert('Asia/Jakarta') if pd.notna(t_departure_utc) else pd.NaT

    actual_visit_time = pd.NA
    if pd.notna(t_arrival_local) and pd.notna(t_departure_local):
        delta_minutes = (t_departure_local - t_arrival_local).total_seconds() / 60
        actual_visit_time = int(round(delta_minutes))

    et_sequence = task.get('routePlannedOrder') # Bisa None
    et_sequence_val = et_sequence if et_sequence is not None else 0 # Default ke 0 jika None

    real_sequence = real_sequence_map.get(task['_id'], 0) # Real sequence dari map (Code 1 logic)

    # --- PERBAIKAN (dari chat sebelumnya): Gabungkan SEMUA sumber status ---
    status_delivery_1 = task.get('statusDelivery', [])
    status_delivery_2 = task.get('statusGr', [])
    raw_labels = task.get('label')

    # Pastikan status_delivery_1 & 2 adalah list (API bisa saja 'None')
    if not isinstance(status_delivery_1, list): status_delivery_1 = []
    if not isinstance(status_delivery_2, list): status_delivery_2 = []
    
    # Gabungkan list awal
    combined_list = status_delivery_1 + status_delivery_2
    
    # Tambahkan 'label' ke combined_list
    if isinstance(raw_labels, str):
        combined_list.append(raw_labels) # Tambahkan string
    elif isinstance(raw_labels, list):
        combined_list.extend(raw_labels) # Tambahkan list
        
    # Buat list unik (dan hapus string kosong/None)
    final_status_list = list(dict.fromkeys(s for s in combined_list if s)) 
    
    # Tetapkan hasil ke *kedua* variabel untuk konsistensi
    labels_list = final_status_list
    combined_status_delivery = final_status_list
    # --- AKHIR PERBAIKAN ---

    return {
        'task_id': task['_id'],
        'flow': task.get('flow', ''),
        'license_plat': license_plat, # Variabel baru
        'driver_name': driver_name, # Variabel baru
        'assignee_email': final_assignee_email, # Variabel baru
        'customer_name': customer_name,
        'status_delivery': ', '.join(combined_status_delivery), # Ini tetap dipakai Pending SO
        'status_delivery_list': combined_status_delivery, # Ini tetap dipakai Pending SO
        'open_time': task.get('openTime', ''),
        'close_time': task.get('closeTime', ''),
        'eta': (task.get('eta') or '')[:5],
        'etd': (task.get('etd') or '')[:5],
        'actual_arrival': t_arrival_local.strftime('%H:%M') if pd.notna(t_arrival_local) else '',
        'actual_departure': t_departure_local.strftime('%H:%M') if pd.notna(t_departure_local) else '',
        'visit_time': task.get('visitTime', ''),
        'actual_visit_time': actual_visit_time,
        'et_sequence': et_sequence_val,
        'real_sequence': real_sequence,
        'is_same_sequence': "SAMA" if et_sequence is not None and et_sequence_val == real_sequence else "TIDAK SAMA",
        'labels': labels_list,
        'alasan': task.get('alasan', ''),
        # Tambahan untuk perhitungan Real Sequence & Actual Visit Time di RO vs Real
        '_arrival_dt_local': t_arrival_local,
        '_departure_dt_local': t_departure_local,
        # Tambahkan _arrival_utc untuk sorting RO vs Real di Code 2 logic
        '_arrival_utc': t_arrival_utc
    }
# =============================================================================

def format_excel_sheet(writer, df, sheet_name, centered_cols, colored_cols=None):
    """Menulis DataFrame ke sheet dan menerapkan format."""
    df.to_excel(writer, index=False, sheet_name=sheet_name)
    workbook = writer.book
    worksheet = writer.sheets[sheet_name]
    center_align = Alignment(horizontal='center', vertical='center')

    for idx, col_name in enumerate(df.columns):
        col_letter = chr(65 + idx)
        try:
            # Hitung panjang data dan header
            data_lengths = df[col_name].astype(str).map(len)
            max_len_data = data_lengths.max() if not data_lengths.empty else 0
            header_len = len(str(col_name))
            max_len = max(max_len_data, header_len) + 2
            worksheet.column_dimensions[col_letter].width = max_len
        except (ValueError, TypeError, AttributeError): # Tangani error
             worksheet.column_dimensions[col_letter].width = len(str(col_name)) + 2 # Fallback

        if col_name in centered_cols:
            for cell in worksheet[col_letter][1:]: # Mulai dari baris ke-2
                if cell.value is not None and cell.value != '':
                    cell.alignment = center_align

        # --- PERBAIKAN DI BLOK INI ---
        if colored_cols and col_name in colored_cols:
            fill = PatternFill(start_color=colored_cols[col_name], end_color=colored_cols[col_name], fill_type="solid")
            for cell in worksheet[col_letter]:
                 # Kondisi 'if cell.value' dihapus agar seluruh kolom diwarnai
                 cell.fill = fill
        # --- AKHIR PERBAIKAN ---

    header_align = Alignment(horizontal='center', vertical='center')
    for cell in worksheet[1]: # Hanya header
        cell.alignment = header_align

def panggil_api_dan_simpan(dates, app_instance):
    """
    Fungsi utama untuk memanggil API, memproses data, dan menyimpan ke Excel.
    """
    selected_date = dates["ymd"]
    # --- PENGATURAN MENGGUNAKAN SHARED UTILS ---
    constants = load_constants()
    config = load_config()
    secrets = load_secret()
    master_data = load_master_data()

    master_df = master_data["df"]
    hub_ids = master_data["hub_ids"]

    if not constants: show_error_message("Gagal", ERROR_MESSAGES["CONSTANT_FILE_ERROR"]); return False
    if not config: show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"]); return False
    if not secrets: show_error_message("Gagal", ERROR_MESSAGES["SECRET_FILE_ERROR"]); return False
    if master_data is None: show_error_message("Gagal", ERROR_MESSAGES["MASTER_DATA_MISSING"]); return False

    master_map = {row['Email']: row for _, row in master_df.iterrows()}

    API_TOKEN = secrets.get('token')
    LOKASI_FILTER = config.get('lokasi') # Digunakan untuk filter master & real sequence
    HUB_ID = hub_ids.get(LOKASI_FILTER)
    location_id = constants.get('location_id', {})
    show_pending_gr = LOKASI_FILTER in ["plck", "pldm"]

    if not API_TOKEN: show_error_message("Error Token API", ERROR_MESSAGES["API_TOKEN_MISSING"]); return False
    if not LOKASI_FILTER or not HUB_ID: show_error_message("Konfigurasi Salah", ERROR_MESSAGES["HUB_ID_MISSING"]); return False

    base_url = constants.get('base_url')
    api_url = f"{base_url}/tasks"
    params = {
        "status": "DONE", "hubId": HUB_ID,
        "timeFrom": f"{selected_date} 00:00:00", "timeTo": f"{selected_date} 23:59:59",
        "timeBy": "doneTime", "limit": 1000
    }
    headers = {"Authorization": f"Bearer {API_TOKEN}", "Content-Type": "application/json"}

    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=60)
        response.raise_for_status()
        tasks_data = response.json().get('tasks', {}).get('data')

        if not tasks_data:
            show_error_message("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
            return False

    except requests.exceptions.RequestException as e: handle_requests_error(e); return False
    except Exception as e: show_error_message("Error API", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=f"{e}\n\n{traceback.format_exc()}")); return False

    # --- Data Processing ---
    # real_sequence map tetap perlu filter lokasi email dari assignedVehicle
    tasks_by_assignee_for_seq = {}
    for task in tasks_data:
        assignee_email_vehicle = (task.get('assignedVehicle') or {}).get('assignee')
        if assignee_email_vehicle and LOKASI_FILTER in assignee_email_vehicle:
            tasks_by_assignee_for_seq.setdefault(assignee_email_vehicle, []).append(task)

    real_sequence_map = {}
    for assignee, tasks in tasks_by_assignee_for_seq.items():
        # Urutkan berdasarkan doneTime untuk real sequence (sesuai kode asli)
        sorted_tasks = sorted(tasks, key=lambda x: x.get('doneTime') or '9999-12-31T23:59:59Z')
        for i, task in enumerate(sorted_tasks):
            real_sequence_map[task['_id']] = i + 1

    # =====================================================================
    # ▼▼▼ LOGIKA PENGOLAHAN DATA UNTUK SEMUA SHEET ▼▼▼
    # =====================================================================
    # Inisialisasi dictionary untuk agregasi 'Total Delivered' (pakai assignedTo.name - Logic Code 2)
    summary_data_total_delivered_new = {}

    # Inisialisasi list untuk data mentah 'Hasil RO vs Real' (berdasarkan assignedTo.name - Logic Code 2)
    ro_vs_real_raw_data = {} # Pakai dict untuk grouping by assignedTo.name

    processed_tasks_list_pending = [] 
    
    pending_undelivered_labels = ["PENDING", "BATAL", "TERIMA SEBAGIAN", "PENDING GR"]
    
    if show_pending_gr:
        pending_undelivered_labels.append("PENDING GR")

    # Iterasi utama: Proses data untuk semua sheet
    for task in tasks_data:

        # --- Proses untuk 'Total Delivered' (Filter Baru by assignedTo.name - Logic Code 2) ---
        assigned_to_data = task.get('assignedTo')
        driver_name_from_assigned_to = None
        driver_email_from_assigned_to = None

        if isinstance(assigned_to_data, dict):
            driver_name_from_assigned_to = assigned_to_data.get('name')
            driver_email_from_assigned_to = assigned_to_data.get('email')

        # Hanya proses jika 'assignedTo.name' ada
        if driver_name_from_assigned_to:
            if driver_name_from_assigned_to not in summary_data_total_delivered_new:
                # Cari plat berdasarkan email jika ada, jika tidak, coba cari dari assignedVehicle
                plat_td = "N/A_Plat"
                if driver_email_from_assigned_to and driver_email_from_assigned_to in master_map:
                    plat_td = master_map[driver_email_from_assigned_to].get('Plat', 'N/A_Plat')
                elif task.get('assignedVehicle') and isinstance(task['assignedVehicle'], dict):
                     plat_td = task['assignedVehicle'].get('name', 'N/A_Plat')

                summary_data_total_delivered_new[driver_name_from_assigned_to] = {
                    'License Plat': plat_td,
                    'Driver': driver_name_from_assigned_to,
                    'Total Outlet': 0,
                    'Total Delivered': 0
                }

            # Hitung Total Outlet
            summary_data_total_delivered_new[driver_name_from_assigned_to]['Total Outlet'] += 1

            # Hitung Total Delivered
            # --- PERBAIKAN: Tangani 'label' string atau list ---
            raw_labels_td = task.get('label')
            labels_td_list = []
            if isinstance(raw_labels_td, str):
                labels_td_list = [raw_labels_td]
            elif isinstance(raw_labels_td, list):
                labels_td_list = raw_labels_td
            # --- AKHIR PERBAIKAN ---
            
            failure_labels_td = ["PENDING", "BATAL", "TERIMA SEBAGIAN"]
            if not show_pending_gr: # Jika BUKAN plck/pldm
                failure_labels_td.append("PENDING GR") # Anggap PENDING GR sebagai gagal
                
                # Gunakan list yang sudah bersih
            is_pending_or_batal_td = any(label in failure_labels_td for label in labels_td_list)
            if not is_pending_or_batal_td:
                summary_data_total_delivered_new[driver_name_from_assigned_to]['Total Delivered'] += 1


            # --- Proses untuk 'Hasil RO vs Real' (Filter Baru by assignedTo.name - Logic Code 2) ---
            plat_ro = "N/A_Plat"
            # Prioritaskan lookup plat dari master berdasarkan email assignedTo
            if driver_email_from_assigned_to and driver_email_from_assigned_to in master_map:
                plat_ro = master_map[driver_email_from_assigned_to].get('Plat', 'N/A_Plat')
            # Jika tidak ada di master, coba ambil dari assignedVehicle
            elif task.get('assignedVehicle') and isinstance(task['assignedVehicle'], dict):
                 plat_ro = task['assignedVehicle'].get('name', 'N/A_Plat')

            # Ekstrak data yang dibutuhkan untuk RO vs Real
            customer_name_ro = task.get('customerName', '')
            # --- PERBAIKAN: Tangani 'label' string atau list ---
            raw_labels_ro = task.get('label') # Ambil dari key 'label'
            if isinstance(raw_labels_ro, str):
                status_delivery_ro = raw_labels_ro # Jika string, gunakan langsung
            elif isinstance(raw_labels_ro, list):
                status_delivery_ro = ', '.join(raw_labels_ro) # Jika list, gabungkan
            else:
                status_delivery_ro = '' # Jika None atau tipe lain
            # --- AKHIR PERBAIKAN ---
            
            open_time_ro = task.get('openTime', '')
            close_time_ro = task.get('closeTime', '')

            # --- PERUBAHAN ---
            # Ambil 'flow' terlebih dahulu untuk menentukan 'Actual Arrival'
            flow_ro = task.get('flow', '') # Ambil nilai 'flow'
            
            # Tentukan key untuk 'Actual Arrival' berdasarkan 'flow'
            arrival_key = 'page1DoneTime' if 'Pending GR' in flow_ro else 'klikJikaSudahSampai'
            arrival_utc_ro = pd.to_datetime(task.get(arrival_key), errors='coerce')
            # --- AKHIR PERUBAHAN ---

            departure_utc_ro = pd.to_datetime(task.get('doneTime'), errors='coerce')
            visit_time_api_ro = task.get('visitTime', '')
            ro_sequence_ro = task.get('routePlannedOrder') # Bisa None

            task_details_for_ro = {
                '_task_id': task['_id'],
                'Flow': flow_ro, # Tambahkan kolom Flow
                'Plat': plat_ro,
                'Driver': driver_name_from_assigned_to,
                'Customer': customer_name_ro,
                'Status Delivery': status_delivery_ro, # Gunakan data label
                'Open Time': open_time_ro,
                'Close Time': close_time_ro,
                'Visit Time': visit_time_api_ro,
                'RO Sequence': ro_sequence_ro if ro_sequence_ro is not None else '-',
                '_arrival_utc': arrival_utc_ro,
                '_departure_utc': departure_utc_ro
            }

            # Kelompokkan berdasarkan driver name (dari assignedTo.name)
            if driver_name_from_assigned_to not in ro_vs_real_raw_data:
                ro_vs_real_raw_data[driver_name_from_assigned_to] = []
            ro_vs_real_raw_data[driver_name_from_assigned_to].append(task_details_for_ro)


        # --- Proses untuk 'Pending SO' (Filter Lama by assignedVehicle.assignee - Logic Code 1) ---
        processed_code1 = process_task_data_code1(task, master_map, real_sequence_map) # Panggil fungsi asli (Code 1)
        # Hanya proses lebih lanjut jika processed_code1 valid dan email cocok lokasi
        if processed_code1 and LOKASI_FILTER in processed_code1['assignee_email']:
            # Simpan hasil proses ini HANYA untuk sheet Pending
            if (
                any(label in pending_undelivered_labels for label in processed_code1['labels'])
                or any(status in pending_undelivered_labels for status in processed_code1.get('status_delivery_list', []))
            ):
                 processed_tasks_list_pending.append(processed_code1)


    # --- Finalisasi Sheet 'Total Delivered' (Logic Code 2) ---
    df_delivered = pd.DataFrame(list(summary_data_total_delivered_new.values()))
    # Pengurutan SEWA untuk 'Total Delivered'
    if not df_delivered.empty:
        df_delivered['is_sewa'] = (df_delivered['License Plat'].astype(str).str.contains('SEWA', case=False, na=False) | df_delivered['Driver'].astype(str).str.contains('SEWA', case=False, na=False)).astype(int)
        conditions = [df_delivered['Driver'].astype(str).str.contains('DRY', case=False, na=False), df_delivered['Driver'].astype(str).str.contains('FRZ', case=False, na=False)]
        choices = [1, 2]
        df_delivered['sewa_category'] = np.select(conditions, choices, default=3)
        df_delivered = df_delivered.sort_values(by=['is_sewa', 'sewa_category', 'Driver'], ascending=[True, True, True]).reset_index(drop=True)
        df_delivered = df_delivered.drop(columns=['is_sewa', 'sewa_category'])

    ro_vs_real_final_list = []
    correct_sequence_map = {}

    for driver_name, tasks_list in ro_vs_real_raw_data.items():
        # Sortir task per driver berdasarkan waktu arrival UTC
        sorted_tasks = sorted(tasks_list, key=lambda x: x['_arrival_utc'] if pd.notna(x['_arrival_utc']) else pd.Timestamp.max.tz_localize('UTC'))

        for i, task_detail in enumerate(sorted_tasks):
            real_sequence = i + 1
            arrival_local = task_detail['_arrival_utc'].tz_convert('Asia/Jakarta') if pd.notna(task_detail['_arrival_utc']) else pd.NaT
            departure_local = task_detail['_departure_utc'].tz_convert('Asia/Jakarta') if pd.notna(task_detail['_departure_utc']) else pd.NaT

            actual_visit_time_ro = pd.NA
            if pd.notna(arrival_local) and pd.notna(departure_local):
                 delta_minutes = (departure_local - arrival_local).total_seconds() / 60
                 actual_visit_time_ro = int(round(delta_minutes))

            ro_sequence_val = task_detail['RO Sequence']
            is_same = "SAMA" if ro_sequence_val != '-' and pd.to_numeric(ro_sequence_val, errors='coerce') == real_sequence else "TIDAK SAMA"

            task_id = task_detail.get('_task_id')
            if task_id:
                correct_sequence_map[task_id] = {
                    'ro': ro_sequence_val,
                    'real': real_sequence
                }

            ro_vs_real_final_list.append({
                'Flow': task_detail['Flow'], # Tambahkan Flow
                'Plat': task_detail['Plat'],
                'Driver': task_detail['Driver'],
                'Customer': task_detail['Customer'],
                'Status Delivery': task_detail['Status Delivery'], # Sudah dari label
                'Open Time': task_detail['Open Time'],
                'Close Time': task_detail['Close Time'],
                'Actual Arrival': arrival_local.strftime('%H:%M') if pd.notna(arrival_local) else '',
                'Actual Departure': departure_local.strftime('%H:%M') if pd.notna(departure_local) else '',
                'Visit Time': task_detail['Visit Time'],
                'Actual Visit Time': actual_visit_time_ro,
                'RO Sequence': ro_sequence_val,
                'Real Sequence': real_sequence,
                'Is Same Sequence': is_same
            })

    df_ro_vs_real = pd.DataFrame(ro_vs_real_final_list)
    if not df_ro_vs_real.empty:
        # Sortir keseluruhan berdasarkan Driver, lalu Real Sequence
        df_ro_vs_real = df_ro_vs_real.sort_values(by=['Driver', 'Real Sequence'], ascending=[True, True])
        final_ro_rows_formatted = []
        last_driver = None
        # Definisikan urutan kolom yang benar di sini
        ro_cols_order = ['Flow', 'Plat', 'Driver', 'Customer', 'Status Delivery', 'Open Time', 'Close Time', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'RO Sequence', 'Real Sequence', 'Is Same Sequence']
        df_ro_vs_real = df_ro_vs_real[ro_cols_order] # Susun ulang kolom DataFrame

        for _, row in df_ro_vs_real.iterrows():
            if last_driver is not None and row['Driver'] != last_driver:
                # Sisipkan baris kosong
                final_ro_rows_formatted.append({col: '' for col in ro_cols_order}) # Gunakan urutan kolom
            final_ro_rows_formatted.append(row.to_dict())
            last_driver = row['Driver']
        df_ro_vs_real = pd.DataFrame(final_ro_rows_formatted) # Buat ulang DataFrame


# --- Finalisasi Sheet 'Hasil Pending SO' (dari processed_tasks_list_pending - Logic Code 1) ---
    
    # --- ▼▼▼ PERBAIKAN UNTUK UnboundLocalError ▼▼▼ ---
    pending_so_data = [] # Re-inisialisasi
    fill_values = None   # Inisialisasi fill_values di sini
    df_pending = pd.DataFrame() # Inisialisasi df_pending di sini
    # --- ▲▲▲ AKHIR PERBAIKAN ▲▲▲ ---

    for processed in processed_tasks_list_pending: # Gunakan list ini
        # (Filter tidak perlu diubah)
        labels_list = processed.get('labels', []) + processed.get('status_delivery_list', [])
        if not any(label in pending_undelivered_labels for label in labels_list):
            continue

        # (Logika sequence tidak perlu diubah)
        task_id = processed['task_id']
        correct_seqs = correct_sequence_map.get(task_id, {})
        ro_sequence = correct_seqs.get('ro', processed['et_sequence'])
        real_sequence = correct_seqs.get('real', processed['real_sequence'])
        
        match = re.search(r'(C0[0-9]+)', processed['customer_name'])
        
        # (Logika Reason tidak perlu diubah)
        reason = '' 
        if any(label in pending_undelivered_labels for label in labels_list):
            reason = processed['alasan']

        # (Logika pewarnaan tidak perlu diubah)
        is_pending_gr = "PENDING GR" in labels_list
        is_pending = "PENDING" in labels_list
        is_batal = "BATAL" in labels_list
        is_sebagian = "TERIMA SEBAGIAN" in labels_list
        fill_color = None
        is_redirected_gr = (is_pending_gr and not show_pending_gr)
        if is_redirected_gr:
            fill_color = "FF0000" 
        should_be_in_pending_col = (is_pending and not is_pending_gr) or is_redirected_gr
        
        # (Logika pending_row tidak perlu diubah)
        pending_row = {
            'Flow': processed['flow'],
            'License Plat': processed['license_plat'], 'Driver': processed['driver_name'],
            'Faktur Batal/ Tolakan SO': processed['customer_name'] if is_batal else '',
            'Terkirim Sebagian': processed['customer_name'] if is_sebagian else '',
            'Pending': processed['customer_name'] if should_be_in_pending_col else '',
            'Reason': reason,
            'Open Time': processed['open_time'], 'Close Time': processed['close_time'],
            'ETA': processed['eta'], 'ETD': processed['etd'], 'Actual Arrival': processed['actual_arrival'],
            'Actual Departure': processed['actual_departure'], 'Visit Time': processed['visit_time'],
            'Actual Visit Time': processed['actual_visit_time'], 'Customer ID': match.group(1) if match else 'N/A',
            'RO Sequence': ro_sequence,
            'Real Sequence': real_sequence,
            'Temperature': ('DRY' if processed['driver_name'].startswith("'DRY'") else 'FRZ' if processed['driver_name'].startswith("'FRZ'") else 'N/A'),
            '_fill': fill_color 
        }

        if show_pending_gr:
            pending_row['Pending GR'] = processed['customer_name'] if is_pending_gr else ''

        pending_so_data.append(pending_row)

    # --- ▼▼▼ PERBAIKAN LOGIKA SETELAH LOOP ▼▼▼ ---
    
    # 1. Buat DataFrame dari data (jika ada)
    if pending_so_data:
        df_pending = pd.DataFrame(pending_so_data)

    # 2. Lakukan sorting, HANYA jika DataFrame tidak kosong
    if not df_pending.empty:
        # Lakukan sorting DULU
        df_pending = df_pending.sort_values(by='Driver', ascending=True)

        # Ekstrak data warna (fill_values) SETELAH di-sort
        if '_fill' in df_pending.columns:
            fill_values = df_pending['_fill']

        # Siapkan daftar kolom untuk Excel (tanpa _fill)
        cols = list(df_pending.columns)
        if '_fill' in cols:
            cols.remove('_fill')
        
        # Atur ulang 'Pending GR' (jika ada)
        if 'Pending GR' in cols and 'Pending' in cols:
            cols.insert(cols.index('Pending') + 1, cols.pop(cols.index('Pending GR')))
            
        # Sisipkan kolom kosong ' ' (jika 'Reason' ada)
        if 'Reason' in cols:
            df_pending[' '] = '' # Tambahkan kolom ke DataFrame
            cols.insert(cols.index('Reason') + 1, ' ') # Atur posisinya di daftar
            
        # Terapkan urutan kolom ('cols') yang sudah benar
        df_pending = df_pending[cols] 
    
    update_longlat_data = []
    for task in tasks_data: # Iterasi tasks_data asli
        new_longlat = task.get('klikLokasiClient', '')
        old_longlat = task.get('longlat', '')
        if not new_longlat: continue
        beda_jarak = haversine_distance(old_longlat, new_longlat)
        title = task.get('title', '')
        match_id = re.search(r'C0\d{6,}', title)
        customer_id = match_id.group(0) if match_id else 'N/A'
        customer_name = title.split(' - ')[0].strip()
        parts = title.split(' - ')
        location_code_longlat = parts[-1].strip() if len(parts) > 2 else 'N/A'
        update_longlat_data.append({
            'Customer ID': customer_id, 'Customer Name': customer_name,
            'Location ID': location_code_longlat, 'New Longlat': new_longlat,
            'Beda Jarak (m)': beda_jarak
        })

    if update_longlat_data:
        df_longlat = pd.DataFrame(update_longlat_data)
        df_longlat = df_longlat.sort_values(by='Beda Jarak (m)', ascending=True)
    else:
        df_longlat = pd.DataFrame({"": ["Tidak Ada Update Longlat"]})

    # --- Simpan ke Excel ---
    lokasi_name = next((name for name, code in location_id.items() if code == LOKASI_FILTER), LOKASI_FILTER)
    selected_date_for_filename = dates["dmy"].replace("-", ".")
    base_name = f"Delivery Summary - {selected_date_for_filename} - {lokasi_name}"
    NAMA_FILE_OUTPUT = get_save_path(base_name)

    if not NAMA_FILE_OUTPUT: show_info_message("Dibatalkan", INFO_MESSAGES["CANCELED_BY_USER"]); return False

    try:
        with pd.ExcelWriter(NAMA_FILE_OUTPUT, engine='openpyxl') as writer:
            # Sheet Total Delivered (Logika agregasi baru - Code 2)
            if not df_delivered.empty:
                 format_excel_sheet(writer, df_delivered, 'Total Delivered', centered_cols=['Total Outlet', 'Total Delivered'])
            else:
                 pd.DataFrame([{" ": "Tidak ada data kunjungan valid (filter Code 2)"}]) \
                   .to_excel(writer, sheet_name='Total Delivered', index=False)

            # Sheet Hasil Pending SO (Logika Code 1)
            if not df_pending.empty:
                # --- KASUS JIKA ADA DATA PENDING ---
                pending_centered_cols = ['Flow', 'Open Time', 'Close Time', 'ETA', 'ETD', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'Customer ID', 'RO Sequence', 'Real Sequence', 'Temperature']
                format_excel_sheet(writer, df_pending, 'Hasil Pending SO', centered_cols=pending_centered_cols, colored_cols={' ': "FFC0CB"})

                # --- ▼▼▼ KODE UNTUK WARNA SEL (REQ 1) ▼▼▼ ---
                if fill_values is not None:
                    ws_pending = writer.sheets["Hasil Pending SO"]
                    bright_red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                    
                    # Cari kolom "Pending"
                    pending_col_idx = None
                    for cell in ws_pending[1]: # Loop header (baris 1)
                        if cell.value == "Pending":
                            pending_col_idx = cell.column # Ambil nomor kolom (mis: 5)
                            break
                    
                    if pending_col_idx is not None:
                        # Iterate data fill_values (yang sudah di-sort)
                        for i, fill_val in enumerate(fill_values.values, start=2):
                            if pd.notna(fill_val): # Jika nilainya bukan None (ada warna)
                                ws_pending.cell(row=i, column=pending_col_idx).fill = bright_red_fill
                # --- ▲▲▲ AKHIR KODE WARNA ▲▲▲ ---
            
            else:
                # --- KASUS JIKA TIDAK ADA DATA PENDING ---
                df_placeholder = pd.DataFrame({"Semua Pengiriman Sukses": []}) # Buat df dgn header saja
                df_placeholder.to_excel(writer, index=False, sheet_name='Hasil Pending SO')
                
                # Ambil worksheet-nya
                ws_pending = writer.sheets['Hasil Pending SO']
                
                # Terapkan bold ke A1
                ws_pending['A1'].font = Font(bold=True)
                
                # Sesuaikan lebar kolom A
                ws_pending.column_dimensions['A'].width = len("Semua Pengiriman Sukses") + 5

            # <-- KODE YANG HILANG DITAMBAHKAN DI SINI (Sheet 3 & 4)
            
            # Sheet Hasil RO vs Real (Logika baru - Code 2, dengan kolom Flow & Status Delivery dari Label)
            ro_centered_cols = ['Flow', 'Status Delivery', 'Open Time', 'Close Time', 'Actual Arrival', 'Actual Departure', 'Visit Time', 'Actual Visit Time', 'RO Sequence', 'Real Sequence', 'Is Same Sequence'] # Tambahkan 'Flow'
            format_excel_sheet(writer, df_ro_vs_real, 'Hasil RO vs Real', centered_cols=ro_centered_cols)

            # Sheet Update Longlat (Logika Code 1)
            if "Customer ID" in df_longlat.columns:
                longlat_centered_cols = ['Customer ID', 'Location ID', 'New Longlat', 'Beda Jarak (m)']
                format_excel_sheet(writer, df_longlat, 'Update Longlat', centered_cols=longlat_centered_cols)
            else:
                df_longlat.to_excel(writer, index=False, sheet_name='Update Longlat')

            comment_author = "System" # Definisikan author satu kali

            if "Total Delivered" in writer.sheets:
                ws_ro = writer.sheets["Total Delivered"]
                
                # Buat kamus (dictionary) untuk semua komentar di sheet ini
                comments_ro = {
                    "Total Delivered": "Total Outlet - (Pending + Batal + Terima Sebagian)",
                }
                
                # Loop baris header satu kali saja
                for cell in ws_ro[1]: 
                    if cell.value in comments_ro:
                        comment_text = comments_ro[cell.value]
                        cell.comment = Comment(comment_text, comment_author)
            # --- Komentar untuk sheet 'Hasil RO vs Real' ---
            if "Hasil RO vs Real" in writer.sheets:
                ws_ro = writer.sheets["Hasil RO vs Real"]
                
                # Buat kamus (dictionary) untuk semua komentar di sheet ini
                comments_ro = {
                    "RO Sequence": "Urutan berdasarkan hasil routing",
                    "Real Sequence": "Urutan kunjungan aktual di lapangan.",
                }
                
                # Loop baris header satu kali saja
                for cell in ws_ro[1]: 
                    if cell.value in comments_ro:
                        comment_text = comments_ro[cell.value]
                        cell.comment = Comment(comment_text, comment_author)

            # --- Komentar untuk sheet 'Update Longlat' ---
            if "Update Longlat" in writer.sheets:
                ws_longlat = writer.sheets["Update Longlat"]
                
                # Kamus untuk komentar di sheet Longlat
                comments_longlat = {
                    "Beda Jarak (m)": "Perhitungan jarak secara garis lurus antara koordinat lama dan baru."
                }
                
                # Loop baris header satu kali saja
                for cell in ws_longlat[1]: 
                    if cell.value in comments_longlat:
                        comment_text = comments_longlat[cell.value]
                        cell.comment = Comment(comment_text, comment_author)
            
            # --- ▼▼▼ KODE BARU UNTUK KOMENTAR KONDISIONAL (REQ 2) ▼▼▼ ---
            # Cek apakah ada baris yang dialihkan (fill_values punya data)
            if fill_values is not None and fill_values.notna().any():
                if "Hasil Pending SO" in writer.sheets:
                    ws_pending_comment = writer.sheets["Hasil Pending SO"]
                    target_header = "Pending"
                    comment_text = 'Warna merah menandakan harusnya pilih "Pending" bukan "Pending GR"'
                    
                    for cell in ws_pending_comment[1]: # Loop header
                        if cell.value == target_header:
                            cell.comment = Comment(comment_text, comment_author)
                            break # Hentikan pencarian

        open_file_externally(NAMA_FILE_OUTPUT)
        return True
    except Exception as e: show_error_message("Gagal Menyimpan", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=f"GAGAL MENYIMPAN FILE EXCEL: {e}\n\n{traceback.format_exc()}")); return False


def main():
    def process_wrapper(dates, app_instance):
        def safe_close():
            if app_instance and app_instance.winfo_exists():
                app_instance.progress.stop()
                app_instance.destroy()
        try:
            panggil_api_dan_simpan(dates, app_instance)
        finally:
            if app_instance and app_instance.winfo_exists():
                app_instance.after(100, safe_close)

    create_date_picker_window("Delivery Summary", process_wrapper)

if __name__ == "__main__":
    main()