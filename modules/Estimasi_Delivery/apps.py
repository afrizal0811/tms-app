import requests
import json 
import re 
import tkinter as tk
from tkinter import ttk 
import threading 
from datetime import datetime, timedelta
# Import utulity dari folder utils
from utils.function import load_config, load_secret, load_constants, show_error_message, load_master_data
from utils.messages import ERROR_MESSAGES
from utils.gui import create_date_picker_window
from utils.api_handler import handle_requests_error

# Batasan untuk tab/tombol kendaraan per halaman
VEHICLES_PER_PAGE = 10 

# =============================================================================
# FUNGSI UTILITY DATA & REGEX (Sama seperti modul sebelumnya)
# =============================================================================
def extract_outlet_name(visit_name):
    """
    Mengekstrak nama outlet dari visitName, yaitu teks sebelum pemisah ' - ' yang pertama.
    Contoh: "Primafood International, PT - C0200302..." -> "Primafood International, PT"
    """
    if not visit_name:
        return "N/A"
    
    # Mencari posisi pemisah ' - '
    separator = ' - '
    index = visit_name.find(separator)
    
    if index != -1:
        # Mengambil substring dari awal hingga sebelum pemisah
        return visit_name[:index].strip()
    
    # Jika pemisah tidak ditemukan, kembalikan seluruh string
    return visit_name.strip()

def get_hub_id():
    """Mengambil Hub ID dari master.json."""
    config = load_config()
    lokasi_code = config.get('lokasi') if config else None
    
    if not config or not lokasi_code:
        show_error_message("Error Konfigurasi", ERROR_MESSAGES.get("LOCATION_CODE_MISSING", "Kode lokasi hilang."))
        return None

    master_data = load_master_data() 
    hub_ids_map = master_data.get('hub_ids') if master_data else None
    
    hub_id = hub_ids_map.get(lokasi_code) if hub_ids_map else None

    if not hub_id:
        show_error_message("Error Hub ID", f"Hub ID untuk lokasi '{lokasi_code}' tidak ditemukan di master data.")
        return None
        
    return hub_id

def extract_customer_and_location(visit_name):
    """Mengekstrak Kode Customer (C0...) dan Kode Lokasi (LOC/MAIN/SHIPTO) dari visitName."""
    if not visit_name: return "", ""
    
    cust_code = re.search(r'(C0\d+)', visit_name)
    loc_code = re.search(r'(MAIN|SHIPTO|LOC\d+)', visit_name)
            
    return cust_code.group(1) if cust_code else "", loc_code.group(1) if loc_code else ""

# =============================================================================
# FUNGSI UTILITY GUI (Diubah untuk Treeview)
# =============================================================================

def create_vehicle_tab(notebook, vehicle_data):
    """Membuat satu tab (sheet) untuk satu kendaraan, menampilkan data dalam Treeview."""
    
    vehicle_name = vehicle_data['vehicleName']
    stop_details = vehicle_data['stopDetails'] # Data yang sudah diparsing
    num_trips = vehicle_data['numTrips']
    
    tab_frame = ttk.Frame(notebook, padding="10")
    notebook.add(tab_frame, text=vehicle_name)

    # --- Informasi Summary di Atas ---
    summary_frame = tk.Frame(tab_frame, padx=10, pady=5, relief="raised", bd=1)
    summary_frame.pack(fill='x', pady=(0, 10))
    
    ttk.Label(summary_frame, text=f"Total Stop:", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
    ttk.Label(summary_frame, text=f"{num_trips}", font=("Arial", 10, "bold")).pack(side=tk.LEFT)

    # --- Treeview (Tabel) ---
    columns = ("Urutan", "Outlet", "Jam Buka", "Jam Tutup", "Perkiraan Sampai", "Perkiraan Selesai")
    tree = ttk.Treeview(tab_frame, columns=columns, show='headings')
    tree.pack(expand=True, fill='both')

    # Konfigurasi Kolom
    tree.column("Urutan", width=60, anchor='center')
    tree.column("Outlet", width=250, anchor='w')
    tree.column("Jam Buka", width=100, anchor='center')
    tree.column("Jam Tutup", width=100, anchor='center')
    tree.column("Perkiraan Sampai", width=120, anchor='center')
    tree.column("Perkiraan Selesai", width=120, anchor='center')

    # Headings
    tree.heading("Urutan", text="Urutan")
    tree.heading("Outlet", text="Outlet")
    tree.heading("Jam Buka", text="Jam Buka")
    tree.heading("Jam Tutup", text="Jam Tutup")
    tree.heading("Perkiraan Sampai", text="Perkiraan Sampai")
    tree.heading("Perkiraan Selesai", text="Perkiraan Selesai")

    # Masukkan Data
    for stop in sorted(stop_details, key=lambda x: x['order']):
        # Format waktu menjadi HH:MM saja (dari HH:MM:SS)
        eta_short = stop['eta'].split(':')[0] + ':' + stop['eta'].split(':')[1]
        etd_short = stop['etd'].split(':')[0] + ':' + stop['etd'].split(':')[1]
        
        tree.insert('', 'end', values=(
            stop['order'], 
            stop['visitName'], 
            stop['startTime'], 
            stop['endTime'], 
            eta_short, # Perkiraan Sampai
            etd_short  # Perkiraan Selesai
        ))
        
    # Scrollbar
    scrollbar = ttk.Scrollbar(tree, orient="vertical", command=tree.yview)
    tree.configure(yscrollcommand=scrollbar.set)
    scrollbar.pack(side='right', fill='y')


def create_summary_tab(notebook, parsed_data):
    """Membuat tab Rangkuman (Hanya menunjukkan jumlah kendaraan)."""
    
    num_vehicles = len(parsed_data)
    
    summary_frame = ttk.Frame(notebook, padding="20 20 20 20")
    notebook.add(summary_frame, text="Rangkuman", sticky="nsew") 

    center_frame = tk.Frame(summary_frame)
    center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    tk.Label(center_frame, text="REKAPITULASI ESTIMASI DELIVERY", font=("Arial", 16, "bold"), pady=15).pack()

    # Info Jumlah Kendaraan
    tk.Label(center_frame, text=f"Total {num_vehicles} Kendaraan Aktif", font=("Arial", 14, "bold"), fg="blue").pack(pady=5)
    tk.Label(center_frame, text="Lihat detail estimasi waktu di setiap tab kendaraan.", font=("Arial", 10), fg="gray").pack(pady=5)

# Fungsi display_result_gui dan update_vehicle_tabs sama persis dengan modul sebelumnya
# Saya akan menyertakannya di sini untuk menjaga kelengkapan dan menghindari pemanggilan fungsi yang hilang.
# ... (lanjutan di bawah) ...

def update_vehicle_tabs(notebook, vehicle_data_list, current_page, pagination_control_frame):
    """Menghapus tab kendaraan lama dan membuat ulang tab untuk halaman saat ini."""
    
    # Hapus semua tab KECUALI tab "Rangkuman" (index 0)
    for tab_id in notebook.tabs()[1:]: notebook.forget(tab_id)

    # Hitung indeks
    start_index = current_page * VEHICLES_PER_PAGE
    end_index = start_index + VEHICLES_PER_PAGE
    vehicles_to_display = vehicle_data_list[start_index:end_index]
    
    # Buat ulang tab kendaraan
    for vehicle_data in vehicles_to_display: create_vehicle_tab(notebook, vehicle_data)

    # Update tombol panah
    total_pages = (len(vehicle_data_list) + VEHICLES_PER_PAGE - 1) // VEHICLES_PER_PAGE
    
    # Hapus tombol lama
    for widget in pagination_control_frame.winfo_children(): widget.destroy()
        
    result_window = pagination_control_frame.winfo_toplevel()
    move_page = result_window.move_page
    
    # Tombol KIRI
    left_button = ttk.Button(pagination_control_frame, text="<", command=lambda: move_page(-1))
    left_button.pack(side=tk.LEFT, padx=(0, 5))
    if current_page == 0: left_button.config(state=tk.DISABLED)

    # Label Halaman
    ttk.Label(pagination_control_frame, text=f"Halaman {current_page + 1}/{total_pages}", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)

    # Tombol KANAN
    right_button = ttk.Button(pagination_control_frame, text=">", command=lambda: move_page(1))
    right_button.pack(side=tk.LEFT, padx=(5, 0))
    if current_page == total_pages - 1: right_button.config(state=tk.DISABLED)

    notebook.select(1 if notebook.index("end") > 0 else 0)

def display_result_gui(parent_instance, parsed_data, date_str):
    """Membuat jendela Toplevel baru dengan Pagination untuk tab kendaraan."""
    if not parent_instance.winfo_exists(): return

    result_window = tk.Toplevel(parent_instance) 
    parent_instance.withdraw() 
    result_window.title(f"Hasil Estimasi Delivery - {date_str}")
    
    window_width, window_height = 1000, 700
    screen_width, screen_height = result_window.winfo_screenwidth(), result_window.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    result_window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    main_container = tk.Frame(result_window); main_container.pack(expand=True, fill='both', padx=10, pady=10)
    
    # Frame Kontrol (di atas Notebook)
    control_frame = tk.Frame(main_container, pady=5); control_frame.pack(fill='x', anchor='n') 
    
    # Frame KONTROL NAVIGASI KHUSUS (DI TENGAH)
    pagination_control_frame = tk.Frame(control_frame); pagination_control_frame.pack(padx=5) 
    
    notebook = ttk.Notebook(main_container); notebook.pack(expand=True, fill='both')

    if parsed_data:
        create_summary_tab(notebook, parsed_data)
    else:
        error_tab = ttk.Frame(notebook, padding="10"); notebook.add(error_tab, text="Kosong")
        tk.Label(error_tab, text="Tidak ada data estimasi yang ditemukan.", font=("Arial", 14)).pack(pady=20)
        return

    # Logic Pagination
    result_window.current_page = 0
    result_window.vehicle_data_list = parsed_data
    result_window.notebook = notebook
    result_window.pagination_control_frame = pagination_control_frame 
    
    def move_page(delta):
        new_page = result_window.current_page + delta
        total_pages = (len(result_window.vehicle_data_list) + VEHICLES_PER_PAGE - 1) // VEHICLES_PER_PAGE
        
        if 0 <= new_page < total_pages:
            result_window.current_page = new_page
            update_vehicle_tabs(notebook, parsed_data, result_window.current_page, pagination_control_frame) 

    result_window.move_page = move_page
    
    update_vehicle_tabs(notebook, parsed_data, result_window.current_page, pagination_control_frame)
    notebook.select(0)


# =============================================================================
# LOGIKA INTI DATA PROCESSING
# =============================================================================

# =============================================================================
# LOGIKA INTI DATA PROCESSING
# =============================================================================

def _parse_delivery_data(routing_results):
    """Mengekstrak data estimasi delivery yang diperlukan dari hasil API mentah."""
    parsed_data = []
    
    for route_item in routing_results:
        for route in route_item.get('result', {}).get('routing', []):
            vehicle_name = route.get('vehicleName', 'N/A')
            trips = route.get('trips', [])
            
            non_hub_trips = [trip for trip in trips if not trip.get('isHub')]
            num_trips = len(non_hub_trips) # Jumlah stop non-hub
            stop_details = []
            
            for trip in non_hub_trips:
                # Pastikan semua field yang dibutuhkan ada
                time_window = trip.get('timeWindow', {})
                
                if trip.get('visitName') and time_window.get('startTime') and trip.get('eta'):
                    # --- PERUBAHAN DI SINI: Panggil fungsi extract_outlet_name ---
                    outlet_name = extract_outlet_name(trip.get('visitName'))
                    # -----------------------------------------------------------
                    
                    stop_details.append({
                        "order": trip.get('order', 999),
                        "visitName": outlet_name, # Sekarang hanya nama outlet
                        "startTime": time_window.get('startTime', 'N/A'), # Jam Buka
                        "endTime": time_window.get('endTime', 'N/A'),     # Jam Tutup
                        "eta": trip.get('eta', 'N/A'),                   # Perkiraan Sampai
                        "etd": trip.get('etd', 'N/A'),                   # Perkiraan Selesai
                    })

            if stop_details:
                parsed_data.append({
                    "vehicleName": vehicle_name,
                    "numTrips": num_trips,
                    "stopDetails": stop_details
                })
    return parsed_data

def _handle_api_request_and_parse_data(app_instance, date_obj, hub_id):
    """Mengurus semua request API, filter, dan parsing data."""
    
    secret = load_secret()
    constants = load_constants()
    
    base_url = constants.get('base_url')
    token = secret.get('token')
    
    if not base_url or not token:
        show_error_message("Error API", ERROR_MESSAGES.get("API_TOKEN_MISSING", "Token API hilang."))
        return None, None

    # LOGIKA PERUBAHAN TANGGAL (SAMA PERSIS DENGAN SEBELUMNYA)
    day_of_week = date_obj.weekday()
    
    if day_of_week == 0:  # Senin
        target_date_obj = date_obj - timedelta(days=2)  # ke Sabtu
    else:
        target_date_obj = date_obj - timedelta(days=1)  # selain itu, mundur 1 hari

    mileapp_date_format = target_date_obj.strftime('%Y-%m-%d')
    date_str = date_obj.strftime('%d-%m-%Y')
    
    url = f"{base_url}/results"
    headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
    params = {'dateFrom': mileapp_date_format, 'dateTo': mileapp_date_format, 'limit': 100, 'hubId': hub_id}

    response = requests.get(url, headers=headers, params=params, timeout=30)
    response.raise_for_status()
    data = response.json()
    
    # 4. FILTERING DATA (dispatchStatus: done)
    app_instance.update_status("Memfilter data...")
    routing_results = [
        item for item in data.get('data', {}).get('data', [])
        if item.get("dispatchStatus") == "done"
    ]

    if not routing_results:
        app_instance.display_error("Data Tidak Ditemukan", ERROR_MESSAGES.get("DATA_NOT_FOUND", "Data tidak ditemukan."))
        return None, date_str
        
    app_instance.update_status("Mengekstrak dan memformat data estimasi...")
    parsed_data = _parse_delivery_data(routing_results)

    # ============================================================
    # 🔽 Tambahan baru: Urutkan kendaraan berdasarkan ETA terawal
    # ============================================================
    def get_first_eta(vehicle):
        """Ambil waktu ETA pertama (terawal) dari setiap kendaraan."""
        try:
            # Ambil eta terawal dari daftar stopDetails kendaraan
            eta_list = [datetime.strptime(stop['eta'], "%H:%M:%S") for stop in vehicle.get('stopDetails', []) if stop.get('eta')]
            return min(eta_list) if eta_list else datetime.max
        except Exception:
            return datetime.max

    # Urutkan berdasarkan ETA terawal, lalu nama kendaraan ascending
    parsed_data.sort(key=lambda v: (get_first_eta(v), v.get('vehicleName', '')))

    # ============================================================

    return parsed_data, date_str


def process_data(date_input, app_instance):
    """Fungsi utama untuk memproses data Estimasi Delivery."""
    
    date_str_input = date_input.get('dmy') if isinstance(date_input, dict) else (date_input if isinstance(date_input, str) else None)
    
    if not date_str_input:
        app_instance.display_error("Kesalahan Input", "Input tanggal tidak valid. Proses dibatalkan.")
        app_instance.after(1000, app_instance.destroy) 
        return

    try:
        date_obj = datetime.strptime(date_str_input, '%d-%m-%Y') 
        
        # LOGIKA KHUSUS: Tolak Hari Minggu (weekday(): Senin=0, Minggu=6)
        if date_obj.weekday() == 6:
            app_instance.display_error("Data Tidak Ditemukan", ERROR_MESSAGES.get("DATA_NOT_FOUND", "Data tidak ditemukan."))
            app_instance.after(1000, app_instance.destroy) 
            return

        hub_id = get_hub_id() 

        if not hub_id:
            app_instance.after(1000, app_instance.destroy) 
            return

        parsed_data, date_str = _handle_api_request_and_parse_data(app_instance, date_obj, hub_id)
        
        # 6. TAMPILKAN HASIL DI GUI BARU
        if parsed_data:
            app_instance.after(0, lambda: display_result_gui(app_instance, parsed_data, date_str))
        elif date_str:
            app_instance.display_error("Data Kosong", "Tidak ada kendaraan yang lolos filter atau tidak ada data estimasi waktu yang lengkap.")
            app_instance.after(1000, app_instance.destroy)
        
    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
        app_instance.after(1000, app_instance.destroy)
    except Exception as e:
        error_msg = ERROR_MESSAGES.get("UNKNOWN_ERROR", "Kesalahan tak terduga: {error_detail}").format(error_detail=str(e))
        app_instance.display_error("Kesalahan Tak Terduga", error_msg)
        app_instance.after(1000, app_instance.destroy) 

def main():
    """Fungsi entry point untuk modul Estimasi Delivery."""
    create_date_picker_window("Estimasi Delivery", process_data)