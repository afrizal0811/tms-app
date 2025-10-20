import requests
import json 
import re 
import tkinter as tk
from tkinter import ttk 
import threading 
from datetime import datetime, timedelta
from utils.function import load_config, load_secret, load_constants, show_error_message, load_master_data
from utils.messages import ERROR_MESSAGES
from utils.gui import create_date_picker_window
from utils.api_handler import handle_requests_error

# Batasan untuk tab/tombol kendaraan per halaman
VEHICLES_PER_PAGE = 10 

# =============================================================================
# FUNGSI UTILITY DATA
# =============================================================================

def get_hub_id():
    """Mengambil Hub ID dari master.json."""
    config = load_config()
    if not config:
        show_error_message("Error Konfigurasi", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
        return None

    lokasi_code = config.get('lokasi')
    
    if not lokasi_code:
        show_error_message("Error Konfigurasi", ERROR_MESSAGES["LOCATION_CODE_MISSING"])
        return None

    master_data = load_master_data() 
    if not master_data:
        show_error_message("Error Master Data", ERROR_MESSAGES["MASTER_FILE_ERROR"])
        return None
    
    hub_ids_map = master_data.get('hub_ids')
    
    if not hub_ids_map:
        show_error_message("Error Master Data", ERROR_MESSAGES["MASTER_DATA_MISSING"].format(details="Hub IDs map tidak ditemukan di master data."))
        return None

    hub_id = hub_ids_map.get(lokasi_code)

    if not hub_id:
        show_error_message("Error Hub ID", f"Hub ID untuk lokasi '{lokasi_code}' tidak ditemukan di master data.")
        return None
        
    return hub_id

def extract_customer_and_location(visit_name):
    """
    Mengekstrak Kode Customer (C0...) dan Kode Lokasi (LOC/MAIN/SHIPTO) dari visitName.
    """
    cust_code = ""
    loc_code = ""
    
    if not visit_name:
        return cust_code, loc_code
    
    cust_match = re.search(r'(C0\d+)', visit_name)
    if cust_match:
        cust_code = cust_match.group(1)
        
    loc_match = re.search(r'(MAIN|SHIPTO|LOC\d+)', visit_name)
    if loc_match:
        loc_code = loc_match.group(1)
            
    return cust_code, loc_code

# =============================================================================
# FUNGSI GUI OUTPUT BARU
# =============================================================================

def copy_to_clipboard(value):
    """Menyalin nilai ke clipboard TANPA menampilkan notifikasi pop-up."""
    try:
        root = tk.Tk()
        root.withdraw() 
        root.clipboard_clear()
        root.clipboard_append(value)
        root.update()
        root.destroy()
    except tk.TclError:
        pass 

def create_task_box(parent_frame, detail):
    """
    Membuat box Task (Stop Detail) sesuai format gambar. 
    """
    
    so_count = len(detail['soNumbers'])
    
    # Border
    box_frame = tk.Frame(parent_frame, bd=1, relief="solid", padx=5, pady=5, bg="#F9F9F9")

    # --- Baris 1: Customer ID ---
    cust_frame = tk.Frame(box_frame, bg="#F9F9F9")
    cust_frame.pack(fill='x', pady=2)
    tk.Label(cust_frame, text="Customer ID:", font=("Arial", 9, "bold"), bg="#F9F9F9").pack(side=tk.LEFT)
    tk.Label(cust_frame, text=detail['customerID'], font=("Consolas", 9), bg="#F9F9F9").pack(side=tk.LEFT, padx=(5, 0))
    # Tombol Copy Customer ID
    ttk.Button(cust_frame, text="Copy", command=lambda v=detail['customerID']: copy_to_clipboard(v), width=5).pack(side=tk.RIGHT)

    # --- Baris 2: Location ID ---
    loc_frame = tk.Frame(box_frame, bg="#F9F9F9")
    loc_frame.pack(fill='x', pady=2)
    tk.Label(loc_frame, text="Location ID:", font=("Arial", 9, "bold"), bg="#F9F9F9").pack(side=tk.LEFT)
    tk.Label(loc_frame, text=detail['locationCode'], font=("Consolas", 9), bg="#F9F9F9").pack(side=tk.LEFT, padx=(5, 0))
    # Tombol Copy Location ID
    ttk.Button(loc_frame, text="Copy", command=lambda v=detail['locationCode']: copy_to_clipboard(v), width=5).pack(side=tk.RIGHT)

    # --- Baris 3: Nomor SO ---
    tk.Label(box_frame, text="Nomor SO:", font=("Arial", 9, "bold"), bg="#F9F9F9").pack(anchor='w', pady=(5, 2))
    
    so_text_content = "\n".join(sorted(detail['soNumbers']))
    
    display_height = max(so_count, 2) 
    
    # Widget Text fleksibel mengikuti lebar kolom Grid
    so_text_widget = tk.Text(box_frame, wrap=tk.NONE, height=display_height, 
                             font=("Consolas", 9), padx=3, pady=3, bd=1, relief="sunken")
    so_text_widget.insert(tk.END, so_text_content)
    so_text_widget.config(state=tk.DISABLED)
    so_text_widget.pack(fill='x', expand=True)

    return box_frame # Mengembalikan frame agar bisa di-grid/pack oleh pemanggil

def create_vehicle_tab(notebook, vehicle_data):
    """Membuat satu tab (sheet) untuk satu kendaraan."""
    
    vehicle_name = vehicle_data['vehicleName']
    num_trips = vehicle_data['numTrips']
    details_per_stop = vehicle_data['detailsPerStop']
    total_so = sum(len(d['soNumbers']) for d in details_per_stop)

    tab_frame = ttk.Frame(notebook, padding="10 10 10 10")
    notebook.add(tab_frame, text=vehicle_name)

    # --- Informasi Summary di Atas ---
    summary_frame = tk.Frame(tab_frame, padx=10, pady=5, relief="raised", bd=1)
    summary_frame.pack(fill='x', pady=(0, 10))
    
    # Label: Total Task
    ttk.Label(summary_frame, text=f"Total Task:", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
    ttk.Label(summary_frame, text=f"{num_trips}", font=("Arial", 10, "bold")).pack(side=tk.LEFT)

    ttk.Label(summary_frame, text="|", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
    
    # Label: Total SO
    ttk.Label(summary_frame, text=f"Total SO:", font=("Arial", 10)).pack(side=tk.LEFT, padx=10)
    ttk.Label(summary_frame, text=f"{total_so}", font=("Arial", 10, "bold")).pack(side=tk.LEFT)

    # --- Area Konten Utama (Scrollable) ---
    main_canvas = tk.Canvas(tab_frame)
    main_canvas.pack(side="left", fill="both", expand=True)

    v_scrollbar = ttk.Scrollbar(tab_frame, orient="vertical", command=main_canvas.yview)
    v_scrollbar.pack(side="right", fill="y")
    
    content_frame = tk.Frame(main_canvas)
    # Set width awal 1 agar on_canvas_resize dapat mengatur lebar yang benar
    main_canvas.create_window((0, 0), window=content_frame, anchor="nw", width=1) 

    def on_canvas_resize(event):
        # Resize window canvas (agar content_frame mengikuti lebar canvas)
        main_canvas.itemconfig(main_canvas.find_all()[-1], width=event.width)
        # Atur scrollregion setelah resize
        main_canvas.configure(scrollregion = main_canvas.bbox("all"))
        
    main_canvas.bind('<Configure>', on_canvas_resize)
    main_canvas.configure(yscrollcommand=v_scrollbar.set)
    
    # --- FUNGSI BINDING MOUSE WHEEL ---
    def bind_mouse_wheel(widget):
        # Binding untuk Windows/X11 (roda mouse)
        widget.bind('<MouseWheel>', lambda e: main_canvas.yview_scroll(int(-1*(e.delta/120)), "units"))
        # Binding untuk Linux/Mac (Button-4, Button-5)
        widget.bind('<Button-4>', lambda e: main_canvas.yview_scroll(-1, "units"))
        widget.bind('<Button-5>', lambda e: main_canvas.yview_scroll(1, "units"))

    bind_mouse_wheel(main_canvas)
    bind_mouse_wheel(content_frame)

    # --- Penempatan Task Box (Grid Dinamis) ---
    num_stops = len(details_per_stop)
    
    all_widgets_to_bind = []
    
    for i in range(0, num_stops, 3):
        row_frame = tk.Frame(content_frame)
        row_frame.pack(fill='x', expand=True) 
        all_widgets_to_bind.append(row_frame)
        
        stops_in_row = details_per_stop[i:i+3]
        num_in_row = len(stops_in_row)
        
        is_last_row = (i + num_in_row == num_stops)
        
        # Inisialisasi column weights
        for col in range(3):
            row_frame.grid_columnconfigure(col, weight=0) 

        # 1. Konfigurasi Grid (Dynamic Column Weight)
        if is_last_row:
            if num_in_row == 1:
                # Modulo 3 = 1 -> Full width
                row_frame.grid_columnconfigure(0, weight=1)
            elif num_in_row == 2:
                # Modulo 3 = 2 -> Divided by 2
                row_frame.grid_columnconfigure(0, weight=1)
                row_frame.grid_columnconfigure(1, weight=1)
            else:
                # Full row (3 boxes)
                row_frame.grid_columnconfigure(0, weight=1)
                row_frame.grid_columnconfigure(1, weight=1)
                row_frame.grid_columnconfigure(2, weight=1)
        else:
            # Full row (3 boxes)
            row_frame.grid_columnconfigure(0, weight=1)
            row_frame.grid_columnconfigure(1, weight=1)
            row_frame.grid_columnconfigure(2, weight=1)


        # 2. Penempatan Kotak Tugas
        for col_index in range(num_in_row):
            detail = stops_in_row[col_index]
            
            cell_frame = tk.Frame(row_frame, padx=5, pady=5) 
            all_widgets_to_bind.append(cell_frame)

            task_box = create_task_box(cell_frame, detail)
            all_widgets_to_bind.append(task_box)
            
            # Pack kotak tugas ke dalam cell_frame
            task_box.pack(fill=tk.BOTH, expand=True)
            
            # Tentukan columnspan khusus untuk kasus 1 box di baris terakhir
            current_column_span = 3 if (is_last_row and num_in_row == 1 and col_index == 0) else 1

            # Tempatkan cell_frame di grid
            cell_frame.grid(row=0, column=col_index, sticky='nsew', columnspan=current_column_span)
            
            # Ambil semua widget anak di dalam task_box untuk di-bind (agar mouse wheel seamless)
            for inner_widget in task_box.winfo_children():
                all_widgets_to_bind.append(inner_widget)

    # 3. Rekursif Binding Mouse Wheel ke SEMUA widget anak
    for widget in all_widgets_to_bind:
        bind_mouse_wheel(widget)


def create_summary_tab(notebook, parsed_data):
    """
    Membuat tab Rangkuman di posisi awal yang HANYA menampilkan Akumulasi Total.
    """
    
    # Hitung Akumulasi Total
    total_task_all = sum(data['numTrips'] for data in parsed_data)
    total_so_all = sum(len(detail['soNumbers']) for data in parsed_data for detail in data['detailsPerStop'])
    
    summary_frame = ttk.Frame(notebook, padding="20 20 20 20")
    notebook.add(summary_frame, text="Rangkuman", sticky="nsew") 

    # --- Area Konten Utama (Diposisikan di tengah) ---
    center_frame = tk.Frame(summary_frame)
    center_frame.place(relx=0.5, rely=0.5, anchor=tk.CENTER)

    # --- Header ---
    tk.Label(center_frame, text="REKAPITULASI ROUTING TRANSACTION", font=("Arial", 16, "bold"), pady=15).pack()

    # --- Bagian Akumulasi Total (Paling Atas) ---
    total_frame = tk.Frame(center_frame, bd=2, relief="solid", padx=20, pady=10, bg="#E0F7FA")
    total_frame.pack(pady=15)
    
    # Total Task Akumulasi
    tk.Label(total_frame, text="TOTAL TASK AKUMULASI", font=("Arial", 14), bg="#E0F7FA").grid(row=0, column=0, padx=20, pady=5, sticky="w")
    tk.Label(total_frame, text=f": {total_task_all}", font=("Consolas", 20, "bold"), fg="blue", bg="#E0F7FA").grid(row=0, column=1, pady=5, sticky="w")
    
    # Total SO Akumulasi
    tk.Label(total_frame, text="TOTAL SO AKUMULASI", font=("Arial", 14), bg="#E0F7FA").grid(row=1, column=0, padx=20, pady=5, sticky="w")
    tk.Label(total_frame, text=f": {total_so_all}", font=("Consolas", 20, "bold"), fg="red", bg="#E0F7FA").grid(row=1, column=1, pady=5, sticky="w")
    
    # Info Jumlah Kendaraan
    tk.Label(center_frame, text=f"Total {len(parsed_data)} Kendaraan Aktif", font=("Arial", 10), fg="gray").pack(pady=5)


# =============================================================================
# PAGINATION LOGIC & GUI
# =============================================================================

def update_vehicle_tabs(notebook, vehicle_data_list, current_page, pagination_control_frame):
    """Menghapus tab kendaraan lama dan membuat ulang tab untuk halaman saat ini."""
    
    # Hapus semua tab KECUALI tab "Rangkuman" (index 0)
    tabs = notebook.tabs()
    for tab_id in tabs[1:]:
        notebook.forget(tab_id)

    # Hitung indeks awal dan akhir untuk data kendaraan di halaman saat ini
    start_index = current_page * VEHICLES_PER_PAGE
    end_index = start_index + VEHICLES_PER_PAGE
    
    vehicles_to_display = vehicle_data_list[start_index:end_index]
    
    # Buat ulang tab kendaraan untuk halaman saat ini
    for vehicle_data in vehicles_to_display:
        create_vehicle_tab(notebook, vehicle_data)

    # Update tombol panah (pagination_control_frame)
    total_pages = (len(vehicle_data_list) + VEHICLES_PER_PAGE - 1) // VEHICLES_PER_PAGE
    
    # Hapus tombol lama
    for widget in pagination_control_frame.winfo_children():
        widget.destroy()
        
    # Ambil fungsi navigasi dari jendela utama (Top-level window)
    result_window = pagination_control_frame.winfo_toplevel()
    move_page = result_window.move_page
    
    # Tombol KIRI
    left_button = ttk.Button(pagination_control_frame, text="<", command=lambda: move_page(-1))
    left_button.pack(side=tk.LEFT, padx=(0, 5))
    if current_page == 0:
        left_button.config(state=tk.DISABLED)

    # Label Halaman
    ttk.Label(pagination_control_frame, text=f"Halaman {current_page + 1}/{total_pages}", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)

    # Tombol KANAN
    right_button = ttk.Button(pagination_control_frame, text=">", command=lambda: move_page(1))
    right_button.pack(side=tk.LEFT, padx=(5, 0))
    if current_page == total_pages - 1:
        right_button.config(state=tk.DISABLED)

    # Pilih tab kendaraan pertama di halaman baru (jika ada) atau tab Rangkuman
    if notebook.index("end") > 0:
        notebook.select(1)
    else:
        notebook.select(0)


def display_result_gui(parent_instance, parsed_data, date_str):
    """
    Membuat jendela Toplevel baru dengan Pagination untuk tab kendaraan.
    """
    if not parent_instance.winfo_exists(): return

    # 1. Buat Jendela Baru (Toplevel)
    result_window = tk.Toplevel(parent_instance) 
    
    # 2. SEMBUNYIKAN GUI Date Selector (parent_instance: tk.Tk)
    parent_instance.withdraw() 

    # 3. Setup result_window
    result_window.title(f"Hasil Routing - {date_str}")
    
    # 4. Atur ukuran dan posisi jendela (Non-Maximized - Ukuran Awal)
    window_width = 1000
    window_height = 700
    screen_width = result_window.winfo_screenwidth()
    screen_height = result_window.winfo_screenheight()
    center_x = int(screen_width/2 - window_width / 2)
    center_y = int(screen_height/2 - window_height / 2)
    result_window.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
    
    # --- Frame Utama untuk Konten dan Pagination ---
    main_container = tk.Frame(result_window)
    main_container.pack(expand=True, fill='both', padx=10, pady=10)
    
    # --- Frame untuk Kontrol Pagination (di atas Notebook) ---
    control_frame = tk.Frame(main_container, pady=5)
    control_frame.pack(fill='x', anchor='n') 
    
    # Frame KONTROL NAVIGASI KHUSUS (Button < dan >) - DIPOSISIKAN DI TENGAH (MENGGUNAKAN PACK TANPA SIDE)
    pagination_control_frame = tk.Frame(control_frame)
    pagination_control_frame.pack(padx=5) 
    
    # 4. Buat Notebook (Tabbed Interface)
    notebook = ttk.Notebook(main_container) 
    notebook.pack(expand=True, fill='both')

    # 5. Isi Notebook (Tab Rangkuman)
    if parsed_data:
        create_summary_tab(notebook, parsed_data)
        
    else:
        error_tab = ttk.Frame(notebook, padding="10")
        notebook.add(error_tab, text="Kosong")
        tk.Label(error_tab, text="Tidak ada data routing yang ditemukan setelah filter.", font=("Arial", 14)).pack(pady=20)
        return

    # --- Logic Pagination ---
    # Inisialisasi status pagination
    result_window.current_page = 0
    result_window.vehicle_data_list = parsed_data
    result_window.notebook = notebook
    result_window.pagination_control_frame = pagination_control_frame 
    
    # Fungsi navigasi (disimpan sebagai atribut window)
    def move_page(delta):
        new_page = result_window.current_page + delta
        total_pages = (len(result_window.vehicle_data_list) + VEHICLES_PER_PAGE - 1) // VEHICLES_PER_PAGE
        
        if 0 <= new_page < total_pages:
            result_window.current_page = new_page
            update_vehicle_tabs(result_window.notebook, result_window.vehicle_data_list, result_window.current_page, result_window.pagination_control_frame) 

    result_window.move_page = move_page
    
    # Tampilkan halaman pertama
    update_vehicle_tabs(notebook, parsed_data, result_window.current_page, pagination_control_frame)
    # Setelah tab dibuat, kembali ke tab Rangkuman
    notebook.select(0)


# =============================================================================
# FUNGSI UTAMA PROSES DATA
# =============================================================================

def process_data(date_input, app_instance):
    """Fungsi utama untuk memproses data Routing Transaction dari API.
    Urutkan kendaraan berdasarkan ETD dari HUB pertama (berdasarkan 'order')."""
    date_str = date_input.get('dmy') if isinstance(date_input, dict) else (date_input if isinstance(date_input, str) else None)

    if not date_str:
        app_instance.display_error("Kesalahan Input", "Input tanggal tidak valid. Proses dibatalkan.")
        app_instance.after(1000, app_instance.destroy)
        return

    try:
        secret = load_secret()
        constants = load_constants()
        hub_id = get_hub_id()

        if not secret or not constants or not hub_id:
            app_instance.after(1000, app_instance.destroy)
            return

        base_url = constants.get('base_url')
        token = secret.get('token')

        if not base_url or not token:
            show_error_message("Error API", ERROR_MESSAGES["API_TOKEN_MISSING"])
            app_instance.after(1000, app_instance.destroy)
            return

        date_obj = datetime.strptime(date_str, '%d-%m-%Y')
        day_of_week = date_obj.weekday()

        if day_of_week == 6:
            app_instance.display_error("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
            app_instance.after(1000, app_instance.destroy)
            return
        elif day_of_week == 0:
            target_date_obj = date_obj - timedelta(days=2)
        else:
            target_date_obj = date_obj - timedelta(days=1)

        mileapp_date_format = target_date_obj.strftime('%Y-%m-%d')
        url = f"{base_url}/results"
        headers = {"Authorization": f"Bearer {token}", "Content-Type": "application/json"}
        params = {"dateFrom": mileapp_date_format, "dateTo": mileapp_date_format, "limit": 100, "hubId": hub_id}

        response = requests.get(url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        data = response.json()

        app_instance.update_status("Memfilter data...")
        routing_results = [
            item for item in data.get('data', {}).get('data', [])
            if item.get("dispatchStatus") == "done"
        ]

        if not routing_results:
            app_instance.display_error("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
            app_instance.after(1000, app_instance.destroy)
            return

        app_instance.update_status("Mengekstrak dan memformat data trips dan stops...")
        parsed_data = []

        for route_item in routing_results:
            routing_list = route_item.get('result', {}).get('routing', [])
            for route in routing_list:
                vehicle_name = route.get('vehicleName', 'N/A')
                trips = route.get('trips', []) or []
                trips_sorted = sorted(trips, key=lambda t: t.get('order', 9999))
                
                # ============================================================
                # ▼▼▼ LOGIKA BARU UNTUK MENCARI ETD HUB PERTAMA ▼▼▼
                # ============================================================
                hub_etd = datetime.max # Nilai default jika ETD HUB tidak ditemukan
                
                first_hub_trip = next((trip for trip in trips_sorted if trip.get('isHub')), None)
                
                if first_hub_trip:
                    etd_raw = first_hub_trip.get('etd')
                    if etd_raw:
                        try:
                            if re.fullmatch(r"\d{2}:\d{2}(:\d{2})?", etd_raw):
                                hub_etd = datetime.combine(datetime.today().date(), datetime.strptime(etd_raw, "%H:%M:%S").time())
                            else:
                                hub_etd = datetime.fromisoformat(etd_raw.replace('Z', '+00:00'))
                        except Exception:
                            pass # Biarkan hub_etd tetap datetime.max jika format salah
                # ============================================================

                non_hub_trips = [trip for trip in trips_sorted if not trip.get('isHub')]
                num_trips = len(non_hub_trips)
                details_per_stop = []

                for trip in non_hub_trips:
                    visit_name = trip.get('visitName')
                    if visit_name:
                        cust_code, loc_code = extract_customer_and_location(visit_name)
                        if cust_code and loc_code:
                            so_numbers = []
                            so_list_start_match = re.search(r'(SO\d+-\d+)', visit_name)
                            if so_list_start_match:
                                so_start_index = so_list_start_match.start()
                                so_list_string = visit_name[so_start_index:].strip()
                                so_numbers.extend([so.strip() for so in so_list_string.split(',')])
                            if so_numbers:
                                details_per_stop.append({
                                    "customerID": cust_code,
                                    "locationCode": loc_code,
                                    "soNumbers": so_numbers
                                })

                if details_per_stop:
                    parsed_data.append({
                        "vehicleName": vehicle_name,
                        "numTrips": num_trips,
                        "detailsPerStop": details_per_stop,
                        "hub_etd": hub_etd  # Tambahkan ETD HUB ke data
                    })

        # Urutkan berdasarkan ETD HUB, lalu nama kendaraan
        parsed_data.sort(key=lambda x: (x.get('hub_etd', datetime.max), x.get('vehicleName', '')))

        if parsed_data:
            app_instance.after(0, lambda: display_result_gui(app_instance, parsed_data, date_str))
        else:
            app_instance.display_error(
                "Data Kosong",
                "Tidak ada kendaraan yang lolos filter (status 'done' dan memiliki SO, Customer ID, dan Lokasi)."
            )
            app_instance.after(1000, app_instance.destroy)

    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
        app_instance.after(1000, app_instance.destroy)
    except Exception as e:
        error_msg = ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=str(e))
        app_instance.display_error("Kesalahan Tak Terduga", error_msg)
        app_instance.after(1000, app_instance.destroy)

def main():
    """Fungsi entry point untuk modul Routing Transaction."""
    create_date_picker_window("Routing Transaction", process_data)