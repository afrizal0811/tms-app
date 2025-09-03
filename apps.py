# @GUI/apps.py (KODE YANG DIMODIFIKASI)

from tkinter import ttk
from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK
import os
import requests
import sys
import threading
import tkinter as tk
import webbrowser
import time
from utils.function import (
    CONFIG_PATH,
    MASTER_JSON_PATH,
    TYPE_MAP_PATH,
    ensure_config_exists,
    load_config,
    load_constants,
    resource_path,
    save_json_data,
    show_error_message,
    show_info_message,
    show_ask_message
)
from utils.messages import (
    ERROR_MESSAGES,
    INFO_MESSAGES,
    ASK_MESSAGES,
)

# --- Impor Modul Aplikasi ---
from modules.Routing_Summary.apps import main as routing_summary_main
from modules.Delivery_Summary.apps import main as delivery_summary_main
from modules.Auto_Delivery_Summary.apps import main as auto_delivery_summary_main
from modules.Start_Finish_Time.apps import main as start_finish_time_main
from modules.Sync_Data.apps import main as sync_data_main
from modules.Check_User.apps import main as check_user_main
from modules.Auto_Routing_Summary.apps import main as auto_routing_summary_main
from modules.Vehicles_Data.apps import main as vehicles_data_main # Impor modul baru

ensure_config_exists()
# ==============================================================================
# FUNGSI BANTUAN LOKAL DAN KONFIGURASI AWAL
# ==============================================================================

# Muat konstanta di awal menggunakan shared_utils
CONSTANTS = load_constants()
if CONSTANTS is None:
    # Pesan error sudah ditampilkan di dalam shared_utils, kita hanya perlu keluar.
    sys.exit(1)

LOKASI_MAPPING = CONSTANTS.get('lokasi_mapping', {})
LOKASI_DISPLAY = CONSTANTS.get('lokasi_display', {})
KODE_KE_LOKASI = {v: k for k, v in LOKASI_MAPPING.items()}
USER_GUIDE_PLANNER = CONSTANTS.get('guide_planner', '')
USER_GUIDE_DRIVER = CONSTANTS.get('guide_driver', '')

def reset_config_and_exit():
    """Menghapus config.json, master.json, dan type_map agar setup wajib diulang, lalu keluar aplikasi."""
    try:
        if os.path.exists(CONFIG_PATH):
            os.remove(CONFIG_PATH)
        if os.path.exists(MASTER_JSON_PATH):
            os.remove(MASTER_JSON_PATH)
        if os.path.exists(TYPE_MAP_PATH):
            os.remove(TYPE_MAP_PATH)
        show_error_message("Setup Tidak Lengkap", ERROR_MESSAGES["SETUP_CANCELED"])
        on_closing()
    except Exception:
        pass


def update_title(root_window):
    """Membaca konfigurasi dan memperbarui judul window utama."""
    config = load_config()
    title = "TMS Data Processing"
    
    if config:
        kode = config.get("lokasi")
        if kode and kode in KODE_KE_LOKASI:
            nama_lokasi = KODE_KE_LOKASI[kode]
            title += f" - {nama_lokasi}"
            
    root_window.title(title)

def toggle_main_controls(is_enabled: bool):
    """
    Mengaktifkan atau menonaktifkan semua tombol utama dan item menu terkait.
    :param is_enabled: True untuk enable, False untuk disable.
    """
    state = 'normal' if is_enabled else 'disabled'

    for button in main_buttons:
        button.config(state=state)

    bantuan_menu.entryconfig("Panduan Pengguna - Planner", state=state)
    bantuan_menu.entryconfig("Panduan Pengguna - Driver", state=state)
    bantuan_menu.entryconfig("Tentang", state=state)

    konfigurasi_menu.entryconfig("Sinkronisasi Data", state=state)
    
    laporan_menu.entryconfig("Routing Summary", state=state)
    laporan_menu.entryconfig("Delivery Summary", state=state)
    laporan_menu.entryconfig("Data Kendaraan", state=state)

def pilih_lokasi(parent_window, initial_setup=False):
    reverse_dict = {v: k for k, v in LOKASI_DISPLAY.items()}
    selected_display_name = list(LOKASI_DISPLAY.keys())[0]

    config_data = load_config() or {}
    kode_lokasi = config_data.get("lokasi", "")
    if kode_lokasi in reverse_dict:
        selected_display_name = reverse_dict[kode_lokasi]

    dialog = tk.Toplevel(parent_window)
    dialog.title("Pilih Lokasi Cabang")
    lebar, tinggi = 350, 180
    x, y = (dialog.winfo_screenwidth() - lebar) // 2, (dialog.winfo_screenheight() - tinggi) // 2
    dialog.geometry(f"{lebar}x{tinggi}+{x}+{y}")
    dialog.resizable(False, False)

    tk.Label(dialog, text="Pilih Lokasi Cabang:", font=("Arial", 14)).pack(pady=10)
    selected_var = tk.StringVar(value=selected_display_name)
    combo = ttk.Combobox(dialog, values=list(LOKASI_DISPLAY.keys()), textvariable=selected_var, font=("Arial", 12), state="readonly")
    combo.pack(pady=10)
    combo.set(selected_display_name)
    
    toggle_main_controls(False)

    def on_select():
        selected = combo.get()
        if selected in LOKASI_DISPLAY:
            kode = LOKASI_DISPLAY[selected]
            config_data['lokasi'] = kode
            save_json_data(config_data, CONFIG_PATH)
            dialog.destroy()
            sync_data_main(reset_config_and_exit)

    # --- Tambahkan handler close ---
    def on_cancel():
        """Handler saat user menutup jendela konfigurasi."""
        if initial_setup:
            # Tampilkan konfirmasi
            if show_ask_message("Konfirmasi", ASK_MESSAGES["CONFIRM_CANCEL_SETUP"]):
                reset_config_and_exit()  # jika YA
                return
            else:
                return  # jika TIDAK, tetap di window ini (tidak destroy)
        dialog.destroy()  # jika bukan setup awal, cukup tutup dialog

    dialog.protocol("WM_DELETE_WINDOW", on_cancel)
    tk.Button(dialog, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)
    dialog.transient(parent_window)
    dialog.grab_set()
    parent_window.wait_window(dialog)
    toggle_main_controls(True)

def pilih_pengguna_awal(parent_window):
    check_user_main(parent_window)  # buka GUI pilih pengguna
    config_after = load_config()
    if not config_after or not config_after.get("user_checked"):
        # User belum menyelesaikan setup pengguna
        if show_ask_message("Konfirmasi", ASK_MESSAGES["CONFIRM_CANCEL_SETUP"]):
            reset_config_and_exit()  # kalau YA → reset dan keluar
        else:
            # kalau TIDAK → buka lagi proses pilih pengguna
            pilih_pengguna_awal(parent_window)

def on_closing():
    try:
        root.destroy()
    finally:
        import os
        os._exit(0)

def show_update_dialog(latest_version, current_version, download_link, show_checkbox=True):
    """Menampilkan dialog update dengan opsi 'Jangan tampilkan lagi'."""
    dialog = tk.Toplevel(root)
    dialog.title("Update Tersedia")
    dialog.resizable(False, False)
    dialog.transient(root)
    dialog.grab_set()

    # --- posisi dialog di tengah layar ---
    dialog.update_idletasks()
    w, h = 300, 150
    x = (dialog.winfo_screenwidth() // 2) - (w // 2)
    y = (dialog.winfo_screenheight() // 2) - (h // 2)
    dialog.geometry(f"{w}x{h}+{x}+{y}")

    # --- isi pesan ---
    message = f"Versi terbaru {latest_version} tersedia.\nSaat ini Anda menggunakan versi {current_version}."
    tk.Label(
        dialog,
        text=message,
        font=("Arial", 10),
        wraplength=320,
        justify="left"
    ).pack(pady=10, padx=10)

    # --- frame tombol (tengah) ---
    btn_frame = tk.Frame(dialog)
    if show_checkbox:
        btn_frame.pack(pady=5)
    else:
        btn_frame.pack(pady=25)
    btn_update = tk.Button(btn_frame, text="Update", width=10)
    btn_update.grid(row=0, column=0, padx=5)
    btn_skip = tk.Button(btn_frame, text="Nanti Saja", width=10)
    btn_skip.grid(row=0, column=1, padx=5)

    # --- checkbox di pojok kiri bawah (opsional) ---
    if show_checkbox:
        skip_var = tk.BooleanVar(value=False)
        chk = tk.Checkbutton(
            dialog,
            text="Jangan tampilkan lagi untuk versi ini",
            variable=skip_var,
            font=("Arial", 8),
            anchor="w",
            justify="left"
        )
        chk.pack(side="left", anchor="sw", padx=8, pady=5)
    else:
        skip_var = None  # dummy

    # --- handler tombol ---
    def on_update():
        webbrowser.open(download_link)
        dialog.destroy()  # tidak simpan skip_update_version

    def on_skip():
        if show_checkbox and skip_var.get():
            config = load_config() or {}
            config["skip_update_version"] = latest_version
            save_json_data(config, CONFIG_PATH)
        dialog.destroy()

    btn_update.config(command=on_update)
    btn_skip.config(command=on_skip)

    dialog.wait_window(dialog)

def check_update(ignore_skip=False, show_checkbox=True):
    """Memeriksa versi baru dan tampilkan dialog update jika perlu."""
    try:
        response = requests.get(REMOTE_VERSION_URL, timeout=5)
        response.raise_for_status()
        latest_version = response.text.strip()

        config = load_config() or {}
        skipped_version = config.get("skip_update_version")

        if latest_version > CURRENT_VERSION and (ignore_skip or skipped_version != latest_version):
            show_update_dialog(latest_version, CURRENT_VERSION, DOWNLOAD_LINK, show_checkbox=show_checkbox)
        else:
            if ignore_skip:  # dipanggil dari menu Bantuan
                show_info_message("Update", "Anda sudah menggunakan versi terbaru.")

    except requests.exceptions.RequestException:
        if ignore_skip:  # kalau dicek manual tapi gagal koneksi
            show_error_message("Update", INFO_MESSAGES["ALREADY_UPDATED"])


def periksa_konfigurasi_awal(parent_window):
    config = load_config()
    if not config or not config.get("lokasi"):
        show_info_message("Setup Lokasi", INFO_MESSAGES["LOCATION_SETUP"])
        pilih_lokasi(parent_window, initial_setup=True)
        update_title(parent_window)
        config = load_config()


    if not config or not config.get("user_checked"):
        show_info_message("Setup Akun", INFO_MESSAGES["USER_SETUP"])
        pilih_pengguna_awal(parent_window)

def atur_visibilitas_menu(menu_bar):
    """Mengatur visibilitas item menu berdasarkan role pengguna."""
    config = load_config()
    constants = load_constants()
    user_info = config.get("user_checked", {})
    user_role_id = user_info.get("role_id")
    role_ids = constants.get("role_ids", {})
    restricted_role_id_list = [
        role_ids.get("planner"),
        role_ids.get("checker")
    ]

    try:
        if user_role_id and user_role_id in restricted_role_id_list:
            konfigurasi_menu.delete("Ganti Lokasi Cabang")
    except tk.TclError:
        pass

def run_sync_in_background(root_window):
    """Menjalankan proses sinkronisasi hub dan driver di background thread."""
    loading_window = tk.Toplevel(root_window)
    loading_window.title("Sinkronisasi")
    loading_window.geometry("300x120")
    x, y = root_window.winfo_x() + (root_window.winfo_width() // 2) - 150, root_window.winfo_y() + (root_window.winfo_height() // 2) - 60
    loading_window.geometry(f"+{x}+{y}")
    loading_window.transient(root_window)
    loading_window.grab_set()

    # --- Tambahkan konfigurasi style di sini ---
    style = ttk.Style(loading_window)
    style.theme_use('clam')
    style.configure("TButton", font=("Helvetica", 12), padding=5)
    style.configure("TLabel", background='SystemButtonFace', font=("Helvetica", 16, "bold"))
    style.configure("TProgressbar", thickness=20)
    # --- Akhir konfigurasi style ---

    status_label = ttk.Label(loading_window, text="Sinkronisasi master data...", font=("Arial", 12))
    status_label.pack(pady=(10, 0))

    timer_label = ttk.Label(loading_window, text="00:00:00", font=("Arial", 10), foreground="gray")
    timer_label.pack(pady=(0, 5))

    # --- Terapkan style yang baru di sini ---
    progress = ttk.Progressbar(loading_window, mode='indeterminate', style="Custom.Horizontal.TProgressbar")
    progress.pack(pady=10, padx=20, fill=tk.X)
    progress.start()
    
    toggle_main_controls(False)

    start_time = time.time()
    timer_running = True

    def update_timer():
        if timer_running and loading_window.winfo_exists():
            elapsed_time = int(time.time() - start_time)
            hours = elapsed_time // 3600
            minutes = (elapsed_time % 3600) // 60
            seconds = elapsed_time % 60
            timer_label.config(text=f"{hours:02}:{minutes:02}:{seconds:02}")
            loading_window.after(1000, update_timer)

    def on_sync_complete():
        nonlocal timer_running
        timer_running = False
        
        if loading_window.winfo_exists():
            progress.stop()
            loading_window.destroy()

        toggle_main_controls(True)

    def thread_target():
        try:
            sync_data_main(reset_config_and_exit) # PERBAIKAN: Berikan argumen yang sesuai
        finally:
            root_window.after(0, on_sync_complete)
            
    update_timer()
    threading.Thread(target=thread_target, daemon=True).start()

# ==============================================================================
# ALUR UTAMA APLIKASI GUI
# ==============================================================================

# --- Setup Window Utama ---
root = tk.Tk()
root.withdraw() 

def ganti_lokasi():
    pilih_lokasi(root, initial_setup=False)
    update_title(root)

# --- Setup Menu Bar ---
menu_bar = tk.Menu(root)

laporan_menu = tk.Menu(menu_bar, tearoff=0)
laporan_menu.add_command(label="Routing Summary", command=routing_summary_main)
laporan_menu.add_command(label="Delivery Summary", command=delivery_summary_main)
laporan_menu.add_separator()
laporan_menu.add_command(label="Data Kendaraan", command=vehicles_data_main)
menu_bar.add_cascade(label="Laporan", menu=laporan_menu)

konfigurasi_menu = tk.Menu(menu_bar, tearoff=0)
konfigurasi_menu.add_command(label="Ganti Lokasi Cabang", command=ganti_lokasi)
konfigurasi_menu.add_command(label="Sinkronisasi Data", command=lambda: run_sync_in_background(root))
menu_bar.add_cascade(label="Konfigurasi", menu=konfigurasi_menu)

def show_about():
    show_info_message(
        "Tentang Aplikasi",
        INFO_MESSAGES["APP_VERSION"].format(version=CURRENT_VERSION)
        + "\n\n"
        + INFO_MESSAGES["APP_BUILD_BY"]
    )

def show_user_guide(link):
    """Fungsi untuk membuka panduan pengguna di browser web."""
    webbrowser.open(link) 

bantuan_menu = tk.Menu(menu_bar, tearoff=0)
bantuan_menu.add_command(label="Panduan Pengguna - Planner", command=lambda:show_user_guide(USER_GUIDE_PLANNER))
bantuan_menu.add_command(label="Panduan Pengguna - Driver", command=lambda: show_user_guide(USER_GUIDE_DRIVER))
bantuan_menu.add_separator()
bantuan_menu.add_command(
    label="Periksa Pembaruan",
    command=lambda: check_update(ignore_skip=True, show_checkbox=False)
)
bantuan_menu.add_command(label="Tentang", command=show_about)
menu_bar.add_cascade(label="Bantuan", menu=bantuan_menu)

root.config(menu=menu_bar)

# --- Setup Tampilan Utama ---
update_title(root)
window_width, window_height = 400, 300
position_x = (root.winfo_screenwidth() - window_width) // 2
position_y = (root.winfo_screenheight() - window_height) // 2
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
root.resizable(False, False)

try:
    root.iconbitmap(resource_path("icon.ico"))
except tk.TclError:
    pass
    
frame = tk.Frame(root)
frame.pack(expand=True)
button_font = ("Arial", 14, "bold")

buttons_config = [
    ("Routing Summary", auto_routing_summary_main),
    ("Delivery Summary", auto_delivery_summary_main),
    ("Start-Finish Time", start_finish_time_main),
]

main_buttons = []
for i, (text, command) in enumerate(buttons_config):
    btn = tk.Button(frame, text=text, command=command, font=button_font, padx=20, pady=10, width=15)
    btn.grid(row=i, column=0, padx=10, pady=10)
    main_buttons.append(btn)

footer_label = tk.Label(root, text=INFO_MESSAGES["APP_BUILD_BY"], font=("Arial", 8), fg="gray")
footer_label.pack(side="bottom", pady=5)

# --- Tampilkan Window dan Jalankan Proses Latar Belakang ---
root.deiconify() 
periksa_konfigurasi_awal(root)
atur_visibilitas_menu(menu_bar)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.after(500, check_update) 
run_sync_in_background(root)

root.mainloop()
