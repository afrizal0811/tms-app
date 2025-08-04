# @GUI/apps.py (KODE YANG DIMODIFIKASI)

from tkinter import ttk
import requests
import sys
import threading
import tkinter as tk
import webbrowser
from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK
from utils.function import (
    CONFIG_PATH,
    ensure_config_exists,
    load_config,
    load_constants,
    resource_path,
    save_json_data,
    show_error_message,
    show_info_message,
)
from utils.messages import (
    ERROR_MESSAGES,
    INFO_MESSAGES
)

# --- Impor Modul Aplikasi ---
from modules.Routing_Summary.apps import main as routing_summary_main
from modules.Delivery_Summary.apps import main as delivery_summary_main
from modules.Auto_Delivery_Summary.apps import main as auto_delivery_summary_main
from modules.Start_Finish_Time.apps import main as start_finish_time_main
from modules.Sync_Driver.apps import main as sync_driver_main
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

def pilih_lokasi(parent_window):
    """Menampilkan GUI modal untuk memilih lokasi cabang."""
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
    
    def on_select():
        selected = combo.get()
        if selected in LOKASI_DISPLAY:
            kode = LOKASI_DISPLAY[selected]
            config_data['lokasi'] = kode
            save_json_data(config_data, CONFIG_PATH)
            dialog.destroy()

    tk.Button(dialog, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)
    dialog.transient(parent_window)
    dialog.grab_set()
    parent_window.wait_window(dialog)

def pilih_pengguna_awal(parent_window):
    """Menjalankan proses pemilihan pengguna menggunakan modul Check_User."""
    check_user_main(parent_window)

def on_closing():
    try:
        root.destroy()
    finally:
        import os
        os._exit(0)

def check_update():
    """Memeriksa versi baru dari URL remote."""
    try:
        response = requests.get(REMOTE_VERSION_URL, timeout=5)
        response.raise_for_status()
        latest_version = response.text.strip()
        if latest_version > CURRENT_VERSION:
            message = INFO_MESSAGES["UPDATE_AVAILABLE"].format(latest_version=latest_version, current_version=CURRENT_VERSION)
            if show_info_message("Update Tersedia", message):
                webbrowser.open(DOWNLOAD_LINK)
    except requests.exceptions.RequestException:
        pass

def periksa_konfigurasi_awal(parent_window):
    """
    Memeriksa apakah lokasi dan pengguna sudah diatur.
    """
    config = load_config()
    if not config or not config.get("lokasi"):
        show_info_message("Setup Awal", INFO_MESSAGES["WELCOME_SETUP"])
        pilih_lokasi(parent_window)
        update_title(parent_window)
        config = load_config()
    
    if not config or not config.get("user_checked"):
        show_info_message("Setup Akun", INFO_MESSAGES["USER_SETUP"])
        pilih_pengguna_awal(parent_window)
        if not (load_config() or {}).get("user_checked"):
            show_error_message("Setup Tidak Lengkap", ERROR_MESSAGES["USER_SETUP_CANCELED"])
            on_closing()

def atur_visibilitas_menu(menu_bar):
    """Mengatur visibilitas item menu berdasarkan role pengguna."""
    config = load_config()
    constants = load_constants()

    user_info = config.get("user_checked", {})
    user_role_id = user_info.get("role_id")
    
    restricted_roles = constants.get("restricted_role_ids", {})
    
    restricted_role_id_list = list(restricted_roles.values())

    try:
        if user_role_id and user_role_id in restricted_role_id_list:
            pengaturan_menu.delete("Ganti Lokasi Cabang")
    except tk.TclError:
        pass

def run_sync_in_background(root_window):
    """Menjalankan proses sinkronisasi driver di background thread."""
    loading_window = tk.Toplevel(root_window)
    loading_window.title("Loading")
    loading_window.geometry("300x100")
    x, y = root_window.winfo_x() + (root_window.winfo_width() // 2) - 150, root_window.winfo_y() + (root_window.winfo_height() // 2) - 50
    loading_window.geometry(f"+{x}+{y}")
    loading_window.transient(root_window)
    loading_window.grab_set()
    ttk.Label(loading_window, text="Sinkronisasi data driver...", font=("Arial", 12)).pack(pady=20)
    progress = ttk.Progressbar(loading_window, mode='indeterminate')
    progress.pack(pady=10, padx=20, fill=tk.X)
    progress.start()
    
    for button in main_buttons: button.config(state='disabled')
    pengaturan_menu.entryconfig("Sinkronisasi Driver", state="disabled")
    proses_menu.entryconfig("Routing Summary", state="disabled")
    proses_menu.entryconfig("Delivery Summary", state="disabled")
    proses_menu.entryconfig("Vehicles Data", state="disabled") # Nonaktifkan menu baru

    def on_sync_complete():
        # --- [PERBAIKAN] ---
        # Hentikan progress bar sebelum menutup jendela untuk menghindari error
        if loading_window.winfo_exists():
            progress.stop()
            loading_window.destroy()

        for button in main_buttons: button.config(state='normal')
        pengaturan_menu.entryconfig("Sinkronisasi Driver", state="normal")
        proses_menu.entryconfig("Routing Summary", state="normal")
        proses_menu.entryconfig("Delivery Summary", state="normal")
        proses_menu.entryconfig("Vehicles Data", state="normal") 

    def thread_target():
        try:
            sync_driver_main()
        finally:
            root_window.after(0, on_sync_complete)
            
    sync_thread = threading.Thread(target=thread_target, daemon=True)
    sync_thread.start()


# ==============================================================================
# ALUR UTAMA APLIKASI GUI
# ==============================================================================

# --- Setup Window Utama ---
root = tk.Tk()
root.withdraw() 

def ganti_lokasi():
    pilih_lokasi(root)
    update_title(root)

# --- Setup Menu Bar ---
menu_bar = tk.Menu(root)

proses_menu = tk.Menu(menu_bar, tearoff=0)
proses_menu.add_command(label="Routing Summary", command=routing_summary_main)
proses_menu.add_command(label="Delivery Summary", command=delivery_summary_main)
proses_menu.add_separator()
proses_menu.add_command(label="Vehicles Data", command=vehicles_data_main)
menu_bar.add_cascade(label="Proses", menu=proses_menu)

pengaturan_menu = tk.Menu(menu_bar, tearoff=0)
pengaturan_menu.add_command(label="Ganti Lokasi Cabang", command=ganti_lokasi)
pengaturan_menu.add_command(label="Sinkronisasi Driver", command=lambda: run_sync_in_background(root))

def show_about():
    show_info_message(
        "Tentang Aplikasi",
        INFO_MESSAGES["APP_VERSION"].format(version=CURRENT_VERSION)
        + "\n\n"
        + INFO_MESSAGES["APP_BUILD_BY"]
    )


pengaturan_menu.add_separator()
pengaturan_menu.add_command(label="Tentang", command=show_about)
menu_bar.add_cascade(label="Pengaturan", menu=pengaturan_menu)

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
    ("Auto Routing Summary", auto_routing_summary_main),
    ("Auto Delivery Summary", auto_delivery_summary_main),
    ("Start-Finish Time", start_finish_time_main),
]

main_buttons = []
for i, (text, command) in enumerate(buttons_config):
    btn = tk.Button(frame, text=text, command=command, font=button_font, padx=20, pady=10, width=20)
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
root.after(1500, lambda: run_sync_in_background(root))

root.mainloop()
