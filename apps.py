# @GUI/apps.py (KODE YANG DIMODIFIKASI)

import tkinter as tk
from tkinter import messagebox, ttk
import sys
import requests
import webbrowser
import threading

# --- Impor Lokal ---
from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK

# 1. Impor fungsi terpusat dari shared_utils
from modules.shared_utils import (
    load_config,
    load_constants,
    save_json_data,
    ensure_config_exists,
    resource_path,
    CONFIG_PATH # Diperlukan untuk menyimpan config
)

# --- Impor Modul Aplikasi ---
# Fungsi 'main' dari modul manual yang akan dipanggil dari menu
from modules.Routing_Summary.apps import main as routing_summary_main
from modules.Delivery_Summary.apps import main as delivery_summary_main

# Fungsi 'main' untuk tombol di halaman utama
# === PERUBAHAN DIMULAI DI SINI ===
# from modules.Auto_Routing_Summary.apps import main as auto_routing_summary_main # Baris ini dinonaktifkan
# === AKHIR DARI PERUBAHAN ===
from modules.Auto_Delivery_Summary.apps import main as auto_delivery_summary_main
from modules.Start_Finish_Time.apps import main as start_finish_time_main
from modules.Sync_Driver.apps import main as sync_driver_main


ensure_config_exists()
# ==============================================================================
# FUNGSI BANTUAN LOKAL DAN KONfigurasi AWAL
# ==============================================================================

# === PERUBAHAN DIMULAI DI SINI ===
def show_wip_popup():
    """Menampilkan pop-up bahwa fitur masih dalam pengembangan."""
    messagebox.showinfo("Segera Hadir", "Fitur ini masih dalam tahap pengembangan dan akan segera tersedia.")
# === AKHIR DARI PERUBAHAN ===

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
            if messagebox.askyesno("Update Tersedia", f"Versi baru: {latest_version}\nVersi Anda: {CURRENT_VERSION}\n\nBuka halaman update?"):
                webbrowser.open(DOWNLOAD_LINK)
    except requests.exceptions.RequestException:
        pass

def periksa_konfigurasi_awal(parent_window):
    """Memeriksa apakah lokasi sudah diatur saat pertama kali membuka aplikasi."""
    config = load_config()
    if not config or not config.get("lokasi"):
        messagebox.showinfo("Setup Awal", "Selamat datang! Silakan pilih lokasi cabang Anda terlebih dahulu.")
        pilih_lokasi(parent_window)
        update_title(parent_window) 

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
    manual_menu.entryconfig("Routing Summary", state="disabled")
    manual_menu.entryconfig("Delivery Summary", state="disabled")

    def on_sync_complete():
        if loading_window.winfo_exists(): loading_window.destroy()
        for button in main_buttons: button.config(state='normal')
        pengaturan_menu.entryconfig("Sinkronisasi Driver", state="normal")
        manual_menu.entryconfig("Routing Summary", state="normal")
        manual_menu.entryconfig("Delivery Summary", state="normal")

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

# Menu "Manual"
manual_menu = tk.Menu(menu_bar, tearoff=0)
manual_menu.add_command(label="Routing Summary", command=routing_summary_main)
manual_menu.add_command(label="Delivery Summary", command=delivery_summary_main)
menu_bar.add_cascade(label="Manual", menu=manual_menu)

# Menu Pengaturan
pengaturan_menu = tk.Menu(menu_bar, tearoff=0)
pengaturan_menu.add_command(label="Ganti Lokasi Cabang", command=ganti_lokasi)
pengaturan_menu.add_command(label="Sinkronisasi Driver", command=lambda: run_sync_in_background(root))
menu_bar.add_cascade(label="Pengaturan", menu=pengaturan_menu)

def show_about():
    messagebox.showinfo(
        "Tentang Aplikasi",
        f"TMS Data Processing\nVersi: {CURRENT_VERSION}\n\nDibuat oleh: Afrizal Maulana - EDP © 2025"
    )

help_menu = tk.Menu(menu_bar, tearoff=0)
help_menu.add_command(label="Tentang", command=show_about)
menu_bar.add_cascade(label="Bantuan", menu=help_menu)
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
    None
    
frame = tk.Frame(root)
frame.pack(expand=True)
button_font = ("Arial", 14, "bold")

# === PERUBAHAN DIMULAI DI SINI ===
# Definisikan ulang tombol untuk memanggil pop-up untuk Auto Routing
buttons_config = [
    ("Auto Routing Summary", show_wip_popup), # Menggunakan fungsi pop-up
    ("Auto Delivery Summary", auto_delivery_summary_main),
    ("Start-Finish Time", start_finish_time_main),
]
# === AKHIR DARI PERUBAHAN ===

main_buttons = []
for i, (text, command) in enumerate(buttons_config):
    btn = tk.Button(frame, text=text, command=command, font=button_font, padx=20, pady=10, width=20)
    btn.grid(row=i, column=0, padx=10, pady=10)
    main_buttons.append(btn)

footer_label = tk.Label(root, text="Dibuat oleh: Afrizal Maulana - EDP © 2025", font=("Arial", 8), fg="gray")
footer_label.pack(side="bottom", pady=5)

# --- Tampilkan Window dan Jalankan Proses Latar Belakang ---
root.deiconify() 
periksa_konfigurasi_awal(root)

root.protocol("WM_DELETE_WINDOW", on_closing)
root.after(500, check_update) 
root.after(1500, lambda: run_sync_in_background(root))

root.mainloop()