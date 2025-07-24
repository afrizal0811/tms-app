import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys
import requests
import webbrowser
import json
import threading

from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK

from modules.Routing_Summary import apps as routing_summary
from modules.Delivery_Summary import apps as delivery_summary
from modules.Start_Finish_Time import apps as start_finish_time
from modules.Sync_Driver import apps as sync_driver

def get_base_path():
    """Mendapatkan path dasar (base path) baik untuk script maupun executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)

def resource_path(relative_path):
    """Mendapatkan path absolut ke resource"""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def load_constants():
    """Memuat data dari constant.json saat startup."""
    # Gunakan fungsi resource_path yang sudah diperbaiki
    constants_path = resource_path("modules/constant.json")
    try:
        with open(constants_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        messagebox.showerror("Kritikal Error", f"File 'constant.json' tidak ditemukan.\nPastikan file tersebut ada di folder yang sama dengan aplikasi atau sudah dibundel dengan benar.")
        sys.exit()
    except json.JSONDecodeError:
        messagebox.showerror("Kritikal Error", "File 'constant.json' memiliki format yang salah.")
        sys.exit()
    except Exception as e:
        messagebox.showerror("Kritikal Error", f"Gagal memuat 'constant.json':\n{e}")
        sys.exit()

# Muat konstanta di awal dan definisikan variabel global
CONSTANTS = load_constants()
LOKASI_MAPPING = CONSTANTS.get('lokasi_mapping', {})
LOKASI_DISPLAY = CONSTANTS.get('lokasi_display', {})
KODE_KE_LOKASI = {v: k for k, v in LOKASI_MAPPING.items()}


def update_title(root_window):
    """Membaca konfigurasi dan memperbarui judul window utama."""
    config_path = os.path.join(get_base_path(), "config.json")
    title = "TMS Data Processing"
    try:
        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                data = json.load(f)
                kode = data.get("lokasi")
                if kode and kode in KODE_KE_LOKASI:
                    nama_lokasi = KODE_KE_LOKASI[kode]
                    title += f" - {nama_lokasi}"
    except (json.JSONDecodeError, IOError):
        pass
    finally:
        root_window.title(title)

def pilih_lokasi(parent_window):
    """Menampilkan GUI modal untuk memilih lokasi cabang (default sesuai config)."""
    reverse_dict = {v: k for k, v in LOKASI_DISPLAY.items()}
    config_path = os.path.join(get_base_path(), "config.json")
    selected_display_name = list(LOKASI_DISPLAY.keys())[0]
    
    config_data = {}
    if os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                config_data = json.load(f)
                kode_lokasi = config_data.get("lokasi", "")
                if kode_lokasi in reverse_dict:
                    selected_display_name = reverse_dict[kode_lokasi]
        except (json.JSONDecodeError, IOError):
            pass 

    dialog = tk.Toplevel(parent_window)
    dialog.title("Pilih Lokasi Cabang")
    lebar, tinggi = 350, 180
    x = (dialog.winfo_screenwidth() - lebar) // 2
    y = (dialog.winfo_screenheight() - tinggi) // 2
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
            try:
                with open(config_path, "w") as f:
                    json.dump(config_data, f, indent=2) 
            except IOError as e:
                messagebox.showerror("Error", f"Gagal menyimpan konfigurasi:\n{e}")
            dialog.destroy()

    tk.Button(dialog, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)
    dialog.transient(parent_window)
    dialog.grab_set()
    parent_window.wait_window(dialog)

def button1_action(): routing_summary.main()
def button2_action(): delivery_summary.main()
def button3_action(): start_finish_time.main()

def on_closing():
    try:
        root.destroy()
    finally:
        os._exit(0)

def check_update():
    try:
        response = requests.get(REMOTE_VERSION_URL)
        latest_version = response.text.strip()
        if latest_version > CURRENT_VERSION:
            if messagebox.askyesno("Update Tersedia", f"Versi terbaru: {latest_version}\nVersi kamu: {CURRENT_VERSION}\n\nMau buka halaman update?"):
                webbrowser.open(DOWNLOAD_LINK)
    except Exception as e:
        messagebox.showerror("Gagal Cek Update", f"Gagal mengecek versi terbaru:\n{e}")

def periksa_konfigurasi_awal(parent_window):
    config_path = os.path.join(get_base_path(), "config.json")
    konfigurasi_valid = False
    if os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                data = json.load(f)
                if "lokasi" in data and data["lokasi"]:
                    konfigurasi_valid = True
        except (json.JSONDecodeError, IOError):
            konfigurasi_valid = False
    if not konfigurasi_valid:
        messagebox.showinfo("Setup Awal", "Silakan pilih lokasi cabang Anda terlebih dahulu.")
        pilih_lokasi(parent_window)

def toggle_main_buttons(state):
    for button in main_buttons:
        button.config(state=state)

def run_sync_in_background(root_window):
    loading_window = tk.Toplevel(root_window)
    loading_window.title("Loading")
    loading_window.geometry("300x100")
    x = root_window.winfo_x() + (root_window.winfo_width() // 2) - 150
    y = root_window.winfo_y() + (root_window.winfo_height() // 2) - 50
    loading_window.geometry(f"+{x}+{y}")
    loading_window.transient(root_window)
    loading_window.grab_set()
    ttk.Label(loading_window, text="Sinkronisasi sedang berjalan...", font=("Arial", 12)).pack(pady=20)
    progress = ttk.Progressbar(loading_window, mode='indeterminate')
    progress.pack(pady=10, padx=20, fill=tk.X)
    progress.start()
    toggle_main_buttons('disabled')

    def on_sync_complete(success, error_message=None):
        if loading_window:
            loading_window.destroy()
        toggle_main_buttons('normal')
        if not success and error_message:
            messagebox.showerror("Error Sinkronisasi", f"Terjadi kegagalan saat sinkronisasi:\n\n{error_message}")
        try:
            pengaturan_menu.entryconfig("Sinkronisasi Driver", state="normal")
        except tk.TclError:
            pass 

    def thread_target():
        try:
            base_path = get_base_path()
            sync_driver.main(base_path)
            root_window.after(0, on_sync_complete, True)
        except Exception as e:
            root_window.after(0, on_sync_complete, False, str(e))
            
    try:
        pengaturan_menu.entryconfig("Sinkronisasi Driver", state="disabled")
    except tk.TclError:
        pass
    sync_thread = threading.Thread(target=thread_target, daemon=True)
    sync_thread.start()

# --- Alur Utama Aplikasi ---
root = tk.Tk()
root.withdraw()

def ganti_lokasi():
    pilih_lokasi(root)
    update_title(root)

menu_bar = tk.Menu(root)
pengaturan_menu = tk.Menu(menu_bar, tearoff=0)
pengaturan_menu.add_command(label="Ganti Lokasi Cabang", command=ganti_lokasi)
pengaturan_menu.add_separator()
pengaturan_menu.add_command(label="Sinkronisasi Driver", command=lambda: run_sync_in_background(root))
menu_bar.add_cascade(label="Pengaturan", menu=pengaturan_menu)
root.config(menu=menu_bar)

update_title(root)

window_width, window_height = 400, 300
position_x = (root.winfo_screenwidth() // 2) - (window_width // 2)
position_y = (root.winfo_screenheight() // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

root.deiconify()
periksa_konfigurasi_awal(root)

frame = tk.Frame(root)
frame.pack(expand=True)
button_font = ("Arial", 14, "bold")
buttons_config = [
    ("Routing Summary", button1_action, 0, 0, "normal"),
    ("Delivery Summary", button2_action, 1, 0, "normal"),
    ("Start-Finish Time", button3_action, 2, 0, "normal"),
]
main_buttons = []
for text, command, row, col, state in buttons_config:
    btn = tk.Button(frame, text=text, command=command, font=button_font, padx=20, pady=10, width=15, state=state)
    btn.grid(row=row, column=col, padx=10, pady=10)
    main_buttons.append(btn)

root.after(1000, lambda: run_sync_in_background(root))
root.protocol("WM_DELETE_WINDOW", on_closing)
root.after(2000, check_update)
root.mainloop()
