import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys
import requests
import webbrowser
import json
from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK

from Routing_Summary import apps as routing_summary
from Start_Finish_Time import apps as start_finish_time
from Delivery_Summary import apps as delivery_summary

# Dictionary untuk memetakan nama lengkap lokasi ke kodenya
lokasi_dict_nama_ke_kode = {
    "Sidoarjo": "plsda", "Jakarta": "pljkt", "Bandung": "plbdg",
    "Semarang": "plsmg", "Yogyakarta": "plygy", "Malang": "plmlg",
    "Denpasar": "pldps", "Makasar": "plmks", "Jember": "pljbr"
}
# Dictionary terbalik untuk mencari nama dari kode (untuk judul)
kode_ke_lokasi = {v: k for k, v in lokasi_dict_nama_ke_kode.items()}


def get_base_path():
    """Mendapatkan path dasar (base path) baik untuk script maupun executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)

def update_title(root_window):
    """Membaca konfigurasi dan memperbarui judul window utama."""
    config_path = os.path.join(get_base_path(), "config.json")
    title = "TMS Data Processing"
    try:
        if os.path.exists(config_path):
            with open(config_path, "r") as f:
                data = json.load(f)
                kode = data.get("lokasi")
                if kode and kode in kode_ke_lokasi:
                    nama_lokasi = kode_ke_lokasi[kode]
                    title += f" - {nama_lokasi}"
    except (json.JSONDecodeError, IOError):
        pass
    finally:
        root_window.title(title)

# --- PERBAIKAN 1: Modifikasi `pilih_lokasi` untuk menggunakan Toplevel modal ---
def pilih_lokasi(parent_window):
    """Menampilkan GUI modal untuk memilih lokasi cabang (default sesuai config)."""
    lokasi_dict_display = {
        "01. Sidoarjo": "plsda", "02. Jakarta": "pljkt", "03. Bandung": "plbdg",
        "04. Semarang": "plsmg", "05. Yogyakarta": "plygy", "06. Malang": "plmlg",
        "07. Denpasar": "pldps", "08. Makasar": "plmks", "09. Jember": "pljbr"
    }
    reverse_dict = {v: k for k, v in lokasi_dict_display.items()}  # "plsda" -> "01. Sidoarjo"
    config_path = os.path.join(get_base_path(), "config.json")

    # Ambil lokasi saat ini dari config.json
    selected_display_name = list(lokasi_dict_display.keys())[0]  # default ke yang pertama
    if os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                data = json.load(f)
                kode_lokasi = data.get("lokasi", "")
                if kode_lokasi in reverse_dict:
                    selected_display_name = reverse_dict[kode_lokasi]
        except (json.JSONDecodeError, IOError):
            pass

    # --- UI dialog ---
    dialog = tk.Toplevel(parent_window)
    dialog.title("Pilih Lokasi Cabang")

    lebar = 350
    tinggi = 180
    x = (dialog.winfo_screenwidth() - lebar) // 2
    y = (dialog.winfo_screenheight() - tinggi) // 2
    dialog.geometry(f"{lebar}x{tinggi}+{x}+{y}")

    tk.Label(dialog, text="Pilih Lokasi Cabang:", font=("Arial", 14)).pack(pady=10)

    selected_var = tk.StringVar(value=selected_display_name)
    combo = ttk.Combobox(dialog, values=list(lokasi_dict_display.keys()), textvariable=selected_var, font=("Arial", 12), state="readonly")
    combo.pack(pady=10)
    combo.set(selected_display_name)

    def on_select():
        selected = combo.get()
        if selected in lokasi_dict_display:
            kode = lokasi_dict_display[selected]
            with open(config_path, "w") as f:
                json.dump({"lokasi": kode}, f)
            dialog.destroy()

    tk.Button(dialog, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)

    dialog.transient(parent_window)
    dialog.grab_set()
    parent_window.wait_window(dialog)

# ... (fungsi button_action tetap sama) ...
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

# --- PERBAIKAN 3: Modifikasi `periksa_konfigurasi_awal` ---
def periksa_konfigurasi_awal(parent_window):
    """Memeriksa config.json saat startup. Jika tidak ada, panggil GUI pilih lokasi."""
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
        # Panggil pilih_lokasi dengan parent_window
        pilih_lokasi(parent_window)

# --- Alur Utama Aplikasi ---
# 1. Buat jendela utama
root = tk.Tk()
root.withdraw() # Sembunyikan dulu jendela utama

# --- PERBAIKAN 4: Modifikasi `ganti_lokasi` untuk memanggil dialog modal ---
def ganti_lokasi():
    pilih_lokasi(root)   # Panggil dialog dengan root sebagai parent
    update_title(root)   # Perbarui judul SETELAH dialog ditutup

# 1. Setup Menu
menu_bar = tk.Menu(root)
pengaturan_menu = tk.Menu(menu_bar, tearoff=0)
pengaturan_menu.add_command(label="Ganti Lokasi Cabang", command=ganti_lokasi)
menu_bar.add_cascade(label="Pengaturan", menu=pengaturan_menu)
root.config(menu=menu_bar)

# 2. Atur judul berdasarkan konfigurasi yang ada
update_title(root)

# 4. Atur geometri dan tampilkan jendela utama
window_width = 400
window_height = 300
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
position_x = (screen_width // 2) - (window_width // 2)
position_y = (screen_height // 2) - (window_height // 2)
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# 4. Tampilkan jendela utama terlebih dahulu
root.deiconify()

# 5. Lakukan pemeriksaan konfigurasi awal
periksa_konfigurasi_awal(root)

# 6. Setup Frame dan Tombol
frame = tk.Frame(root)
frame.pack(expand=True)
button_font = ("Arial", 14, "bold")
buttons_config = [
    ("Routing Summary", button1_action, 0, 0, "normal"),
    ("Delivery Summary", button2_action, 1, 0, "normal"),
    ("Start-Finish Time", button3_action, 2, 0, "normal"),
]
for text, command, row, col, state in buttons_config:
    btn = tk.Button(frame, text=text, command=command, font=button_font, padx=20, pady=10, width=15, state=state)
    btn.grid(row=row, column=col, padx=10, pady=10)

# 7. Tampilkan jendela utama dan jalankan aplikasi

root.protocol("WM_DELETE_WINDOW", on_closing)
# Cek update setelah jendela ditampilkan
root.after(1000, check_update)
root.mainloop()
