import tkinter as tk
from tkinter import messagebox, ttk
import os
import sys
import requests
import webbrowser
import json
from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK

from Routing_Summary import apps as routing_summary
from RO_vs_Real import apps as ro_vs_real
from Pending_SO import apps as pending_so
from Start_Finish_Time import apps as start_finish_time

def get_base_path():
    """Mendapatkan path dasar (base path) baik untuk script maupun executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(__file__)
    
def pilih_lokasi():
    """Menampilkan GUI untuk memilih lokasi cabang."""
    lokasi_dict = {
        "01. Sidoarjo": "plsda", "02. Jakarta": "pljkt", "03. Bandung": "plbdg",
        "04. Semarang": "plsmg", "05. Yogyakarta": "plygy", "06. Malang": "plmlg",
        "07. Denpasar": "pldps", "08. Makasar": "plmks", "09. Jember": "pljbr"
    }
    config_path = os.path.join(get_base_path(), "config.json")

    selected_value = None
    def on_select():
        nonlocal selected_value
        selected = combo.get()
        if selected in lokasi_dict:
            selected_value = lokasi_dict[selected]
            with open(config_path, "w") as f:
                json.dump({"lokasi": selected_value}, f)
            root.destroy()

    root = tk.Tk()
    root.title("Pilih Lokasi Cabang")
    lebar = 350
    tinggi = 180
    x = (root.winfo_screenwidth() - lebar) // 2
    y = (root.winfo_screenheight() - tinggi) // 2
    root.geometry(f"{lebar}x{tinggi}+{x}+{y}")

    tk.Label(root, text="Pilih Lokasi Cabang:", font=("Arial", 14)).pack(pady=10)
    combo = ttk.Combobox(root, values=list(lokasi_dict.keys()), font=("Arial", 12))
    combo.pack(pady=10)
    combo.current(0)
    tk.Button(root, text="Pilih", command=on_select, font=("Arial", 12)).pack(pady=10)
    root.mainloop()
    return selected_value

# Fungsi untuk tombol 1
def button1_action():
    routing_summary.main()

# Fungsi untuk tombol 2
def button2_action():
    ro_vs_real.main()

# Fungsi untuk tombol 3: Menjalankan truck_detail_routing.py
def button3_action():
    pending_so.main()

# Fungsi untuk tombol 4: Menjalankan truck_detail_task.py
def button4_action():
    start_finish_time.main()

# Fungsi untuk tombol 5
def button5_action():
    messagebox.showinfo("Info", "Tombol 5 di klik!")

# Fungsi untuk tombol 6
def button6_action():
    messagebox.showinfo("Info", "Tombol 6 di klik!")

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

has_closed = False  # Di atas, sebelum fungsi apa pun

def on_closing():
    global has_closed
    if not has_closed:
        has_closed = True
        try:
            root.destroy()
        except tk.TclError:
            pass

def check_update():
    try:
        response = requests.get(REMOTE_VERSION_URL)
        latest_version = response.text.strip()

        if latest_version != CURRENT_VERSION:
            result = messagebox.askyesno(
                "Update Tersedia",
                f"Versi terbaru: {latest_version}\nVersi kamu: {CURRENT_VERSION}\n\nMau buka halaman update?"
            )
            if result:
                webbrowser.open(DOWNLOAD_LINK)
    except Exception as e:
        messagebox.showerror("Gagal Cek Update", f"Gagal mengecek versi terbaru:\n{e}")

def periksa_konfigurasi_awal():
    """Memeriksa config.json saat startup. Jika tidak ada, panggil GUI pilih lokasi."""
    config_path = os.path.join(get_base_path(), "config.json")
    konfigurasi_valid = False

    if os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                data = json.load(f)
                # Pastikan key 'lokasi' ada dan tidak kosong
                if "lokasi" in data and data["lokasi"]:
                    konfigurasi_valid = True
        except (json.JSONDecodeError, IOError):
            # Jika file rusak atau tidak bisa dibaca, anggap tidak valid
            konfigurasi_valid = False

    # Jika setelah semua pengecekan konfigurasi tetap tidak valid,
    # maka tampilkan paksa jendela pemilihan lokasi.
    if not konfigurasi_valid:
        messagebox.showinfo("Setup Awal", "Silakan pilih lokasi cabang Anda terlebih dahulu.")
        pilih_lokasi()

# Membuat jendela utama
root = tk.Tk()
root.title("TMS Data Processing")

# Fungsi untuk menu "Ganti Lokasi Cabang"
def ganti_lokasi():
    pilih_lokasi()
    messagebox.showinfo("Informasi", "Lokasi cabang berhasil diperbarui.\nSilakan restart aplikasi untuk menerapkan perubahan.")

# Membuat menu bar
menu_bar = tk.Menu(root)

# Menu 'Pengaturan'
pengaturan_menu = tk.Menu(menu_bar, tearoff=0)
pengaturan_menu.add_command(label="Ganti Lokasi Cabang", command=ganti_lokasi)
menu_bar.add_cascade(label="Pengaturan", menu=pengaturan_menu)

# Pasang menu bar ke window
root.config(menu=menu_bar)

# Cek update otomatis setelah 1 detik app dibuka
root.after(10, check_update)

# Menentukan ukuran jendela
window_width = 700
window_height = 350

# Mendapatkan ukuran layar
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()

# Menentukan posisi tengah
position_x = (screen_width // 2) - (window_width // 2)
position_y = (screen_height // 2) - (window_height // 2)

# Menentukan ukuran dan posisi jendela
root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")

# Frame untuk menyusun tombol dalam grid
frame = tk.Frame(root)
frame.pack(expand=True)

# Styling tombol
button_font = ("Arial", 14, "bold")

# Membuat tombol 1
button1 = tk.Button(frame, text="Routing Summary", command=button1_action, font=button_font, padx=20, pady=10, width=15)
button1.grid(row=0, column=0, padx=10, pady=10)

# Membuat tombol 2
button2 = tk.Button(frame, text="RO vs Real", command=button2_action, font=button_font, padx=20, pady=10, width=15)
button2.grid(row=0, column=1, padx=10, pady=10)

# Membuat tombol 3
button3 = tk.Button(frame, text="Pending SO", command=button3_action, font=button_font, padx=20, pady=10, width=15)
button3.grid(row=1, column=0, padx=10, pady=10)

# Membuat tombol 3
button3 = tk.Button(frame, text="Start-Finish Time", command=button4_action, font=button_font, padx=20, pady=10, width=15)
button3.grid(row=1, column=1, padx=10, pady=10)

# Membuat tombol 4
button5 = tk.Button(frame, text="Tombol Disabled", command=button5_action, font=button_font, padx=20, pady=10, width=15, state="disabled")
button5.grid(row=2, column=0, padx=10, pady=10) 

# Membuat tombol 6
button6 = tk.Button(frame, text="Tombol Disabled", command=button6_action, font=button_font, padx=20, pady=10, width=15, state="disabled")
button6.grid(row=2, column=1, padx=10, pady=10) 

periksa_konfigurasi_awal()
# Menjalankan loop utama aplikasi
try:
    root.protocol("WM_DELETE_WINDOW", on_closing)
except tk.TclError:
    pass  # Abaikan jika root sudah dihancurkan
root.mainloop()
