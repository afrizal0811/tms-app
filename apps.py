import tkinter as tk
from tkinter import messagebox
import os
import sys
import requests
import webbrowser
from version import CURRENT_VERSION, REMOTE_VERSION_URL, DOWNLOAD_LINK

from Combine_Routing import apps as combine_routing
from RO_vs_Real import apps as ro_vs_real
from Pending_SO import apps as pending_so
from Truck_Detail import apps as truck_detail
from Truck_Usage import apps as truck_usage

# Fungsi untuk tombol 1
def button1_action():
    combine_routing.main()

# Fungsi untuk tombol 2
def button2_action():
    ro_vs_real.main()

# Fungsi untuk tombol 3: Menjalankan truck_detail_routing.py
def button3_action():
    pending_so.main()

# Fungsi untuk tombol 4: Menjalankan truck_detail_task.py
def button4_action():
    truck_detail.main()

# Fungsi untuk tombol 5
def button5_action():
    truck_usage.main()

# Fungsi untuk tombol 6
def button6_action():
    messagebox.showinfo("Info", "Tombol 6 di klik!")

def resource_path(relative_path):
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)

def on_closing():
    root.destroy()
    sys.exit()  # pastikan exit beneran clean

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


# Membuat jendela utama
root = tk.Tk()
root.title("TMS Data Processing")

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
button1 = tk.Button(frame, text="Combine Routing", command=button1_action, font=button_font, padx=20, pady=10, width=15)
button1.grid(row=0, column=0, padx=10, pady=10)

# Membuat tombol 2
button2 = tk.Button(frame, text="RO vs Real", command=button2_action, font=button_font, padx=20, pady=10, width=15)
button2.grid(row=0, column=1, padx=10, pady=10)

# Membuat tombol 3
button3 = tk.Button(frame, text="Pending SO", command=button3_action, font=button_font, padx=20, pady=10, width=15)
button3.grid(row=1, column=0, padx=10, pady=10)

# Membuat tombol 3
button3 = tk.Button(frame, text="Truck Detail", command=button4_action, font=button_font, padx=20, pady=10, width=15)
button3.grid(row=1, column=1, padx=10, pady=10)

# Membuat tombol 4
button5 = tk.Button(frame, text="Truck Usage", command=button5_action, font=button_font, padx=20, pady=10, width=15)
button5.grid(row=2, column=0, padx=10, pady=10) 

# Membuat tombol 6
button6 = tk.Button(frame, text="Tombol Disabled", command=button6_action, font=button_font, padx=20, pady=10, width=15, state="disabled")
button6.grid(row=2, column=1, padx=10, pady=10) 

# Menjalankan loop utama aplikasi
root.protocol("WM_DELETE_WINDOW", on_closing)
root.mainloop()
