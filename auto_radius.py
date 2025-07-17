import pyautogui
import time
import subprocess
import tkinter as tk
from tkinter import messagebox
import threading
import os
import sys

# =================================================================================
# BAGIAN 1: SEMUA LOGIKA AUTOMASI (TIDAK DIUBAH)
# =================================================================================

def resource_path(relative_path):
    """ Fungsi untuk mendapatkan path absolut ke resource, berfungsi untuk mode dev dan PyInstaller """
    try:
        # PyInstaller membuat folder sementara dan menyimpan path di _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

# --- PENGATURAN GAMBAR MENGGUNAKAN FUNGSI BARU ---
DRIVER_OPTION_IMG = resource_path('driver_option.png')
DRIVER_CONFIRM_IMG = resource_path('driver_confirm.png')
START_RADIUS_BOX_IMG = resource_path('start_radius_box.png')
SAVE_BUTTON_IMG = resource_path('save.png')

# --- FUNGSI-FUNGSI PYAUTOGUI ---
def wait_for_image(image_path, timeout=15):
    start_time = time.time()
    while True:
        try:
            location = pyautogui.locateOnScreen(image_path, confidence=0.7)
            if location:
                return location
        except pyautogui.ImageNotFoundException:
            pass
        if time.time() - start_time > timeout:
            raise Exception(f"TIMEOUT: Tidak dapat menemukan gambar '{image_path}'")
        time.sleep(0.5)

def find_and_click(image_path, timeout=10):
    location = wait_for_image(image_path, timeout=timeout)
    pyautogui.click(pyautogui.center(location))

# --- FUNGSI UTAMA AUTOMASI ---
def run_automation(radius_value):
    """Fungsi ini berisi seluruh alur kerja automasi."""
    try:
        # 1. Buka URL
        chrome_path = 'C:/Program Files/Google/Chrome/Application/chrome.exe'
        url = 'https://web.mile.app/setting/permission'
        subprocess.Popen([chrome_path, '-new-window', url])
        wait_for_image(DRIVER_OPTION_IMG, timeout=20)

        # 2. Pilih role "DRIVER"
        find_and_click(DRIVER_OPTION_IMG)
        wait_for_image(DRIVER_CONFIRM_IMG)

        # 3. Scroll ke bawah
        pyautogui.scroll(-20000)

        # 4. Ubah nilai "Start radius"
        find_and_click(START_RADIUS_BOX_IMG)
        start_radius_area = wait_for_image(START_RADIUS_BOX_IMG)
        # Horisontal: 50% (tengah), Vertikal: 75% (lebih ke bawah)
        click_x_radius = start_radius_area.left + start_radius_area.width * 0.5
        click_y_radius = start_radius_area.top + start_radius_area.height * 0.75
        pyautogui.click(click_x_radius, click_y_radius)
        time.sleep(0.5)
        pyautogui.hotkey('ctrl', 'a')
        pyautogui.press('backspace')
        pyautogui.write(str(radius_value)) # Menggunakan nilai dari parameter

        # 5. Scroll ke atas
        pyautogui.scroll(20000)
        time.sleep(0.5)

        # 6. Klik tombol Save dengan metode presisi
        save_button_box = wait_for_image(SAVE_BUTTON_IMG)
        click_x = save_button_box.left + save_button_box.width * 0.65
        click_y = save_button_box.top + save_button_box.height * 0.8
        pyautogui.click(click_x, click_y)

    except Exception as e:
        error_type = type(e).__name__
        messagebox.showerror(f"Error: {error_type}", f"Program Gagal!\n\nPenyebab:\n{str(e)}")

# =================================================================================
# BAGIAN 2: KODE UNTUK GUI (TKINTER) - TELAH DIPERBARUI
# =================================================================================

def start_automation_thread(radius_value):
    """Fungsi untuk memulai automasi di thread baru agar GUI tidak macet."""
    btn_aktif.config(state="disabled")
    btn_nonaktif.config(state="disabled")
    
    automation_thread = threading.Thread(target=run_automation, args=(radius_value,))
    automation_thread.start()
    
    check_if_done(automation_thread)

def check_if_done(thread):
    """Memeriksa status thread setiap 100ms."""
    if thread.is_alive():
        root.after(100, lambda: check_if_done(thread))
    else:
        btn_aktif.config(state="normal")
        btn_nonaktif.config(state="normal")

# --- Pengaturan Tampilan GUI ---
WINDOW_WIDTH = 350
WINDOW_HEIGHT = 250
TITLE_FONT = ("Helvetica", 20, "bold")
BUTTON_FONT = ("Helvetica", 14, "bold")

# Membuat jendela utama
root = tk.Tk()
root.title("Radius Automation")
root.resizable(False, False) # Mencegah ukuran jendela diubah

# --- Logika untuk menempatkan jendela di tengah layar ---
screen_width = root.winfo_screenwidth()
screen_height = root.winfo_screenheight()
center_x = int(screen_width/2 - WINDOW_WIDTH / 2)
center_y = int(screen_height/2 - WINDOW_HEIGHT / 2)
root.geometry(f'{WINDOW_WIDTH}x{WINDOW_HEIGHT}+{center_x}+{center_y}')

# Membuat label judul dengan font lebih besar
lbl_title = tk.Label(root, text="Pilih Aksi Radius", font=TITLE_FONT)
lbl_title.pack(pady=20) # Menambah padding vertikal

# Membuat tombol "Aktif" dengan font lebih besar
btn_aktif = tk.Button(
    root, 
    text="Aktif (Set 500)", 
    font=BUTTON_FONT,
    command=lambda: start_automation_thread(500),
    width=20, 
    height=2,
    bg="#4CAF50", # Warna hijau
    fg="white"    # Warna teks putih
)
btn_aktif.pack()

# Membuat tombol "Nonaktif" dengan font lebih besar
btn_nonaktif = tk.Button(
    root, 
    text="Nonaktif (Set 0)", 
    font=BUTTON_FONT,
    command=lambda: start_automation_thread(0),
    width=20, 
    height=2,
    bg="#f44336", # Warna merah
    fg="white"    # Warna teks putih
)
btn_nonaktif.pack(pady=10) # Menambah padding vertikal

# Menjalankan main loop GUI
root.mainloop()

