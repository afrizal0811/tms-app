# modules/shared_utils.py

import os
import sys
import json
import subprocess
import pandas as pd
from tkinter import filedialog, messagebox
import tkinter as tk

# =============================================================================
# PENGELOLAAN PATH TERPUSAT
# =============================================================================

def get_base_path():
    """Mendapatkan path dasar (@GUI) baik untuk script maupun executable."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

BASE_PATH = get_base_path()
CONFIG_PATH = os.path.join(BASE_PATH, "config.json")
MASTER_JSON_PATH = os.path.join(BASE_PATH, "master.json")
CONSTANTS_PATH = os.path.join(BASE_PATH, "constant.json")


# =============================================================================
# FUNGSI UTILITAS UMUM
# =============================================================================

# Ganti fungsi lama di modules/shared_utils.py dengan yang ini

# modules/shared_utils.py

def get_config_path():
    """Mengembalikan path config.json di samping executable atau script utama."""
    if getattr(sys, 'frozen', False):
        # Saat sudah dibundle
        return os.path.join(os.path.dirname(sys.executable), "config.json")
    else:
        # Saat jalan normal
        return os.path.join(get_base_path(), "config.json")

CONFIG_PATH = get_config_path()

def ensure_config_exists(default_data=None):
    """Membuat config.json jika belum ada."""
    if not os.path.exists(CONFIG_PATH):
        if default_data is None:
            default_data = {"lokasi": ""}
        save_json_data(default_data, CONFIG_PATH)

def get_base_path():
    """Mendapatkan path dasar (@GUI) tempat .exe berada."""
    if getattr(sys, 'frozen', False):
        return os.path.dirname(sys.executable)
    else:
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

# 1. TAMBAHKAN FUNGSI RESOURCE_PATH DI SINI
def resource_path(relative_path):
    """Mendapatkan path absolut ke resource, baik di dalam bundle _MEIPASS atau lokal."""
    try:
        # PyInstaller membuat folder sementara dan menyimpan path-nya di sys._MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # Jika tidak di-bundle, gunakan path biasa ke folder root proyek
        base_path = get_base_path()
    return os.path.join(base_path, relative_path)

def load_json_data(file_path):
    """Fungsi generik untuk membaca data dari file JSON."""
    if not os.path.exists(file_path):
        messagebox.showerror("File Tidak Ditemukan", f"File '{os.path.basename(file_path)}' tidak ditemukan.")
        return None
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        messagebox.showerror("Error Membaca File", f"Gagal memuat file '{os.path.basename(file_path)}'. File mungkin rusak.\n\nError: {e}")
        return None

def load_constants():
    """Memuat constant.json dari dalam bundle menggunakan resource_path."""
    # Mencari constant.json di dalam _MEIPASS (jika di-bundle) atau di root proyek
    path_to_constant = resource_path("constant.json")
    return load_json_data(path_to_constant)

def load_config():
    """Memuat config.json dari luar (di samping .exe). TIDAK BERUBAH."""
    return load_json_data(CONFIG_PATH)

def load_master_data(lokasi_cabang=None):
    """Memuat master.json dari luar (di samping .exe). TIDAK BERUBAH."""
    df = load_json_data(MASTER_JSON_PATH)
    if df is None: return None
    # ... (sisa logika fungsi ini tidak berubah) ...
    try:
        df = pd.DataFrame(df)
        df.columns = [col.strip() for col in df.columns]
        if 'Email' not in df.columns or 'Driver' not in df.columns:
            raise ValueError("Key 'Email' dan/atau 'Driver' tidak ditemukan di master data")
        df['Email'] = df['Email'].astype(str).str.strip().str.lower()
        df['Driver'] = df['Driver'].astype(str).str.strip()
        if lokasi_cabang:
            df = df[df['Email'].str.contains(lokasi_cabang, case=False, na=False)].copy()
        return df
    except Exception as e:
        messagebox.showerror("Error Master Data", f"Terjadi kesalahan saat memproses master data:\n{e}")
        return None

def get_save_path(base_name="Laporan", extension=".xlsx"):
    """Membuka dialog untuk memilih folder penyimpanan. TIDAK BERUBAH."""
    root = tk.Tk()
    root.withdraw()
    folder = filedialog.askdirectory(title="Pilih Lokasi Untuk Menyimpan File Laporan")
    if not folder: return None
    save_path = os.path.join(folder, f"{base_name}{extension}")
    counter = 1
    while os.path.exists(save_path):
        save_path = os.path.join(folder, f"{base_name} ({counter}){extension}")
        counter += 1
    return save_path
def open_file_externally(filepath):
    """Membuka file dengan aplikasi default. TIDAK BERUBAH."""
    try:
        if sys.platform.startswith('win'): os.startfile(filepath)
        elif sys.platform.startswith('darwin'): subprocess.call(["open", filepath])
        else: subprocess.call(["xdg-open", filepath])
    except Exception as e:
        messagebox.showerror("Gagal Membuka File", f"Tidak dapat membuka file:\n{filepath}\n\nError: {e}")

# Tambahkan fungsi ini di dalam file modules/shared_utils.py

def save_json_data(data, file_path):
    """Fungsi generik untuk menyimpan data ke file JSON."""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except IOError as e:
        messagebox.showerror("Gagal Menyimpan File", f"Gagal menyimpan data ke '{os.path.basename(file_path)}':\n{e}")
        return False