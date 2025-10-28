import os
import sys
import json
import subprocess
import pandas as pd
from tkinter import filedialog, messagebox
import tkinter as tk
from utils.messages import ERROR_MESSAGES
# =============================================================================
# PENGELOLAAN PATH TERPUSAT (PYINSTALLER-COMPATIBLE)
# =============================================================================


def show_error_message(title, message):
    messagebox.showerror(title, message)

def show_info_message(title, message):
    messagebox.showinfo(title, message)

def show_ask_message(title, message):
    return messagebox.askyesno(title, message)

def get_base_path():
    """
    Mendapatkan path dasar aplikasi, baik saat dijalankan sebagai skrip
    maupun sebagai file executable hasil PyInstaller. Path ini menunjuk ke
    direktori tempat .exe berada atau root folder proyek.
    """
    if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
        # Dijalankan sebagai bundel PyInstaller. Path dasar adalah direktori
        # yang berisi file executable.
        return os.path.dirname(sys.executable)
    else:
        # Dijalankan sebagai skrip normal. Path dasar adalah root proyek.
        # Asumsi file ini ada di dalam folder 'modules'.
        # __file__ -> .../modules/shared_utils.py
        # os.path.dirname -> .../modules
        # os.path.dirname -> .../ (project root)
        return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))

def resource_path(relative_path):
    """
    Mendapatkan path absolut ke resource yang di-bundle (seperti icon, constant.json).
    Ini dicari di dalam bundle _MEIPASS saat frozen.
    """
    try:
        # PyInstaller membuat folder sementara dan menyimpan path di sys._MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        # _MEIPASS tidak ada, berarti dijalankan di lingkungan Python normal.
        # Gunakan path dasar dari root proyek.
        base_path = get_base_path()

    return os.path.join(base_path, relative_path)

# --- Definisi Path ---
SECRET_PATH = resource_path('secret.json')
BASE_DIR = get_base_path()
CONFIG_PATH = os.path.join(BASE_DIR, 'config.json')
MASTER_JSON_PATH = os.path.join(BASE_DIR, 'master.json')
TYPE_MAP_PATH = os.path.join(BASE_DIR, 'type_map.json')
CONSTANT_PATH = resource_path('constant.json')


# =============================================================================
# FUNGSI UTILITAS UMUM
# =============================================================================

def load_json_data(file_path, default_data=None):
    """Fungsi generik untuk membaca data dari file JSON.
       - Jika file tidak ada:
         * Jika default_data diberikan → buat file dengan default_data
         * Jika default_data None → return None
       - Jika file ada tapi rusak → buat ulang jika default_data ada, selain itu return None
    """
    if not os.path.exists(file_path):
        if default_data is not None:
            os.makedirs(os.path.dirname(file_path), exist_ok=True)
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(default_data, f, indent=4)
            return default_data
        else:
            return None

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError):
        if default_data is not None:
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(default_data, f, indent=4)
            return default_data
        else:
            return None

def save_json_data(data, file_path):
    """Fungsi generik untuk menyimpan data ke file JSON."""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
        return True
    except IOError as e:
        messagebox.showerror("Gagal Menyimpan File", f"Gagal menyimpan data ke '{os.path.basename(file_path)}':\n{e}")
        return False

def load_config():
    """Memuat konfigurasi dari config.json."""
    return load_json_data(CONFIG_PATH)

def load_constants():
    """Memuat konstanta dari constant.json."""
    return load_json_data(CONSTANT_PATH)

def load_type_map():
    return load_json_data(TYPE_MAP_PATH)

def ensure_config_exists():
    """
    Memastikan config.json ada di direktori aplikasi.
    Jika tidak ada, buat file default.
    """
    if not os.path.exists(CONFIG_PATH):
        default_config = {"lokasi": "", "user_checked": None}
        save_json_data(default_config, CONFIG_PATH)

def load_master_data(lokasi_cabang=None):
    """Memuat dan memproses master.json."""
    data = load_json_data(MASTER_JSON_PATH)
    if data is None:
        return None
    
    try:
        if 'driver' not in data:
            show_error_message("Gagal Memuat Master Data", ERROR_MESSAGES["MASTER_DATA_MISSING"])
            return None

        df = pd.DataFrame(data['driver'])  # gunakan key "driver"
        if not df.empty: 
            df.columns = [col.strip() for col in df.columns]
            if 'Email' not in df.columns or 'Driver' not in df.columns:
                show_error_message("Gagal Memuat Master Data", ERROR_MESSAGES["MASTER_DATA_MISSING"])
                return None

            df['Email'] = df['Email'].astype(str).str.strip().str.lower()
            df['Driver'] = df['Driver'].astype(str).str.strip()
            if lokasi_cabang:
                df = df[df['Email'].str.contains(lokasi_cabang, case=False, na=False)].copy()
        
        return {
            "df": df,
            "hub_ids": data.get("hub_ids", {})
        }
    except Exception:
        show_error_message("Gagal Memuat Master Data", ERROR_MESSAGES["MASTER_FILE_ERROR"])
        return None

def get_save_path(base_name="Laporan", extension=".xlsx"):
    """Membuka dialog untuk memilih folder penyimpanan."""
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
    """Membuka file dengan aplikasi default sistem operasi."""
    try:
        if sys.platform.startswith('win'):
            os.startfile(filepath)
        elif sys.platform.startswith('darwin'):
            subprocess.call(["open", filepath])
        else:
            subprocess.call(["xdg-open", filepath])
    except Exception as e:
        show_error_message("Gagal Membuka File", ERROR_MESSAGES["FAILED_OPENING_FILE"].format(error_detail=str(e)))

def load_secret():
    """
    Memuat data rahasia dari file secret.json.
    Jika file tidak ada, akan dibuatkan dengan placeholder.
    """
    if not os.path.exists(SECRET_PATH):
        sample_secret = {"token": "PASTE_YOUR_MILEAPP_TOKEN_HERE"}
        with open(SECRET_PATH, 'w') as f:
            json.dump(sample_secret, f, indent=2)
        show_error_message("Gagal", ERROR_MESSAGES["API_TOKEN_MISSING"])
        return None
        
    try:
        with open(SECRET_PATH, 'r') as f:
            return json.load(f)
    except (FileNotFoundError, json.JSONDecodeError):
        show_error_message("Gagal", ERROR_MESSAGES["SECRET_FILE_ERROR"])
        return None