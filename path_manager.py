# path_manager.py
import os
import sys

def get_base_path():
    """
    Mendapatkan path dasar yang benar, baik saat dijalankan sebagai skrip
    maupun sebagai file .exe yang dibuat oleh PyInstaller.
    """
    # sys.frozen akan bernilai True jika program berjalan sebagai .exe
    if getattr(sys, 'frozen', False):
        # Jika .exe, base_path adalah folder sementara (_MEIxxxx)
        # yang dibuat oleh PyInstaller. sys._MEIPASS menunjuk ke sana.
        return sys._MEIPASS
    else:
        # Jika skrip, base_path adalah direktori tempat file skrip ini berada.
        return os.path.dirname(os.path.abspath(__file__))

# Definisikan path ke master.json secara terpusat
BASE_DIR = get_base_path()
MASTER_JSON_PATH = os.path.join(BASE_DIR, 'modules', 'master.json')