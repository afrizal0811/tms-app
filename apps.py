# modules/Sync_Driver/apps.py (KODE BARU)

import requests
import json
from tkinter import messagebox

# 1. Impor semua fungsi dan path yang dibutuhkan dari shared_utils
from ..shared_utils import (
    load_config, 
    load_constants, 
    load_json_data, 
    save_json_data,
    MASTER_JSON_PATH
)

# ==============================================================================
# FUNGSI-FUNGSI BANTU (HELPER FUNCTIONS)
# ==============================================================================

# Fungsi duplikat (resource_path, get_constant_file_path, load_json_data, save_data_to_json)
# TELAH DIHAPUS DARI SINI.

# Fungsi spesifik untuk modul ini bisa tetap di sini.
def get_lokasi_nama_by_kode(mapping, kode):
    """Mencari nama lokasi (misal: "Sidoarjo") dari kode (misal: "plsda")."""
    for nama, kode_value in mapping.items():
        if kode_value == kode:
            return nama
    return None

def fetch_and_process_vehicle_data(token, hub_id, lokasi_nama):
    """Memanggil API kendaraan dan memproses hasilnya."""
    api_url = 'https://apiweb.mile.app/api/v3/vehicles'
    headers = {'Authorization': f'Bearer {token}'}
    params = {'limit': 100, 'hubId': hub_id}
    
    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status() 

        api_response_data = response.json()
        vehicle_data = api_response_data.get('data')

        if not vehicle_data:
            messagebox.showwarning(
                "Data Tidak Ditemukan",
                f"Tidak ada data kendaraan yang ditemukan dari API untuk lokasi '{lokasi_nama}'.\n\nHubungi Admin."
            )
            return []

        processed_data = [
            {'Email': v.get('assignee'), 'Plat': v.get('name')} 
            for v in vehicle_data if v.get('assignee') and v.get('name')
        ]
        return processed_data

    except requests.exceptions.HTTPError as err:
        raise ValueError(f"Gagal mengambil data dari API (HTTP Status: {err.response.status_code}).\nPastikan Token atau Hub ID sudah benar.")
    except requests.exceptions.RequestException as e:
        raise ConnectionError(f"Gagal terhubung ke API: {e}")


def compare_and_update_master(api_data, master_data):
    """Membandingkan data API dengan master dan mengembalikan data yang sudah diperbarui."""
    # Buat salinan untuk dimodifikasi
    updated_master_data = [dict(item) for item in master_data]
    master_map = {item['Email']: item for item in updated_master_data if 'Email' in item}
    was_updated = False

    for vehicle in api_data:
        email = vehicle.get('Email')
        api_plat = vehicle.get('Plat')
        if email in master_map and master_map[email].get('Plat') != api_plat:
            master_map[email]['Plat'] = api_plat
            was_updated = True
            
    return updated_master_data, was_updated


# ==============================================================================
# FUNGSI UTAMA (MAIN CONTROLLER)
# ==============================================================================

def main():
    """
    Fungsi utama untuk menjalankan sinkronisasi berdasarkan lokasi yang dipilih.
    Kini menggunakan shared_utils.
    """
    try:
        # 2. Muat semua file konfigurasi menggunakan shared_utils
        constants = load_constants()
        config = load_config()

        if not constants or not config:
            messagebox.showerror("Gagal", "File 'constant.json' atau 'config.json' gagal dimuat. Proses dibatalkan.")
            return

        lokasi_kode = config.get('lokasi')
        if not lokasi_kode:
            messagebox.showerror("Gagal", "Kode lokasi tidak ditemukan. Silakan atur lokasi terlebih dahulu.")
            return

        # Ambil token dan hubId dari konstanta
        api_token = constants.get('token')
        hub_ids_map = constants.get('hub_ids', {})
        lokasi_mapping = constants.get('lokasi_mapping', {})
        
        hub_id = hub_ids_map.get(lokasi_kode)
        lokasi_nama = get_lokasi_nama_by_kode(lokasi_mapping, lokasi_kode)
        
        if not api_token:
            messagebox.showerror("Gagal", f"Token API tidak ditemukan.\n\nHubungi Admin.")
            return
        
        if not hub_id:
            messagebox.showerror("Gagal", f"Hub ID tidak ditemukan.\n\nHubungi Admin.")
            return
        
        # 3. Proses sinkronisasi
        processed_api_list = fetch_and_process_vehicle_data(api_token, hub_id, lokasi_nama)
        
        master_vehicle_list = load_json_data(MASTER_JSON_PATH)
        if master_vehicle_list is None:
             messagebox.showerror("Gagal", "File master data tidak ditemukan.")
             return

        updated_master, was_updated = compare_and_update_master(processed_api_list, master_vehicle_list)
        
        if was_updated:
            # 4. Simpan data menggunakan shared_utils
            save_json_data(updated_master, MASTER_JSON_PATH)

    except (ValueError, ConnectionError) as e:
        # Menangkap error spesifik dari proses fetch
        messagebox.showerror("Error Sinkronisasi", str(e))
    except Exception as e:
        # Menangkap error lainnya
        messagebox.showerror("Error Tidak Dikenal", f"Terjadi kesalahan tak terduga saat sinkronisasi:\n{e}")

if __name__ == "__main__":
    # Untuk pengujian langsung (jika diperlukan)
    main()