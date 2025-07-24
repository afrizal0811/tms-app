import requests
import json
import os
import sys
from tkinter import messagebox

def resource_path(relative_path):
    """Mendapatkan path absolut ke resource di bundle atau development."""
    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS
    else:
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

def get_constant_file_path(base_path):
    # 1. Cek di bundle (PyInstaller)
    bundle_path = resource_path("modules/constant.json")
    if os.path.exists(bundle_path):
        return bundle_path

    # 2. Cek di root project (development)
    dev_path = os.path.join(base_path, "constant.json")
    if os.path.exists(dev_path):
        return dev_path

    raise FileNotFoundError("File constant.json tidak ditemukan di bundle maupun root project.")

def get_lokasi_nama_by_kode(mapping, kode):
    for nama, kode_value in mapping.items():
        if kode_value == kode:
            return nama
    return None
    

def main(base_path):
    """
    Fungsi utama untuk menjalankan sinkronisasi berdasarkan lokasi yang dipilih.
    """
    constant_file = get_constant_file_path(base_path)
    master_file = os.path.join(base_path, 'master.json')
    config_file = os.path.join(base_path, 'config.json')
    api_url = 'https://apiweb.mile.app/api/v3/vehicles'

    # Muat semua konstanta
    constants = load_json_data(constant_file)
    if not constants:
        messagebox.showerror("File Tidak Ditemukan", "File constant gagal dimuat. \n \n Hubungi Admin.")
        return
    # Muat config untuk mendapatkan kode lokasi
    config = load_json_data(config_file)
    if not config:
        messagebox.showerror("File Tidak Ditemukan", "File config gagal dimuat. \n \n Hubungi Admin.")
        return
    lokasi_kode = config.get('lokasi')
    if not lokasi_kode:
        messagebox.showerror("Data Tidak Ditemukan", "Kode lokasi tidak ditemukan. \n \n Hubungi Admin.")
        return

    # Ambil token dan hubId dari konstanta
    api_token = constants.get('token')
    hub_ids_map = config.get('hub_ids', {})
    lokasi_mapping = constants.get('lokasi_mapping', {})
    hub_id = hub_ids_map.get(lokasi_kode)
    lokasi_nama = get_lokasi_nama_by_kode(lokasi_mapping, lokasi_kode)
    
    if not api_token:
        messagebox.showerror("Data Tidak Ditemukan", "Token tidak ditemukan. \n \n Hubungi Admin.")
        return
    if not hub_id:
        messagebox.showerror(
            "Data Tidak Ditemukan",
            f"Hub ID untuk lokasi '{lokasi_nama}' tidak ditemukan. \n \n Hubungi Admin."
        )
        return
    
    processed_api_list = fetch_and_process_vehicle_data(api_token, api_url, hub_id, lokasi_nama)
    
    master_vehicle_list = load_json_data(master_file)
    if master_vehicle_list:
        updated_master, was_updated = compare_and_update_master(processed_api_list, master_vehicle_list)
        if was_updated:
            save_data_to_json(updated_master, master_file)

def load_json_data(file_path):
    """Fungsi generik untuk membaca data dari file JSON."""
    if not os.path.exists(file_path):
        raise FileNotFoundError(f"File yang dibutuhkan tidak ditemukan di: {file_path}")
    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        raise IOError(f"Gagal membaca atau mem-parsing file '{os.path.basename(file_path)}': {e}")

def fetch_and_process_vehicle_data(token, api_url, hub_id, lokasi_nama):
    """Fungsi untuk memanggil API dengan hubId dinamis dan penanganan error yang lebih baik."""
    headers = {'Authorization': f'Bearer {token}'}
    params = {'limit': 100, 'hubId': hub_id}
    
    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        # Memeriksa jika status code bukan 2xx atau 3xx dan memberikan error
        response.raise_for_status() 

        api_response_data = response.json()
        vehicle_data = api_response_data.get('data')

        # PERUBAHAN: Memeriksa jika 'data' tidak ada, None, atau list kosong []

        if not vehicle_data:
            messagebox.showerror(
                "Data Tidak Ditemukan",
                f"Hub ID untuk lokasi '{lokasi_nama}' tidak ditemukan. \n \n Hubungi Admin."
            )

        processed_data = [{'Email': v.get('assignee'), 'Plat': v.get('name')} for v in vehicle_data if v.get('assignee') and v.get('name')]
        return processed_data

    except requests.exceptions.HTTPError as err:
        # PERUBAHAN: Memberikan pesan error yang lebih spesifik untuk status code non-200
        raise ValueError(f"Gagal mengambil data dari API (HTTP Status: {err.response.status_code}).\nPastikan Token atau Hub ID sudah benar.")
    except requests.exceptions.RequestException as e:
        # Menangkap error koneksi, timeout, dll.
        raise ConnectionError(f"Gagal terhubung ke API: {e}")


def compare_and_update_master(api_data, master_data):
    """Membandingkan data API dengan master dan mengembalikan data yang sudah diperbarui."""
    updated_master_data = list(master_data)
    master_map = {item['Email']: item for item in updated_master_data if 'Email' in item}
    was_updated = False
    for vehicle in api_data:
        email = vehicle.get('Email')
        api_plat = vehicle.get('Plat')
        if email in master_map and master_map[email].get('Plat') != api_plat:
            master_map[email]['Plat'] = api_plat
            was_updated = True
    return updated_master_data, was_updated

def save_data_to_json(data, file_path):
    """Fungsi untuk menyimpan data ke dalam file JSON."""
    try:
        with open(file_path, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)
    except IOError as e:
        raise IOError(f"Gagal menyimpan data ke file '{os.path.basename(file_path)}': {e}")

if __name__ == "__main__":
    pass
