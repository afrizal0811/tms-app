# modules/Sync_Driver/apps.py (KODE BARU)

import requests

# 1. Impor semua fungsi dan path yang dibutuhkan dari shared_utils
from utils.function import (
    MASTER_JSON_PATH,
    load_config, 
    load_constants, 
    load_json_data, 
    load_secret,
    save_json_data,
    show_error_message
)
from utils.messages import ERROR_MESSAGES

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
            show_error_message("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
            return []

        processed_data = [
            {'Email': v.get('assignee'), 'Plat': v.get('name')} 
            for v in vehicle_data if v.get('assignee') and v.get('name')
        ]
        return processed_data

    except requests.exceptions.HTTPError as err:
        show_error_message("Kesalahan HTTP", ERROR_MESSAGES["HTTP_ERROR_GENERIC"].format(status_code=err.response.status_code))
    except requests.exceptions.RequestException as e:
        show_error_message("Kesalahan API", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=e))


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
        secrets = load_secret()

        if not constants or not config or not secrets:
            show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
            return

        lokasi_kode = config.get('lokasi')
        if not lokasi_kode:
            show_error_message("Gagal", ERROR_MESSAGES["LOCATION_CODE_MISSING"])
            return

        # Ambil token dan hubId
        api_token = secrets.get('token')
        hub_ids_map = constants.get('hub_ids', {})
        lokasi_mapping = constants.get('lokasi_mapping', {})
        
        hub_id = hub_ids_map.get(lokasi_kode)
        lokasi_nama = get_lokasi_nama_by_kode(lokasi_mapping, lokasi_kode)
        
        if not api_token:
            show_error_message("Error Token API", ERROR_MESSAGES["API_TOKEN_MISSING"])
            return
        
        if not hub_id:
            show_error_message("Gagal", ERROR_MESSAGES["HUB_ID_MISSING"].format(lokasi_code="ini"))
            return
        
        # 3. Proses sinkronisasi
        processed_api_list = fetch_and_process_vehicle_data(api_token, hub_id, lokasi_nama)
        
        master_vehicle_list = load_json_data(MASTER_JSON_PATH)
        if master_vehicle_list is None:
             show_error_message("Gagal", ERROR_MESSAGES["MASTER_DATA_MISSING"])
             return

        updated_master, was_updated = compare_and_update_master(processed_api_list, master_vehicle_list)
        
        if was_updated:
            # 4. Simpan data menggunakan shared_utils
            save_json_data(updated_master, MASTER_JSON_PATH)

    except (ValueError, ConnectionError) as e:
        # Menangkap error spesifik dari proses fetch
        show_error_message("Error Sinkronisasi", str(e))
    except Exception as e:
        # Menangkap error lainnya
        show_error_message("Error Tidak Dikenal", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=e))

if __name__ == "__main__":
    # Untuk pengujian langsung (jika diperlukan)
    main()