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

def fetch_and_process_vehicle_data(token, hub_id, constants):
    """Memanggil API kendaraan dan memproses hasilnya."""
    base_url = constants.get('base_url')
    api_url = f'{base_url}/vehicles'
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
    try:
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

        api_token = secrets.get('token')
        if not api_token:
            show_error_message("Error Token API", ERROR_MESSAGES["API_TOKEN_MISSING"])
            return

        # Load master.json
        master_json = load_json_data(MASTER_JSON_PATH)
        if not master_json or "driver" not in master_json or "hub_ids" not in master_json:
            show_error_message("Gagal", ERROR_MESSAGES["MASTER_DATA_MISSING"])
            return

        hub_ids_map = master_json.get("hub_ids", {})
        hub_id = hub_ids_map.get(lokasi_kode)
        lokasi_mapping = constants.get('lokasi_mapping', {})
        lokasi_nama = get_lokasi_nama_by_kode(lokasi_mapping, lokasi_kode)

        if not hub_id:
            show_error_message("Gagal", ERROR_MESSAGES["HUB_ID_MISSING"].format(lokasi_code=lokasi_kode))
            return

        # Proses sinkronisasi
        processed_api_list = fetch_and_process_vehicle_data(api_token, hub_id, constants)
        master_vehicle_list = master_json["driver"]

        updated_master, was_updated = compare_and_update_master(processed_api_list, master_vehicle_list)
        if was_updated:
            master_json["driver"] = updated_master
            save_json_data(master_json, MASTER_JSON_PATH)

    except (ValueError, ConnectionError) as e:
        show_error_message("Error Sinkronisasi", str(e))
    except Exception as e:
        show_error_message("Error Tidak Dikenal", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=e))

if __name__ == "__main__":
    # Untuk pengujian langsung (jika diperlukan)
    main()