# modules/Sync_Data/apps.py

import requests
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

# =============================================================================
# SYNC HUB IDS (TETAP UPDATE SEMUA HUB)
# =============================================================================
def sync_hub(api_token, constants):
    base_url = constants.get("base_url")
    url = f"{base_url}/hubs"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json().get("data", [])
    except Exception as e:
        show_error_message("Error Hub API", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=e))
        return {}

    excluded_id = "683924970c29c079e30d862f"
    lokasi_mapping = constants.get("lokasi_mapping", {})
    result = {}
    for hub in data:
        if hub.get("_id") == excluded_id:
            continue
        hub_name = hub.get("name", "")
        for nama, kode in lokasi_mapping.items():
            if nama in hub_name:
                result[kode] = hub.get("_id")
    return dict(sorted(result.items()))

# =============================================================================
# FETCH VEHICLES (gunakan type_map dari master.json)
# =============================================================================
def fetch_and_process_vehicle_data(api_token, hub_id, constants, type_map):
    base_url = constants.get('base_url')
    api_url = f'{base_url}/vehicles'
    headers = {'Authorization': f'Bearer {api_token}'}
    params = {'limit': 100, 'hubId': hub_id}
    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        vehicle_data = response.json().get('data', [])
    except Exception as e:
        show_error_message("Error Kendaraan", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=e))
        return []

    vehicles = []
    for v in vehicle_data:
        if not v.get('assignee') or not v.get('name'):
            continue
        raw_type = (v.get('tags', [])[0] if v.get('tags') else '')
        mapped_type = type_map.get(raw_type, raw_type)
        vehicles.append({
            'Email': v.get('assignee', '').lower(),
            'Plat': v.get('name', ''),
            'Type': mapped_type
        })
    return vehicles

# =============================================================================
# FETCH USERS (DRIVER)
# =============================================================================
def fetch_driver_users(api_token, hub_id, constants):
    base_url = constants.get("base_url")
    api_url = f"{base_url}/users"
    headers = {"Authorization": f"Bearer {api_token}"}
    params = {
        "roleId": "6703410af6be892f3208ecde",
        "hubId": hub_id,
        "status": "active",
        "limit": 500
    }
    try:
        resp = requests.get(api_url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        users = resp.json().get("data", [])
        return [u for u in users if u.get("name") and ("FRZ" in u["name"] or "DRY" in u["name"])]
    except Exception as e:
        show_error_message("Error Users API", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=e))
        return []

# =============================================================================
# UPDATE DRIVER DI MASTER
# =============================================================================
def update_driver_master(master_driver, users, vehicles):
    updated_driver = [dict(item) for item in master_driver] if master_driver else []
    master_map = {d['Email'].lower(): d for d in updated_driver if 'Email' in d}
    vehicle_map = {v['Email']: (v['Plat'], v.get('Type', '')) for v in vehicles}

    was_updated = False
    for user in users:
        email = user.get("email", "").lower()
        name = user.get("name", "")
        plat, vtype = vehicle_map.get(email, ("", ""))

        if email in master_map:
            if master_map[email].get("Driver") != name:
                master_map[email]["Driver"] = name
                was_updated = True
            if plat and master_map[email].get("Plat") != plat:
                master_map[email]["Plat"] = plat
                was_updated = True
            if vtype and master_map[email].get("Type") != vtype:
                master_map[email]["Type"] = vtype
                was_updated = True
        else:
            updated_driver.append({
                "Email": email,
                "Driver": name,
                "Plat": plat,
                "Type": vtype
            })
            was_updated = True

    updated_driver = sorted(updated_driver, key=lambda x: x.get("Email", "").lower())
    return updated_driver, was_updated

# =============================================================================
# MAIN SYNC ALL
# =============================================================================
def main():
    try:
        constants = load_constants()
        config = load_config()
        secrets = load_secret()

        if not constants or not config or not secrets:
            show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
            return

        api_token = secrets.get("token")
        if not api_token:
            show_error_message("Error Token API", ERROR_MESSAGES["API_TOKEN_MISSING"])
            return

        master_json = load_json_data(MASTER_JSON_PATH) or {}
        if "hub_ids" not in master_json:
            master_json["hub_ids"] = {}
        if "driver" not in master_json:
            master_json["driver"] = []
        if "type_map" not in master_json:
            master_json["type_map"] = {}

        type_map = master_json.get("type_map", {})

        # 1. Sync semua hub_ids tanpa filter lokasi aktif
        hub_ids = sync_hub(api_token, constants)
        if hub_ids:
            master_json["hub_ids"] = hub_ids

        lokasi_kode = config.get("lokasi")
        if not lokasi_kode:
            show_error_message("Gagal", ERROR_MESSAGES["LOCATION_CODE_MISSING"])
            return
        hub_id = hub_ids.get(lokasi_kode)
        if not hub_id:
            show_error_message("Gagal", ERROR_MESSAGES["HUB_ID_MISSING"].format(lokasi_code=lokasi_kode))
            return

        # 2. Ambil users dan vehicles hub aktif (pakai type_map dari master.json)
        users = fetch_driver_users(api_token, hub_id, constants)
        vehicles = fetch_and_process_vehicle_data(api_token, hub_id, constants, type_map)

        # 3. Update driver
        updated_driver, driver_updated = update_driver_master(master_json["driver"], users, vehicles)
        if driver_updated or hub_ids:
            master_json["driver"] = updated_driver
            save_json_data(master_json, MASTER_JSON_PATH)

    except Exception as e:
        show_error_message("Error Tidak Dikenal", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=e))

if __name__ == "__main__":
    main()
