import requests
import json
import os
import tkinter as tk
import traceback
from utils.function import (
    MASTER_JSON_PATH,
    TYPE_MAP_PATH,
    load_config,
    load_constants,
    load_master_data,
    load_secret,
    save_json_data,
    show_error_message,
    show_ask_message
)
from utils.messages import ERROR_MESSAGES, ASK_MESSAGES
from utils.api_handler import handle_requests_error

# =============================================================================
# LOAD / SAVE TYPE MAP
# =============================================================================
def load_type_map():
    # Jika belum ada file → buat kosong
    if not os.path.exists(TYPE_MAP_PATH):
        with open(TYPE_MAP_PATH, "w", encoding="utf-8") as f:
            json.dump({"type": {}}, f, indent=4)
        return {}
    try:
        with open(TYPE_MAP_PATH, "r", encoding="utf-8") as f:
            data = json.load(f)
            return data.get("type", {})
    except Exception:
        return {}

def save_type_map(type_map):
    with open(TYPE_MAP_PATH, "w", encoding="utf-8") as f:
        json.dump({"type": type_map}, f, indent=4)

# =============================================================================
# ASK USER UNTUK SUBSTITUSI TYPE (hanya saat type_map kosong)
# =============================================================================
def ask_type_substitution(plat, original_type, assignee_email, driver_users, constants, rest_config):
    prefix = original_type.split("-")[0] if "-" in original_type else original_type
    assignee_email = assignee_email.lower()

    # Cari driver_name dari driver_users
    driver_name = ""
    for user in driver_users:
        if user.get("email", "").lower() == assignee_email:
            driver_name = user.get("name", "")
            break

    if driver_name.startswith("'FRZ'"):
        prefix = "FROZEN"
    elif driver_name.startswith("'DRY'"):
        prefix = "DRY"

    type_options = [f"{prefix}-{suffix}" for suffix in constants.get("vehicle_types", [])]
    selected = tk.StringVar(value="")

    root = tk.Tk()
    root.title("Pilih Tipe Kendaraan")

    def on_close():
        if show_ask_message("Konfirmasi", ASK_MESSAGES["CONFIRM_CANCEL_SETUP"]):
            rest_config()
            root.quit()
        else:
            return 
        
    root.protocol("WM_DELETE_WINDOW", on_close)

    window_width = 300
    window_height = 500
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    x_position = (screen_width // 2) - (window_width // 2)
    y_position = (screen_height // 2) - (window_height // 2)
    root.geometry(f"{window_width}x{window_height}+{x_position}+{y_position}")
    root.resizable(False, False)

    label = tk.Label(
        root,
        text=f"Plat '{plat}' memiliki tipe tidak standar:\n'{original_type}'\n\nPilih tipe kendaraan:",
        wraplength=280,
        justify="left",
        font=("Segoe UI", 12)
    )
    label.pack(pady=(15, 10))

    button_frame = tk.Frame(root)
    button_frame.pack(pady=(5, 10))
    buttons = []

    def on_button_click(option, btn):
        selected.set(option)
        for b in buttons:
            b.config(relief="raised", bg="SystemButtonFace")
        btn.config(relief="sunken", bg="#cce5ff")

    for opt in type_options:
        btn = tk.Button(button_frame, text=opt, width=28, anchor="w", font=("Segoe UI", 12))
        btn.config(command=lambda opt=opt, btn=btn: on_button_click(opt, btn))
        btn.pack(pady=3)
        buttons.append(btn)

    action_frame = tk.Frame(root)
    action_frame.pack(pady=10)

    tk.Button(action_frame, text="OK", width=10, command=root.quit, font=("Segoe UI", 12)).pack(side="left", padx=10)

    root.mainloop()
    root.destroy()

    return selected.get() or None



# =============================================================================
# SYNC HUB IDS
# =============================================================================
def sync_hub(api_token, constants):
    base_url = constants.get("base_url")
    url = f"{base_url}/hubs"
    headers = {"Authorization": f"Bearer {api_token}"}
    try:
        resp = requests.get(url, headers=headers, timeout=30)
        resp.raise_for_status()
        data = resp.json().get("data", [])
        excluded_id = "683924970c29c079e30d862f"
        location_id = constants.get("location_id", {})
        result = {}
        for hub in data:
            if hub.get("_id") == excluded_id:
                continue
            hub_name = hub.get("name", "")
            for nama, kode in location_id.items():
                if nama in hub_name:
                    result[kode] = hub.get("_id")
        return dict(sorted(result.items()))
    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
        return None
    except Exception as e:
        show_error_message("Error Tak Terduga", ERROR_MESSAGES["UNKNOWN_ERROR"].format(
        error_detail=f"{e}\n\n{traceback.format_exc()}"
    ))

# =============================================================================
# FETCH VEHICLES (gunakan type_map + substitusi manual pertama kali)
# =============================================================================
def fetch_and_process_vehicle_data(api_token, hub_id, constants, type_map, driver_users, rest_config):
    base_url = constants.get('base_url')
    api_url = f'{base_url}/vehicles'
    headers = {'Authorization': f'Bearer {api_token}'}
    params = {'limit': 100, 'hubId': hub_id}
    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        vehicle_data = response.json().get('data', [])
        valid_types = constants.get("vehicle_types", [])
        type_map_updated = False

        vehicles = []
        for v in vehicle_data:
            if not v.get('assignee') or not v.get('name') or not v.get('tags'):
                continue

            raw_type = v.get('tags')[0]
            mapped_type = type_map.get(raw_type)

            if not mapped_type:
                if any(t in raw_type for t in valid_types):
                    mapped_type = raw_type
                else:
                    substitute = ask_type_substitution(
                        plat=v.get('name', ''),
                        original_type=raw_type,
                        assignee_email=v.get('assignee', ''),
                        driver_users=driver_users,
                        constants=constants,
                        rest_config=rest_config
                    )
                    if substitute:
                        type_map[raw_type] = substitute
                        mapped_type = substitute
                        type_map_updated = True
                    else:
                        mapped_type = raw_type

            vehicles.append({
                'Email': v.get('assignee', '').lower(),
                'Plat': v.get('name', ''),
                'Type': mapped_type
            })

        if type_map_updated:
            save_type_map(type_map)

        return vehicles

    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
        return None
    except Exception as e:
        show_error_message(
            "Error Kendaraan",
            ERROR_MESSAGES["UNKNOWN_ERROR"].format(
                error_detail=f"{e}\n\n{traceback.format_exc()}"
            )
        )
        return None

# =============================================================================
# FETCH USERS
# =============================================================================
def fetch_driver_users(api_token, hub_id, constants, lokasi_kode):
    base_url = constants.get("base_url")
    role_ids = constants.get("role_ids", {})
    if lokasi_kode in ("plck", "pldm"):
        driver_id = role_ids.get("driverJkt")
    else:
        driver_id = role_ids.get("driver")
    api_url = f"{base_url}/users"
    headers = {"Authorization": f"Bearer {api_token}"}
    params = {
        "roleId": driver_id,
        "hubId": hub_id,
        "status": "active",
        "limit": 500
    }
    try:
        resp = requests.get(api_url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        users = resp.json().get("data", [])
        return [u for u in users if u.get("name") and ("FRZ" in u["name"] or "DRY" in u["name"])]
    
    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
        return None # Mengembalikan None saat terjadi error requests
    except Exception as e:
        show_error_message("Error Users API", ERROR_MESSAGES["UNKNOWN_ERROR"].format(
            error_detail=f"{e}\n\n{traceback.format_exc()}"
        ))
        return None

# =============================================================================
# UPDATE DRIVER
# =============================================================================
def update_driver_master(master_driver, users, vehicles):
    updated_driver = [dict(item) for item in master_driver] if master_driver else []
    master_map = {d['Email'].lower(): d for d in updated_driver if 'Email' in d}
    vehicle_map = {v['Email']: (v['Plat'], v.get('Type', '')) for v in vehicles}

    was_updated = False
    users_with_vehicles = {v['Email'] for v in vehicles}

    for user in users:
        email = user.get("email", "").lower()
        name = user.get("name", "")
        if email not in users_with_vehicles:
            continue
        
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
# MAIN
# =============================================================================
def main(rest_config):
    try:
        constants = load_constants()
        config = load_config()
        secrets = load_secret()
        type_map = load_type_map()

        if not constants or not config or not secrets:
            show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
            return

        api_token = secrets.get("token")
        if not api_token:
            show_error_message("Error Token API", ERROR_MESSAGES["API_TOKEN_MISSING"])
            return

        # -----------------------------------------
        # Load master_data dengan aman
        # -----------------------------------------
        master_data = load_master_data()
        master_df = None
        hub_ids = {}

        if master_data:
            # master_data diharapkan dict dengan kunci "df" dan "hub_ids"
            master_df = master_data.get("df")
            hub_ids = master_data.get("hub_ids", {})

        lokasi_kode = config.get("lokasi")
        if not lokasi_kode:
            show_error_message("Gagal", ERROR_MESSAGES["LOCATION_CODE_MISSING"])
            return

        # Selalu sync hub dari API
        new_hub_ids = sync_hub(api_token, constants)
        if new_hub_ids is None:
            return

        if hub_ids != new_hub_ids:
            hub_ids.update(new_hub_ids)

        # Sort berdasarkan key ascending
        hub_ids = dict(sorted(hub_ids.items()))
        
        hub_id = hub_ids.get(lokasi_kode)
        if not hub_id:
            show_error_message("Gagal", ERROR_MESSAGES["HUB_ID_MISSING"])
            return

        # Ambil users & vehicles (boleh saja master_df None — fungsi aman terhadap itu)
        users = fetch_driver_users(api_token, hub_id, constants, lokasi_kode)
        vehicles = fetch_and_process_vehicle_data(api_token, hub_id, constants, type_map, users, rest_config)

        # Siapkan master records aman untuk update_driver_master
        if master_df is not None and hasattr(master_df, "to_dict"):
            master_records = master_df.to_dict("records")
        else:
            master_records = []

        updated_driver, driver_updated = update_driver_master(master_records, users, vehicles)

        # Simpan jika ada perubahan
        if driver_updated or hub_ids:
            save_json_data({"driver": updated_driver, "hub_ids": hub_ids}, MASTER_JSON_PATH)

    except Exception as e:
        show_error_message("Error Tidak Dikenal", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=e))


if __name__ == "__main__":
    main()