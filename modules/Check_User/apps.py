from tkinter import ttk
import requests
import tkinter as tk
from utils.function import (
    CONFIG_PATH,
    load_config,
    load_constants,
    load_secret,
    save_json_data,
    show_ask_message,
    show_error_message
)
from utils.messages import ERROR_MESSAGES, ASK_MESSAGES

# Role ID yang harus dikecualikan
EXCLUDED_ROLE_ID = "6703410af6be892f3208ecde"

def main(parent_window):
    try:
        constants = load_constants()
        config = load_config()
        secrets = load_secret()

        if not constants:
            show_error_message("Gagal", ERROR_MESSAGES["CONSTANT_FILE_ERROR"])
            return
        if not config:
            # Menggunakan pesan yang relevan meskipun tidak ada di ERROR_MESSAGES asli
            show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
            return
        if not secrets:
            show_error_message("Gagal", ERROR_MESSAGES["SECRET_FILE_ERROR"])
            return

        lokasi_kode = config.get('lokasi')
        api_token = secrets.get('token')
        hub_ids_map = constants.get('hub_ids', {})
        hub_id = hub_ids_map.get(lokasi_kode)

        if not lokasi_kode:
            show_error_message("Gagal", ERROR_MESSAGES["LOCATION_CODE_MISSING"])
            return
        if not api_token:
            show_error_message("Gagal", ERROR_MESSAGES["API_TOKEN_MISSING"])
            return
        if not hub_id:
            show_error_message("Gagal", ERROR_MESSAGES["HUB_ID_MISSING"])
            return

        restricted_roles = list(constants.get('restricted_role_ids', {}).values())
        api_url = "https://apiweb.mile.app/api/v3/users"
        headers = {'Authorization': f'Bearer {api_token}'}
        params = {'limit': 500, 'hubId': hub_id, 'status': 'active'}

        try:
            response = requests.get(api_url, headers=headers, params=params, timeout=30)
            response.raise_for_status()
            users_data = response.json().get('data', [])
        except requests.exceptions.HTTPError as errh:
            status_code = errh.response.status_code
            if status_code == 401:
                show_error_message("Akses Ditolak (401)", ERROR_MESSAGES["API_TOKEN_MISSING"])
            elif status_code >= 500:
                show_error_message("Masalah Server API", ERROR_MESSAGES["SERVER_ERROR"].format(error_detail=status_code))
            else:
                show_error_message("Kesalahan HTTP", ERROR_MESSAGES["HTTP_ERROR_GENERIC"].format(status_code=status_code))
            return
        except requests.exceptions.ConnectionError:
            show_error_message("Koneksi Gagal", ERROR_MESSAGES["CONNECTION_ERROR"].format(error_detail="Tidak dapat terhubung ke server. Periksa koneksi internet Anda."))
            return
        except requests.exceptions.RequestException as e:
            show_error_message("Kesalahan API", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=e))
            return

        # filter default menggunakan restricted_roles dan mengecualikan EXCLUDED_ROLE_ID
        def filter_users():
            return [u for u in users_data if u.get('roleId') in restricted_roles and u.get('roleId') != EXCLUDED_ROLE_ID]

        filtered_users = filter_users()
        filtered_users.sort(key=lambda u: u.get('name', '').lower())
        if not filtered_users:
            show_error_message("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
            return

        dialog = tk.Toplevel(parent_window)
        dialog.title("Pilih Akun Pengguna")
        dialog_width, dialog_height = 500, 500
        x = (dialog.winfo_screenwidth() // 2) - (dialog_width // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog_height // 2)
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        tk.Label(dialog, text="Pilih Akun Anda", font=("Arial", 16, "bold")).pack(pady=5)

        container = tk.Frame(dialog, relief="solid", bd=1)
        container.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        tk.Label(container, text="Pilih salah satu nama di bawah ini:", font=("Arial", 12, "italic"), anchor="w").pack(anchor="w", pady=5, padx=5)

        selected_var = tk.StringVar()
        canvas = tk.Canvas(container, borderwidth=0)
        scrollbar = ttk.Scrollbar(container, orient="vertical", command=canvas.yview)
        scroll_frame = tk.Frame(canvas)
        scroll_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.create_window((0, 0), window=scroll_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        save_button = tk.Button(dialog, text="Simpan Pilihan", state=tk.DISABLED, font=("Arial", 14, "bold"), bg="#4CAF50", fg="white")
        save_button.pack(fill=tk.X, padx=10, pady=10)

        def on_selection(*args):
            if selected_var.get():
                save_button.config(state=tk.NORMAL)
            else:
                save_button.config(state=tk.DISABLED)

        selected_var.trace_add('write', on_selection)

        radio_buttons = []
        def populate_user_list(user_list):
            for rb in radio_buttons:
                rb.destroy()
            radio_buttons.clear()
            for user in user_list:
                name_cap = user.get('name', 'Nama tidak ditemukan').title()
                rb = tk.Radiobutton(scroll_frame, text=name_cap, variable=selected_var,
                                     value=user.get('_id'), anchor='w', font=("Arial", 12))
                rb.pack(fill=tk.X, padx=10, pady=2)
                radio_buttons.append(rb)

        populate_user_list(filtered_users)

        # Rahasia: tekan CTRL+A untuk menghapus filter role (tetap mengecualikan EXCLUDED_ROLE_ID)
        def secret_show_all(event=None):
            all_users = [u for u in users_data if u.get('roleId') != EXCLUDED_ROLE_ID]
            all_users.sort(key=lambda u: u.get('name', '').lower())
            populate_user_list(all_users)
            return "break"

        dialog.bind_all("<Control-a>", secret_show_all)

        def simpan_pengguna():
            user_id = selected_var.get()
            selected_user = next((u for u in users_data if u.get('_id') == user_id), None)
            if not selected_user:
                show_error_message("Error", ERROR_MESSAGES["USER_SELECTION_NOT_FOUND"])
                return
            
            # Konfirmasi sekarang ada di dalam fungsi simpan_pengguna
            if not show_ask_message("Konfirmasi", ASK_MESSAGES["CONFIRM_SAVE_USER"]):
                return
                
            config['user_checked'] = {
                'name': selected_user.get('name'),
                '_id': selected_user.get('_id'),
                'hub_id': selected_user.get('hubId'),
                'role_id': selected_user.get('roleId')
            }
            save_json_data(config, CONFIG_PATH)
            dialog.destroy()

        save_button.config(command=simpan_pengguna)

        dialog.transient(parent_window)
        dialog.grab_set()
        parent_window.wait_window(dialog)

    except Exception as e:
        show_error_message("Error Tak Terduga", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=e))