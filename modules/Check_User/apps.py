from tkinter import ttk
import requests
import tkinter as tk
from utils.function import (
    CONFIG_PATH,
    load_config,
    load_constants,
    load_secret,
    load_master_data,
    save_json_data,
    show_ask_message,
    show_error_message
)
from utils.messages import ERROR_MESSAGES, ASK_MESSAGES
from utils.api_handler import handle_requests_error

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
            show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
            return
        if not secrets:
            show_error_message("Gagal", ERROR_MESSAGES["SECRET_FILE_ERROR"])
            return

        lokasi_kode = config.get('lokasi')
        api_token = secrets.get('token')
        master_data = load_master_data()
        if not master_data or 'hub_ids' not in master_data:
            show_error_message("Gagal", ERROR_MESSAGES["MASTER_DATA_MISSING"])
            return
        
        hub_ids_map = master_data['hub_ids']
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
        base_url = constants.get('base_url')
        api_url = f"{base_url}/users"
        headers = {'Authorization': f'Bearer {api_token}'}
        params = {'limit': 500, 'hubId': hub_id, 'status': 'active'}

        try:
            response = requests.get(api_url, headers=headers, params=params, timeout=30)
            response.raise_for_status()
            users_data = response.json().get('data', [])
        except requests.exceptions.RequestException as e:
            handle_requests_error(e)
        except Exception as e:
            import traceback
            show_error_message("Error Tak Terduga", ERROR_MESSAGES["UNKNOWN_ERROR"].format(
                error_detail=f"Terjadi kesalahan: {e}\n\n{traceback.format_exc()}"
            ))

        # Filter default
        def filter_users():
            return [u for u in users_data if u.get('roleId') in restricted_roles and u.get('roleId') != EXCLUDED_ROLE_ID]

        filtered_users = filter_users()
        filtered_users.sort(key=lambda u: u.get('name', '').lower())
        if not filtered_users:
            show_error_message("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
            return

        # === Dialog Pilihan User ===
        dialog = tk.Toplevel(parent_window)
        dialog.title("Pilih Akun Pengguna")
        dialog_width, dialog_height = 500, 500
        x = (dialog.winfo_screenwidth() // 2) - (dialog_width // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog_height // 2)
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        dialog.rowconfigure(0, weight=0)  # Title
        dialog.rowconfigure(1, weight=1)  # area tombol user
        dialog.rowconfigure(2, weight=0)  # pagination
        dialog.rowconfigure(3, weight=0)  # save button
        dialog.columnconfigure(0, weight=1)

        # Title di tengah atas
        tk.Label(dialog, text="Pilih Akun Anda", font=("Arial", 16, "bold")).grid(row=0, column=0, pady=8)

        container = tk.Frame(dialog)
        container.grid(row=1, column=0, sticky="nsew", padx=10, pady=5)

        pagination_frame = tk.Frame(dialog)  # pagination terpisah
        pagination_frame.grid(row=2, column=0, pady=5)

        save_button = tk.Button(dialog, text="Simpan Pilihan", state=tk.DISABLED,
                                font=("Arial", 14, "bold"), bg="#4CAF50", fg="white")
        save_button.grid(row=3, column=0, sticky="ew", padx=10, pady=10)

        selected_var = tk.StringVar()
        user_buttons = []
        current_page = 0
        ITEMS_PER_PAGE = 10

        def on_selection(*a):
            if selected_var.get():
                save_button.config(state=tk.NORMAL)
            else:
                save_button.config(state=tk.DISABLED)

        def set_selected(user_id, btn):
            for b in user_buttons:
                b.config(relief="raised", bg="SystemButtonFace")
            btn.config(relief="sunken", bg="#d0e0ff")
            selected_var.set(user_id)
            on_selection()

        def change_page(delta, user_list):
            nonlocal current_page
            total_pages = max(1, (len(user_list) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
            current_page = max(0, min(current_page + delta, total_pages - 1))
            populate_user_list(user_list)

        def populate_user_list(user_list):
            nonlocal current_page
            for w in container.winfo_children():
                w.destroy()
            for w in pagination_frame.winfo_children():
                w.destroy()
            user_buttons.clear()

            total_pages = max(1, (len(user_list) + ITEMS_PER_PAGE - 1) // ITEMS_PER_PAGE)
            start = current_page * ITEMS_PER_PAGE
            page_users = user_list[start:start + ITEMS_PER_PAGE]

            # Spacer top & bottom agar konten tengah vertikal
            container.grid_rowconfigure(0, weight=1)
            container.grid_rowconfigure(2, weight=1)
            container.grid_columnconfigure(0, weight=1)
            container.grid_columnconfigure(1, weight=1)

            content_frame = tk.Frame(container)
            content_frame.grid(row=1, column=0, columnspan=2)

            if len(page_users) <= 5:
                for i, user in enumerate(page_users):
                    name_cap = user.get('name', 'Nama tidak ditemukan').title()
                    btn = tk.Button(content_frame, text=name_cap, font=("Arial", 11, "bold"),  # Bold
                                    width=24, height=2, relief="raised")
                    btn.grid(row=i, column=0, columnspan=2, padx=8, pady=6)
                    btn.config(command=lambda u=user, b=btn: set_selected(u.get('_id'), b))
                    user_buttons.append(btn)
            else:
                for i, user in enumerate(page_users):
                    col = 0 if i < 5 else 1
                    row = i if i < 5 else i - 5
                    name_cap = user.get('name', 'Nama tidak ditemukan').title()
                    btn = tk.Button(content_frame, text=name_cap, font=("Arial", 11, "bold"),  # Bold
                                    width=18, height=2, relief="raised")
                    btn.grid(row=row, column=col, padx=8, pady=6)
                    btn.config(command=lambda u=user, b=btn: set_selected(u.get('_id'), b))
                    user_buttons.append(btn)

            # Pagination fixed di bawah
            tk.Button(pagination_frame, text="<< Prev",
                      state=tk.NORMAL if current_page > 0 else tk.DISABLED,
                      command=lambda: change_page(-1, user_list)).pack(side=tk.LEFT, padx=6)
            tk.Label(pagination_frame, text=f"Page {current_page+1} / {total_pages}").pack(side=tk.LEFT, padx=6)
            tk.Button(pagination_frame, text="Next >>",
                      state=tk.NORMAL if current_page < total_pages - 1 else tk.DISABLED,
                      command=lambda: change_page(1, user_list)).pack(side=tk.LEFT, padx=6)

        populate_user_list(filtered_users)

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
