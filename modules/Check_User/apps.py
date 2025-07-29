import tkinter as tk
from tkinter import ttk, messagebox
import requests
import json
import os

# --- Impor dari shared_utils ---
# Menggunakan fungsi terpusat untuk menangani path dan file
from modules.shared_utils import (
    load_config,
    load_constants,
    save_json_data,
    CONFIG_PATH
)


class ApiGuiApp:
    """
    Kelas untuk aplikasi GUI yang memanggil API MileApp.
    """
    def __init__(self, master):
        """
        Inisialisasi aplikasi.
        
        Args:
            master: Jendela utama (root) dari tkinter.
        """
        self.master = master
        master.title("MileApp User Checker")
        
        self.users_data = {}
        self.selected_user_var = tk.StringVar()

        self.bg_color = "#f0f0f0"
        self.font_label_bold = ("Arial", 12, "bold")
        self.font_button = ("Arial", 12, "bold")
        self.font_radio_result = ("Arial", 10) 
        self.font_instruction = ("Arial", 9, "italic") 
        
        master.configure(bg=self.bg_color)

        main_frame = tk.Frame(master, padx=20, pady=20, bg=self.bg_color)
        main_frame.pack(fill=tk.BOTH, expand=True)
        main_frame.grid_rowconfigure(2, weight=1)
        main_frame.grid_columnconfigure(0, weight=1)

        self.label = tk.Label(main_frame, text="Masukkan Nama Pengguna di MileApp", font=self.font_label_bold, bg=self.bg_color)
        self.label.grid(row=0, column=0, pady=(0, 10), sticky="ew")

        self.entry = tk.Entry(main_frame, font=("Arial", 12), width=50)
        self.entry.grid(row=1, column=0, pady=(0, 10), sticky="ew")
        self.entry.bind("<Return>", self.search_user_event)

        # --- Implementasi Scrollbar ---
        result_container = tk.Frame(main_frame, borderwidth=1, relief="solid")
        result_container.grid(row=2, column=0, pady=(5, 15), sticky="nsew")
        result_container.grid_rowconfigure(0, weight=1)
        result_container.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(result_container, bg="white", highlightthickness=0)
        self.scrollbar = ttk.Scrollbar(result_container, orient="vertical", command=self.canvas.yview)
        self.scrollable_frame = tk.Frame(self.canvas, bg="white")

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        )

        self.canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        self.canvas.configure(yscrollcommand=self.scrollbar.set)
        
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.scrollbar.grid(row=0, column=1, sticky="ns")
        
        self.master.bind_all("<MouseWheel>", self._on_mousewheel)
        # --- Akhir Implementasi Scrollbar ---
        
        button_frame = tk.Frame(main_frame, bg=self.bg_color)
        button_frame.grid(row=3, column=0, sticky="ew")
        
        self.search_button = tk.Button(button_frame, text="Cari Pengguna", command=self.search_user, font=self.font_button, bg="#2196F3", fg="white", relief=tk.FLAT)
        self.search_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=5, padx=(0, 5))
        
        self.save_button = tk.Button(button_frame, text="Simpan Pilihan", command=self.save_selection, font=self.font_button, bg="#4CAF50", fg="white", relief=tk.FLAT, state=tk.DISABLED)
        self.save_button.pack(side=tk.LEFT, expand=True, fill=tk.X, ipady=5, padx=(5, 0))

        self._center_window(600, 500)

    def _on_mousewheel(self, event):
        """Fungsi untuk menangani scroll dengan mouse wheel."""
        if self.canvas.winfo_exists():
            self.canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    def _center_window(self, width=600, height=500):
        """Menempatkan jendela di tengah layar."""
        screen_width = self.master.winfo_screenwidth()
        screen_height = self.master.winfo_screenheight()
        x = (screen_width / 2) - (width / 2)
        y = (screen_height / 2) - (height / 2)
        self.master.geometry(f'{width}x{height}+{int(x)}+{int(y)}')

    def clear_results(self):
        """Membersihkan hasil pencarian sebelumnya."""
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()
        self.save_button.config(state=tk.DISABLED)
        self.selected_user_var.set("")

    def search_user(self):
        self.clear_results()
        user_name_input = self.entry.get().strip()
        
        if len(user_name_input) < 4:
            error_label = tk.Label(self.scrollable_frame, 
                                   text="Harap masukkan minimal 4 karakter untuk memulai pencarian.", 
                                   bg="white", fg="red", font=("Arial", 10))
            error_label.pack(pady=10)
            return

        # Menggunakan fungsi terpusat untuk memuat data
        config = load_config()
        constants = load_constants()

        if not config or not constants:
            messagebox.showerror("Error", "Gagal memuat file konfigurasi (config.json atau constant.json).")
            return
            
        lokasi = config.get('lokasi')
        api_token = constants.get('token')
        hub_id = constants.get('hub_ids', {}).get(lokasi)

        if not api_token or not hub_id:
            messagebox.showerror("Error Konfigurasi", "Token API atau Hub ID tidak ditemukan. Periksa file konfigurasi Anda.")
            return

        loading_label = tk.Label(self.scrollable_frame, text=f"Mencari pengguna '{user_name_input}'...", bg="white")
        loading_label.pack(pady=10)
        self.master.update_idletasks()

        params = {'q': user_name_input, 'hubId': hub_id}
        headers = {'Authorization': f'Bearer {api_token}', 'Content-Type': 'application/json'}
        
        try:
            response = requests.get("https://apiweb.mile.app/api/v3/users", headers=headers, params=params, timeout=10)
            response.raise_for_status()
            response_data = response.json()
            users = response_data.get('data')
            
            loading_label.destroy()

            if users and isinstance(users, list):
                self.users_data.clear()
                
                role_id_to_exclude = "6703410af6be892f3208ecde"
                users_after_role_filter = [user for user in users if user.get('roleId') != role_id_to_exclude]

                users_after_name_filter = [
                    user for user in users_after_role_filter 
                    if user_name_input.lower() in user.get('name', '').lower()
                ]

                users_after_name_filter.sort(key=lambda u: u.get('name', '').lower())

                if not users_after_name_filter:
                    tk.Label(self.scrollable_frame, text=f"Tidak ada pengguna yang cocok dengan nama '{user_name_input}'.", bg="white").pack(pady=10)
                    self.master.focus_set() 
                    return

                instruction_label = tk.Label(self.scrollable_frame, text="Pilih salah satu nama di bawah ini:", font=self.font_instruction, bg="white")
                instruction_label.pack(pady=(5, 2), anchor="w", padx=10)

                for user in users_after_name_filter:
                    user_id = user.get('_id')
                    user_name_display = user.get('name', 'Nama tidak ditemukan').title()
                    self.users_data[user_id] = user 
                    
                    rb = tk.Radiobutton(self.scrollable_frame, text=user_name_display, variable=self.selected_user_var, 
                                        value=user_id, bg="white", anchor="w",
                                        font=self.font_radio_result,
                                        command=lambda: self.save_button.config(state=tk.NORMAL))
                    rb.pack(fill=tk.X, padx=10, pady=0) 
                
                self.master.focus_set()
            else:
                tk.Label(self.scrollable_frame, text=f"Tidak ada pengguna yang cocok dengan nama '{user_name_input}'.", bg="white").pack(pady=10)
                self.master.focus_set()

        except requests.exceptions.RequestException as e:
            loading_label.destroy()
            tk.Label(self.scrollable_frame, text=f"Gagal menghubungi API: {e}", bg="white", fg="red", wraplength=550).pack(pady=10)
            self.master.focus_set()

    def save_selection(self):
        selected_id = self.selected_user_var.get()
        if not selected_id:
            messagebox.showwarning("Peringatan", "Tidak ada pengguna yang dipilih.")
            return

        selected_user = self.users_data.get(selected_id)
        if not selected_user:
            messagebox.showerror("Error", "Data pengguna yang dipilih tidak ditemukan.")
            return
        
        confirm_msg = f"Anda akan memilih '{selected_user['name']}'.\n\nPilihan ini bersifat final dan tidak dapat diubah lagi dari aplikasi ini.\n\nApakah Anda yakin?"
        is_confirmed = messagebox.askyesno("Konfirmasi Pilihan", confirm_msg)

        if is_confirmed:
            try:
                # Membaca file config.json untuk diperbarui
                config_data = load_config()
                if config_data is None:
                    messagebox.showerror("Error", "Gagal membaca file 'config.json' sebelum menyimpan.")
                    return
                
                # Menambahkan data user_checked ke data config
                config_data['user_checked'] = {
                    'name': selected_user.get('name'),
                    '_id': selected_user.get('_id'),
                    'hub_id': selected_user.get('hubId'),
                    'role_id': selected_user.get('roleId')
                }

                # Menyimpan kembali data yang sudah diperbarui ke config.json
                if save_json_data(config_data, CONFIG_PATH):
                    messagebox.showinfo("Sukses", f"Pengguna '{selected_user['name']}' telah berhasil disimpan.")
                    self.master.destroy()
                # Pesan error sudah ditangani oleh save_json_data jika gagal

            except Exception as e:
                messagebox.showerror("Error Tak Terduga", f"Terjadi kesalahan saat proses penyimpanan:\n{e}")

    def search_user_event(self, event):
        self.search_user()

if __name__ == "__main__":
    root = tk.Tk()
    app = ApiGuiApp(root)
    root.mainloop()
