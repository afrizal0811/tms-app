# utils/gui.py

from tkcalendar import DateEntry
from tkinter import ttk
import threading
import tkinter as tk
import time
from utils.function import show_error_message, show_info_message

def create_date_picker_window(title, process_callback):
    """
    Membuat dan menampilkan jendela GUI pemilih tanggal yang dapat digunakan kembali,
    termasuk label status, progress bar, dan timer.

    Args:
        title (str): Judul untuk jendela GUI.
        process_callback (function): Fungsi yang akan dipanggil dengan tanggal yang dipilih.
    """
    
    class DatePickerApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(title)
            
            # Pengaturan Geometri Window
            window_width = 350
            window_height = 250
            screen_width = self.winfo_screenwidth()
            screen_height = self.winfo_screenheight()
            center_x = int(screen_width/2 - window_width / 2)
            center_y = int(screen_height/2 - window_height / 2)
            self.geometry(f'{window_width}x{window_height}+{center_x}+{center_y}')
            self.config(bg='SystemButtonFace')

            # Konfigurasi Style
            style = ttk.Style(self)
            style.theme_use('clam')
            style.configure("TButton", font=("Helvetica", 12), padding=5)
            style.configure("TLabel", background='SystemButtonFace', font=("Helvetica", 16, "bold"))
            style.configure("TProgressbar", thickness=20)
            
            main_frame = tk.Frame(self, bg='SystemButtonFace')
            main_frame.pack(expand=True, pady=20)

            label = ttk.Label(main_frame, text="Pilih Tanggal Pengiriman")
            label.pack(pady=(0, 10))

            # Widget DateEntry
            self.cal = DateEntry(main_frame, date_pattern='dd-MM-yyyy', font=("Helvetica", 16), 
                                 width=12, justify='center', borderwidth=2, relief="solid",
                                 style='TEntry')
            self.cal.pack(pady=10, ipady=5)
            self.cal.bind("<Button-1>", self._on_date_click)

            # Tombol Proses
            self.run_button = ttk.Button(main_frame, text="Proses", command=self.run_process_thread, style="TButton")
            self.run_button.pack(pady=10)

            # Label Status
            self.status_label = ttk.Label(main_frame, text="", foreground="blue", font=("Helvetica", 10))
            self.status_label.pack(pady=(5, 0))

            # Widget Timer baru
            self.timer_label = ttk.Label(main_frame, text="", font=("Arial", 9), foreground="gray")
            self.timer_label.pack()

            # Progress Bar (disembunyikan secara default)
            self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="indeterminate")
            
        def _on_date_click(self, event):
            """Membuka kalender dropdown saat widget diklik."""
            self.cal.drop_down()

        def run_process_thread(self):
            """Menjalankan proses di thread terpisah agar GUI tidak freeze."""
            # Sembunyikan tombol dan tampilkan progress bar
            self.run_button.pack_forget()
            self.progress.pack(pady=(10, 5))
            self.progress.start(10)
            
            selected_date_obj = self.cal.get_date()
            date_formats = {
                "dmy": selected_date_obj.strftime('%d-%m-%Y'),
                "ymd": selected_date_obj.strftime('%Y-%m-%d')
            }
            
            # Nonaktifkan kalender
            self.cal.config(state='disabled')
            
            # Mulai timer
            self.start_time = time.time()
            self.timer_running = True
            self.update_timer()

            # Buat thread untuk menjalankan fungsi callback
            self.process_thread = threading.Thread(
                target=process_callback,
                args=(date_formats, self)
            )
            self.process_thread.start()
            self.after(100, self.check_thread)

        def check_thread(self):
            """Memeriksa apakah thread sudah selesai."""
            if self.process_thread.is_alive():
                self.after(100, self.check_thread)
            else:
                self.timer_running = False
                self.progress.stop()
                self.progress.pack_forget()
                self.status_label.config(text="")
                self.timer_label.config(text="")
                self.cal.config(state='normal')
                self.run_button.pack(pady=10)

        def update_status(self, message):
            """Fungsi untuk update status label dari thread lain."""
            # Menggunakan after untuk menjalankan update di thread utama
            if self.winfo_exists():
                self.after(0, lambda: self.status_label.config(text=message))

        def display_error(self, title, message):
            """Fungsi jembatan untuk menampilkan error pop-up dengan aman."""
            if self.winfo_exists():
                self.after(0, lambda: show_error_message(title, message))
                self.update_status(f"Error: {message}")

        def display_info(self, title, message):
            """Fungsi jembatan untuk menampilkan info pop-up dengan aman."""
            if self.winfo_exists():
                self.after(0, lambda: show_info_message(title, message))
                self.update_status(f"Info: {message}")
        
        def update_timer(self):
            """Fungsi untuk update label timer."""
            if self.timer_running and self.winfo_exists():
                elapsed_time = int(time.time() - self.start_time)
                hours = elapsed_time // 3600
                minutes = (elapsed_time % 3600) // 60
                seconds = elapsed_time % 60
                self.timer_label.config(text=f"{hours:02}:{minutes:02}:{seconds:02}")
                self.after(1000, self.update_timer)

    app = DatePickerApp()
    app.mainloop()