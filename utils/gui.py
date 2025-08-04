# modules/gui_utils.py (FILE BARU)

from tkcalendar import DateEntry
from tkinter import ttk
import threading
import tkinter as tk

def create_date_picker_window(title, process_callback):
    """
    Membuat dan menampilkan jendela GUI pemilih tanggal yang dapat digunakan kembali.

    Args:
        title (str): Judul untuk jendela GUI.
        process_callback (function): Fungsi yang akan dipanggil dengan tanggal yang dipilih 
                                     saat tombol 'Proses' diklik.
    """
    
    class DatePickerApp(tk.Tk):
        def __init__(self):
            super().__init__()
            self.title(title)
            
            # Pengaturan Geometri Window
            window_width = 350
            window_height = 220
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

            label = ttk.Label(main_frame, text="Pilih Tanggal")
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

            # Progress Bar dan Status Label
            self.progress = ttk.Progressbar(main_frame, orient="horizontal", length=300, mode="indeterminate")
            self.status_label = ttk.Label(main_frame, text="", foreground="blue", font=("Helvetica", 10))
            self.status_label.pack(pady=5)

        def _on_date_click(self, event):
            """Membuka kalender dropdown saat widget diklik."""
            self.cal.drop_down()

        def run_process_thread(self):
            """Menjalankan proses di thread terpisah agar GUI tidak freeze."""
            self.run_button.pack_forget()
            self.progress.pack(pady=10)
            self.progress.start(10)
            
            selected_date_obj = self.cal.get_date()
            # Callback untuk time.py mengharapkan format 'dd-MM-yyyy'
            # Callback untuk auto_delivery.py mengharapkan format 'yyyy-MM-dd'
            # Kita kirim keduanya untuk fleksibilitas
            date_formats = {
                "dmy": selected_date_obj.strftime('%d-%m-%Y'),
                "ymd": selected_date_obj.strftime('%Y-%m-%d')
            }
            
            # Buat thread untuk menjalankan fungsi callback
            self.process_thread = threading.Thread(
                target=process_callback,
                args=(date_formats, self) # Mengirim instance App sebagai argumen kedua
            )
            self.process_thread.start()
            self.after(100, self.check_thread)

        def check_thread(self):
            """Memeriksa apakah thread sudah selesai."""
            if self.process_thread.is_alive():
                self.after(100, self.check_thread)
            else:
                self.progress.stop()
                self.progress.pack_forget()
                self.status_label.config(text="")
                self.run_button.pack(pady=10)

        def update_status(self, message):
            """Fungsi untuk update status label dari thread lain."""
            if self.winfo_exists():
                self.status_label.config(text=message)
    
    # Inisialisasi dan jalankan aplikasi
    app = DatePickerApp()
    app.mainloop()