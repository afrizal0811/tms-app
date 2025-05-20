import os
import tkinter as tk
from tkinter import filedialog, messagebox, ttk  # tambahkan ttk di sini
import openpyxl
import subprocess
import sys
import json

def pilih_file_routing():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Upload File", "Pilih Export Routing atau Hasil Combine Export Routing")
    return filedialog.askopenfilename(title="Pilih Export Routing", filetypes=[("Excel files", "*.xlsx")])

def simpan_file(wb):
    root = tk.Tk()
    root.withdraw()
    folder_path = filedialog.askdirectory(title="Pilih lokasi untuk menyimpan file")
    if not folder_path:
        return None

    base_filename = "Hasil Truck Detail"
    filename = f"{base_filename}.xlsx"
    full_path = os.path.join(folder_path, filename)
    counter = 1

    while os.path.exists(full_path):
        filename = f"{base_filename} - {counter}.xlsx"
        full_path = os.path.join(folder_path, filename)
        counter += 1

    wb.save(full_path)
    return full_path

def buka_file(path):
    if sys.platform.startswith('darwin'):
        subprocess.call(('open', path))
    elif os.name == 'nt':
        os.startfile(path)
    elif os.name == 'posix':
        subprocess.call(('xdg-open', path))

def buat_mapping_driver(master_path, lokasi_value):
    wb = openpyxl.load_workbook(master_path)
    ws = wb.active

    header = [str(cell.value).strip() if cell.value else "" for cell in ws[1]]
    email_idx = header.index("Email")
    driver_idx = header.index("Driver")

    mapping_list = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        email = row[email_idx]
        driver = row[driver_idx]
        if not email or not driver:
            continue
        if lokasi_value.lower() in str(email).lower():
            mapping_list.append((str(email).strip().lower(), str(driver).strip()))

    # Urutkan berdasarkan nama driver
    mapping_list.sort(key=lambda x: x[1])

    # Kembalikan sebagai dict
    return dict(mapping_list)

def pilih_lokasi():
    lokasi_dict = {
        "Sidoarjo": "plsda",
        "Jakarta": "pljkt",
        "Bandung": "plbdg",
        "Semarang": "plsmg",
        "Yogyakarta": "plygy",
        "Malang": "plmlg",
        "Denpasar": "pldps",
        "Makasar": "plmks",
        "Jember": "pljbr"
    }

    if getattr(sys, 'frozen', False):
        base_path = sys._MEIPASS  # folder temporary extract saat run .exe
        real_base = os.path.dirname(sys.executable)  # folder tempat .exe disimpan
    else:
        real_base = os.path.dirname(__file__)

    config_path = os.path.join(real_base, "config.json")
    if os.path.exists(config_path):
        try:
            with open(config_path, "r") as f:
                data = json.load(f)
                if "lokasi" in data:
                    return data["lokasi"]
        except:
            pass

    selected_value = None

    def simpan_dan_tutup():
        nonlocal selected_value
        selected = combo.get()
        if selected in lokasi_dict:
            selected_value = lokasi_dict[selected]
            with open(config_path, "w") as f:
                json.dump({"lokasi": selected_value}, f)
            root.destroy()

    root = tk.Tk()
    root.title("Lokasi Cabang Pangan Lestari")
    
    # Menentukan ukuran jendela
    window_width = 300
    window_height = 150 
    # Mendapatkan ukuran layar
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()

    # Menentukan posisi tengah
    position_x = (screen_width // 2) - (window_width // 2)
    position_y = (screen_height // 2) - (window_height // 2)

    # Menentukan ukuran dan posisi jendela
    root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
    tk.Label(root, text="Pilih Lokasi Cabang:", font=("Arial", 14)).pack(pady=10)

    combo = tk.ttk.Combobox(root, values=list(lokasi_dict.keys()), font=("Arial", 12))
    combo.pack(pady=10)
    combo.current(0)

    tk.Button(root, text="Pilih", command=simpan_dan_tutup, font=("Arial", 12), height=2, width=10).pack(pady=10)

    root.mainloop()
    return selected_value

def main():
    try:
        lokasi = pilih_lokasi()
        routing_path = pilih_file_routing()
        if not routing_path:
            messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")
            return

        # Ambil Master Driver di lokasi yang sama
        base_dir = os.path.dirname(os.path.abspath(__file__))
        master_path = os.path.join(base_dir, "Master_Driver.xlsx")
        if not os.path.exists(master_path):
            raise FileNotFoundError("File 'Master Driver.xlsx' tidak ditemukan di folder script.")

        email_to_name = buat_mapping_driver(master_path, lokasi)

        wb = openpyxl.load_workbook(routing_path, data_only=True)
        ws = wb.active

        # Cek apakah ada 'Capacity Constraint' di 10 baris pertama
        should_delete = False
        for row in ws.iter_rows(min_row=1, max_row=10, values_only=True):
            if any("capacity constraint" in str(cell).lower() for cell in row if cell):
                should_delete = True
                break

        if should_delete:
            for _ in range(10):
                ws.delete_rows(1)

        # Ambil header dan indeks kolom yang dibutuhkan
        header = [cell.value for cell in ws[1]]
        needed_headers = [
            "Vehicle Name",
            "Assignee",
            "Weight Percentage ",
            "Volume Percentage ",
            "Total Distance (m)",
            "Total Visits"
        ]

        # Mapping header -> index (1-based)
        col_indices = {h: i+1 for i, h in enumerate(header) if h in needed_headers}
        
        if "Assignee" not in col_indices:
            raise ValueError("Kolom 'Assignee' tidak ditemukan.")

        # Buat sheet baru untuk hasil
        new_wb = openpyxl.Workbook()
        new_ws = new_wb.active
        new_ws.title = "Filtered Data"

        # Tulis header dengan urutan baru
        new_header = [
            "Vehicle Name", "Assignee", "Weight Percentage", "Volume Percentage",
            "Total Distance (m)", "Total Visits", "Total Delivered", "Ship Duration"
        ]

        for col_num, header_name in enumerate(new_header, start=1):
            new_ws.cell(row=1, column=col_num, value=header_name)

        # Pastikan kolom 'Total Spent Time (mins)' tersedia
        spent_time_idx = next(
            (i for i, col in enumerate(header) if str(col).strip() == "Total Spent Time (mins)"), None
        )
        if spent_time_idx is None:
            raise ValueError("Kolom 'Total Spent Time (mins)' tidak ditemukan.")

        # Ambil dan salin data sesuai kolom yang diinginkan
        for i, row in enumerate(ws.iter_rows(min_row=2, values_only=True), start=2):
            new_row = []
            for h in new_header:
                if h == "Total Delivered":
                    new_row.append("")  # kolom kosong
                elif h == "Ship Duration":
                    val = row[spent_time_idx] if spent_time_idx < len(row) else ""
                    try:
                        minutes = float(str(val).replace(",", "").strip())
                        hours = int(minutes // 60)
                        mins = int(round(minutes % 60))
                        val = f"'{hours}:{mins:02d}"  # Menambahkan kutip satu untuk Google Sheets
                    except:
                        val = ""
                    new_row.append(val)
                else:
                    idx = next(i for i, col in enumerate(header) if str(col).strip() == h.strip())
                    new_row.append(row[idx] if idx < len(row) else None)

            for j, val in enumerate(new_row, start=1):
                col_name = new_header[j - 1]
                
                # Konversi ke persen jika perlu
                if col_name in ["Weight Percentage", "Volume Percentage"]:
                    try:
                        val_str = str(val).strip()
                        is_percent = val_str.endswith("%")
                        val_str = val_str.replace("%", "").replace(",", "").strip()
                        val_float = float(val_str)

                        if not is_percent and val_float < 10:
                            val_float *= 100

                        val = f"{val_float:.1f}%"
                    except:
                        val = ""

                if col_name == "Total Distance (m)":
                    try:
                        val_str = str(val).replace(",", "").replace(" ", "").strip()
                        val = int(float(val_str))
                    except:
                        val = 0
                new_ws.cell(row=i, column=j, value=val)


        # Ganti email jadi nama di kolom Assignee
        assignee_col_idx = new_header.index("Assignee") + 1
        for row in new_ws.iter_rows(min_row=2):
            cell = row[assignee_col_idx - 1]
            email = str(cell.value).strip().lower()
            if email in email_to_name:
                cell.value = email_to_name[email]

        # Sort data berdasarkan kolom Assignee
        data = list(new_ws.iter_rows(min_row=2, values_only=True))
        data.sort(key=lambda row: row[assignee_col_idx - 1] or "")

        # Hapus data lama, tulis ulang hasil sort
        new_ws.delete_rows(2, new_ws.max_row)
        for i, row_data in enumerate(data, start=2):
            for j, value in enumerate(row_data, start=1):
                new_ws.cell(row=i, column=j, value=value)

        # Tambah driver dari Master Driver yang belum ada di Filtered Data
        existing_drivers = set()
        for row in new_ws.iter_rows(min_row=2, values_only=True):
            if row[1]:
                existing_drivers.add(str(row[1]).strip())

        all_drivers = set(email_to_name.values())
        missing_drivers = all_drivers - existing_drivers

        start_row = new_ws.max_row + 1
        for i, driver in enumerate(sorted(missing_drivers), start=start_row):
            new_ws.cell(row=i, column=2, value=driver)  # isi nama driver di kolom Assignee
            # kolom lain dibiarkan kosong

        # Sort ulang berdasarkan kolom Assignee (kolom 2)
        all_data = list(new_ws.iter_rows(min_row=2, max_row=new_ws.max_row, values_only=True))
        all_data.sort(key=lambda r: r[1] or "")

        # Hapus data lama
        new_ws.delete_rows(2, new_ws.max_row)

        # Tulis ulang data yang sudah diurutkan
        for i, row_data in enumerate(all_data, start=2):
            for j, value in enumerate(row_data, start=1):
                new_ws.cell(row=i, column=j, value=value)

        # Tambahkan sheet 'Type Total Distance'
        type_ws = new_wb.create_sheet(title="Type Total Distance")
        type_ws["A1"] = "DRY"
        type_ws["B1"] = "FRZ"

        # Filter berdasarkan Assignee dan total distance untuk DRY dan FRZ
        dry_total = 0
        frz_total = 0
        for row in data:
            assignee = row[1]  # Assignee berada di kolom 2 (index 1)
            total_distance = row[4]  # Total Distance (m) berada di kolom 5 (index 4)
            if assignee and isinstance(total_distance, (int, float)):
                if "DRY" in assignee.upper():
                    dry_total += total_distance
                elif "FRZ" in assignee.upper():
                    frz_total += total_distance

        type_ws["A2"] = dry_total
        type_ws["B2"] = frz_total

        # Sesuaikan lebar kolom berdasarkan isi terpanjang di 'Filtered Data'
        for col in new_ws.columns:
            max_length = 0
            col_letter = openpyxl.utils.get_column_letter(col[0].column)
            for cell in col:
                try:
                    if cell.value:
                        max_length = max(max_length, len(str(cell.value)))
                except:
                    pass
            adjusted_width = max_length + 2  # Tambahan padding
            new_ws.column_dimensions[col_letter].width = adjusted_width


        save_path = simpan_file(new_wb)
        if save_path:
            buka_file(save_path)
        else:
            messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")

    except Exception as e:
        messagebox.showerror("Terjadi Kesalahan", str(e))

if __name__ == "__main__":
    main()
