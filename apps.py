import os
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
import traceback

def find_vehicle_tag(vehicle_name, plat_list, plat_type_map):
    for plat in plat_list:
        if pd.isna(vehicle_name):
            return ""
        if str(plat) in str(vehicle_name):
            return plat_type_map.get(plat, "")
    return ""

def contains_capacity_constraint(file_path):
    try:
        preview_df = pd.read_excel(file_path, nrows=20, header=None)
        return preview_df.astype(str).apply(lambda row: row.str.contains("capacity constraint", case=False, na=False)).any().any()
    except:
        return False

def main():
    try:
        root = tk.Tk()
        root.withdraw()

        messagebox.showinfo("Pilih File", "Pilih Export Routing atau Hasil Combine Export Routing")
        file_path = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])

        if not file_path:
            messagebox.showwarning("Peringatan", "Proses Dibatalkan")
            return

        master_driver_path = os.path.join(os.path.dirname(__file__), "Master_Driver.xlsx")
        if not os.path.exists(master_driver_path):
            messagebox.showerror("File Tidak Ditemukan", f"Master_Driver.xlsx tidak ditemukan di folder:\n{os.path.dirname(__file__)}")
            return

        master_df = pd.read_excel(master_driver_path)

        skip_rows = 10 if contains_capacity_constraint(file_path) else 0
        upload_df = pd.read_excel(file_path, skiprows=skip_rows)

        # Mapping Assignee ke Driver
        email_to_driver = dict(zip(master_df["Email"], master_df["Driver"]))
        upload_df["Assignee"] = upload_df["Assignee"].map(email_to_driver).fillna(upload_df["Assignee"])

        # Mapping Vehicle Name ke Type → Vehicle Tags
        plat_type_map = dict(zip(master_df["Plat"], master_df["Type"]))
        plat_list = list(master_df["Plat"])
        upload_df["Vehicle Tags"] = upload_df["Vehicle Name"].apply(
            lambda name: find_vehicle_tag(name, plat_list, plat_type_map)
        )

        # Pisahkan DRY dan FROZEN
        tag_series = upload_df["Vehicle Tags"].astype(str)
        dry_df = upload_df[tag_series.str.contains("DRY", case=False, na=False)].reset_index(drop=True)
        frozen_df = upload_df[tag_series.str.contains("FROZEN", case=False, na=False)].reset_index(drop=True)

        # Daftar tipe kendaraan sesuai permintaan
        vehicle_types = ["L300", "CDE", "CDE-Long", "CDD", "CDD-Long", "Fuso"]

        def count_types_separated(tags):
            counts = {vt:0 for vt in vehicle_types}
            used_indices = set()

            # cek tipe Long dulu supaya ga kebaca ke tipe pendek
            long_types = [t for t in vehicle_types if "-Long" in t]
            short_types = [t for t in vehicle_types if "-Long" not in t]

            # cek tipe Long dulu
            for vtype in long_types:
                match = tags[~tags.index.isin(used_indices)].str.contains(rf"\b{vtype}\b", case=False, na=False)
                matched_indices = match[match].index.tolist()
                counts[vtype] = len(matched_indices)
                used_indices.update(matched_indices)

            # cek tipe pendek setelahnya
            for vtype in short_types:
                match = tags[~tags.index.isin(used_indices)].str.contains(rf"\b{vtype}\b", case=False, na=False)
                matched_indices = match[match].index.tolist()
                counts[vtype] = len(matched_indices)
                used_indices.update(matched_indices)

            return counts

        dry_counts = count_types_separated(dry_df["Vehicle Tags"].astype(str))
        frozen_counts = count_types_separated(frozen_df["Vehicle Tags"].astype(str))

        def format_counts(count_dict):
            lines = []
            for vt in vehicle_types:
                c = count_dict.get(vt, 0)
                if c > 0:
                    lines.append(f"{vt.ljust(10)}: {c}")
            return "\n".join(lines)

        dry_result = format_counts(dry_counts)
        frozen_result = format_counts(frozen_counts)

        full_message = ""
        if dry_result:
            full_message += "[DRY]\n" + dry_result + "\n\n"
        if frozen_result:
            full_message += "[FROZEN]\n" + frozen_result
        if not full_message.strip():
            full_message = "Tidak ada kendaraan yang terdeteksi."

        def show_result_window(title, message):
            result_window = tk.Toplevel()
            result_window.title(title)

            text_widget = tk.Text(result_window, font=("Courier New", 11), padx=10, pady=10, bg="white")
            text_widget.insert("1.0", message)
            text_widget.config(state="disabled")
            text_widget.pack(expand=True, fill="both")

            width = 400
            height = 300
            screen_width = result_window.winfo_screenwidth()
            screen_height = result_window.winfo_screenheight()
            x = (screen_width // 2) - (width // 2)
            y = (screen_height // 2) - (height // 2)
            result_window.geometry(f"{width}x{height}+{x}+{y}")

            result_window.transient()
            result_window.grab_set()
            result_window.wait_window()

        show_result_window("Ringkasan Kendaraan", full_message.strip())

    except Exception as e:
        error_message = traceback.format_exc()
        messagebox.showerror("Terjadi Error", error_message)

if __name__ == "__main__":
    main()
