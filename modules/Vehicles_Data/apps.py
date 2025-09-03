import tkinter as tk
from tkinter import ttk, messagebox
import requests
import pandas as pd
import openpyxl
from datetime import datetime
import traceback
import os, sys

# --- import utilmu tetap sama ---
from utils.function import (
    get_save_path,
    load_config,
    load_constants,
    load_master_data,
    load_secret,
    open_file_externally,
    show_error_message,
    show_info_message
)
from utils.messages import ERROR_MESSAGES
from utils.api_handler import handle_requests_error


def auto_size_columns(workbook):
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if cell.value and len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except (ValueError, TypeError):
                    pass
            adjusted_width = min(max_length + 2, 50)
            worksheet.column_dimensions[column].width = adjusted_width


def fetch_and_prepare_data():
    """Ambil data API, kelola, return dict dataframe tanpa simpan file."""
    config = load_config()
    constants = load_constants()
    secrets = load_secret()
    master_data = load_master_data()

    if not (config and constants and secrets and master_data):
        return None, None

    api_token = secrets.get("token")
    lokasi_code = config.get("lokasi")
    hub_ids = master_data.get("hub_ids", {})
    lokasi_mapping = constants.get("lokasi_mapping", {})

    if not (api_token and lokasi_code and lokasi_code in hub_ids):
        return None, None

    hub_id = hub_ids[lokasi_code]
    base_url = constants.get("base_url")
    api_url = f"{base_url}/vehicles"

    params = {"limit": 500, "hubId": hub_id}
    headers = {"Authorization": f"Bearer {api_token}"}

    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        vehicles_data = response.json().get("data", [])
        if not vehicles_data:
            return None, None

        # --- olah data sama seperti kode lamamu ---
        master_df = master_data["df"]
        driver_mapping = dict(zip(master_df["Email"].str.lower(), master_df["Driver"]))

        template_data = []
        master_data_list = []

        for vehicle in vehicles_data:
            working_time = vehicle.get("workingTime", {})
            break_time = vehicle.get("breakTime", {})
            capacity = vehicle.get("capacity", {})
            weight_cap = capacity.get("weight", {})
            volume_cap = capacity.get("volume", {})
            tags = vehicle.get("tags", [])

            template_data.append({
                "Name*": vehicle.get("name", ""),
                "Assignee": vehicle.get("assignee", ""),
                "Start Time": working_time.get("startTime", ""),
                "End Time": working_time.get("endTime", ""),
                "Break Start": break_time.get("startTime") or 0,
                "Break End": break_time.get("endTime") or 0,
                "Multiday": working_time.get("multiday") or 0,
                "Speed Km/h": vehicle.get("speed", 0),
                "Cost Factor": vehicle.get("fixedCost", 0),
                "Vehicle Tags": "; ".join(tags),
                "weight Min": weight_cap.get("min", ""),
                "weight Max": weight_cap.get("max", ""),
                "volume Min": volume_cap.get("min", ""),
                "volume Max": volume_cap.get("max", ""),
            })

            assignee_email = (vehicle.get("assignee", "") or "").lower()
            vehicle_type_raw = tags[0] if tags else ""
            if vehicle_type_raw == "FROZEN-KFC":
                vehicle_type_raw = "FROZEN-CDD-LONG-5000"
            elif vehicle_type_raw == "DRY-HAVI":
                vehicle_type_raw = "DRY-FUSO-LONG"

            driver_name = driver_mapping.get(assignee_email, assignee_email)
            master_data_list.append({
                "License Plat": vehicle.get("name", ""),
                "Type": vehicle_type_raw,
                "Email": assignee_email,
                "Name": driver_name,
            })

        df_template = pd.DataFrame(template_data)
        df_master = pd.DataFrame(master_data_list)

        # --- buat dict sheet ---
        dfs = {"Master Vehicle": df_master, "Template Vehicle": df_template}
        # conditional vehicle dikelola juga kalau ada
        # (aku potong biar singkat, bisa disalin dari kode lama)

        lokasi_name = next((n for n, c in lokasi_mapping.items() if c == lokasi_code), lokasi_code)
        return dfs, lokasi_name

    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
    except Exception as e:
        show_error_message("Error Tak Terduga", ERROR_MESSAGES["UNKNOWN_ERROR"].format(
            error_detail=f"{e}\n\n{traceback.format_exc()}"
        ))
    return None, None


def show_excel_viewer(dfs, lokasi_name):
    """Menampilkan hasil DataFrame ke GUI viewer dengan opsi download."""
    viewer = tk.Toplevel()
    viewer.title(f"Vehicle Data Viewer - {lokasi_name}")

    # --- ukuran dan posisi tengah layar ---
    w, h = 900, 600
    sw, sh = viewer.winfo_screenwidth(), viewer.winfo_screenheight()
    x, y = (sw - w) // 2, (sh - h) // 2
    viewer.geometry(f"{w}x{h}+{x}+{y}")

    # --- notebook (atas) ---
    notebook = ttk.Notebook(viewer)
    notebook.pack(fill="both", expand=True, padx=5, pady=(5, 0))

    # Tampilkan semua DataFrame di tab
    for sheet_name, df in dfs.items():
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=sheet_name)

        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        tree = ttk.Treeview(frame, columns=list(df.columns), show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        # Scrollbar vertikal
        v_scroll = ttk.Scrollbar(frame, orient="vertical", command=tree.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=v_scroll.set)

        # Scrollbar horizontal
        h_scroll = ttk.Scrollbar(frame, orient="horizontal", command=tree.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(xscrollcommand=h_scroll.set)

        # Heading dan lebar kolom
        for col in df.columns:
            tree.heading(col, text=col)
            tree.column(col, width=120, anchor="center")

        # Isi data
        for _, row in df.iterrows():
            values = [("" if pd.isna(x) else x) for x in row]
            tree.insert("", "end", values=values)

    # --- tombol download (bawah) ---
    button_frame = ttk.Frame(viewer)
    button_frame.pack(pady=8)

    def download_excel():
        date_str = datetime.now().strftime("%d.%m.%Y")
        file_basename = f"Vehicle Data {lokasi_name} - {date_str}"
        save_path = get_save_path(file_basename)
        if not save_path:
            return
        try:
            with pd.ExcelWriter(save_path, engine="openpyxl") as writer:
                for sheet_name, df in dfs.items():
                    df.to_excel(writer, index=False, sheet_name=sheet_name, na_rep="")
            workbook = openpyxl.load_workbook(save_path)
            auto_size_columns(workbook)
            workbook.save(save_path)
            open_file_externally(save_path)
        except Exception as e:
            messagebox.showerror("Error", f"Gagal menyimpan file:\n{e}")

    ttk.Button(button_frame, text="Download", command=download_excel).pack()
    viewer.mainloop()

def main():
    dfs, lokasi_name = fetch_and_prepare_data()
    if dfs:
        show_excel_viewer(dfs, lokasi_name)


if __name__ == "__main__":
    main()
