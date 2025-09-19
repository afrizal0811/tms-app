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
    location_id = constants.get("location_id", {})

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

        master_df_config = master_data["df"]
        driver_mapping = dict(zip(master_df_config["Email"].str.lower(), master_df_config["Driver"]))

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
        df_conditional = pd.DataFrame()

        if not df_master.empty and "Email" in df_master.columns and "License Plat" in df_master.columns:
            df_master['plat_len'] = df_master['License Plat'].astype(str).str.len()
            df_master = df_master.sort_values(by=['Email', 'plat_len'], ascending=[True, True])
            duplicate_mask = df_master.duplicated(subset='Email', keep='first')
            
            if duplicate_mask.any():
                df_conditional = df_master[duplicate_mask].copy()
                df_master = df_master[~duplicate_mask].copy()

            df_master = df_master.drop(columns=['plat_len'])
            if not df_conditional.empty:
                df_conditional = df_conditional.drop(columns=['plat_len'])

        if not df_master.empty and "Email" in df_master.columns:
            df_master = df_master.sort_values(by="Email", ascending=True).reset_index(drop=True)
        if not df_template.empty and "Assignee" in df_template.columns:
            df_template = df_template.sort_values(by="Assignee", ascending=True).reset_index(drop=True)

        # Mulai dengan Master Vehicle
        dfs = {"Master Vehicle": df_master}

        # Jika ada data duplikat, sisipkan Conditional Vehicle di urutan kedua
        if not df_conditional.empty:
            dfs["Conditional Vehicle"] = df_conditional.sort_values(by="Email", ascending=True).reset_index(drop=True)
        
        # Tambahkan Template Vehicle di urutan terakhir
        dfs["Template Vehicle"] = df_template

        lokasi_name = next((n for n, c in location_id.items() if c == lokasi_code), lokasi_code)
        return dfs, lokasi_name

    except requests.exceptions.RequestException as e:
        handle_requests_error(e)
    except Exception as e:
        show_error_message("Error Tak Terduga", ERROR_MESSAGES["UNKNOWN_ERROR"].format(
            error_detail=f"{e}\n\n{traceback.format_exc()}"
        ))
    return None, None


def show_excel_viewer(dfs, lokasi_name):
    """Menampilkan hasil DataFrame ke GUI viewer dengan opsi download dan filtering."""
    viewer = tk.Toplevel()
    viewer.title(f"Vehicle Data Viewer - {lokasi_name}")

    w, h = 900, 600
    sw, sh = viewer.winfo_screenwidth(), viewer.winfo_screenheight()
    x, y = (sw - w) // 2, (sh - h) // 2
    viewer.geometry(f"{w}x{h}+{x}+{y}")

    notebook = ttk.Notebook(viewer)
    notebook.pack(fill="both", expand=True, padx=5, pady=(5, 0))
    allowed_filters = {
        "Master Vehicle": ["License Plat", "Email",  "Name", "Type"],
        "Conditional Vehicle": ["License Plat", "Email", "Name", "Type"],
        "Template Vehicle": ["Name*", "Assignee"]
    }

    for sheet_name, df in dfs.items():
        frame = ttk.Frame(notebook)
        notebook.add(frame, text=sheet_name)

        filter_frame = ttk.Frame(frame)
        filter_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(filter_frame, text="Filter by Column:").pack(side="left", padx=(0, 5))
        
        columns_to_show = allowed_filters.get(sheet_name, list(df.columns))
        filter_column = ttk.Combobox(filter_frame, values=columns_to_show, state="readonly")
        
        filter_column.pack(side="left", padx=5)
        if columns_to_show:
            filter_column.current(0)

        ttk.Label(filter_frame, text="Keyword:").pack(side="left", padx=(10, 5))
        
        filter_entry = ttk.Entry(filter_frame)
        filter_entry.pack(side="left", fill="x", expand=True, padx=5)

        tree_frame = ttk.Frame(frame)
        tree_frame.pack(fill="both", expand=True)

        tree = ttk.Treeview(tree_frame, columns=list(df.columns), show="headings")
        tree.grid(row=0, column=0, sticky="nsew")

        v_scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
        v_scroll.grid(row=0, column=1, sticky="ns")
        tree.configure(yscrollcommand=v_scroll.set)

        h_scroll = ttk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
        h_scroll.grid(row=1, column=0, sticky="ew")
        tree.configure(xscrollcommand=h_scroll.set)
        
        tree_frame.rowconfigure(0, weight=1)
        tree_frame.columnconfigure(0, weight=1)

        def populate_tree(tree_widget, dataframe):
            for i in tree_widget.get_children():
                tree_widget.delete(i)
            for col in dataframe.columns:
                tree_widget.heading(col, text=col)
                tree_widget.column(col, width=120, anchor="center")
            for _, row in dataframe.iterrows():
                values = [("" if pd.isna(x) else x) for x in row]
                tree_widget.insert("", "end", values=values)

        populate_tree(tree, df)

        def apply_filter(dataframe, tree_widget, column_widget, entry_widget):
            keyword = entry_widget.get().strip()
            column = column_widget.get()
            if not keyword or not column:
                populate_tree(tree_widget, dataframe)
                return
            filtered_df = dataframe[dataframe[column].astype(str).str.contains(keyword, case=False, na=False)]
            populate_tree(tree_widget, filtered_df)

        def reset_filter(dataframe, tree_widget, entry_widget):
            entry_widget.delete(0, 'end')
            populate_tree(tree_widget, dataframe)

        filter_button = ttk.Button(
            filter_frame,
            text="Filter",
            command=lambda d=df, t=tree, fc=filter_column, fe=filter_entry: apply_filter(d, t, fc, fe)
        )
        filter_button.pack(side="left", padx=5)
        
        reset_button = ttk.Button(
            filter_frame,
            text="Reset",
            command=lambda d=df, t=tree, fe=filter_entry: reset_filter(d, t, fe)
        )
        reset_button.pack(side="left", padx=5)

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
                for sheet_name, df_item in dfs.items():
                    df_item.to_excel(writer, index=False, sheet_name=sheet_name, na_rep="")
            workbook = openpyxl.load_workbook(save_path)
            auto_size_columns(workbook)
            workbook.save(save_path)
            open_file_externally(save_path)
            viewer.destroy()
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
