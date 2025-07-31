import sys
import os
import subprocess
import platform
from datetime import datetime
import tkinter as tk
from tkinter import messagebox, filedialog
import requests
import json
import threading
import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter

# --- Perbaikan untuk ModuleNotFoundError ---
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(project_root)
# --- Akhir Perbaikan ---

# --- Impor dari shared_utils dan gui_utils ---
from modules.shared_utils import (
    load_config,
    load_constants
)
from modules.gui_utils import create_date_picker_window

def get_save_directory():
    """Membuka dialog untuk memilih direktori penyimpanan."""
    root = tk.Tk()
    root.withdraw()
    folder_selected = filedialog.askdirectory(parent=root)
    root.destroy()
    return folder_selected

def get_unique_filepath(directory, basename, extension):
    """Mencari path file yang unik dengan menambahkan angka jika sudah ada."""
    counter = 1
    file_path = os.path.join(directory, f"{basename}{extension}")
    while os.path.exists(file_path):
        file_path = os.path.join(directory, f"{basename} ({counter}){extension}")
        counter += 1
    return file_path

def open_file(filepath):
    """Membuka file dengan aplikasi default sistem operasi."""
    try:
        if platform.system() == 'Darwin':       # macOS
            subprocess.call(('open', filepath))
        elif platform.system() == 'Windows':    # Windows
            os.startfile(filepath)
        else:                                   # linux variants
            subprocess.call(('xdg-open', filepath))
    except Exception as e:
        messagebox.showwarning("Gagal Membuka File", f"Tidak dapat membuka file secara otomatis:\n{e}")

def style_excel(file_path):
    """Menerapkan style pada file Excel yang sudah ada."""
    workbook = load_workbook(file_path)
    # Style semua sheet yang ada di workbook
    for sheet_name in workbook.sheetnames:
        ws = workbook[sheet_name]
        for col in ws.columns:
            max_length = 0
            column_letter = get_column_letter(col[0].column)
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = (max_length + 2)
            ws.column_dimensions[column_letter].width = adjusted_width

    # Terapkan style spesifik jika sheet "Truck Detail" ada
    if "Truck Detail" in workbook.sheetnames:
        ws_detail = workbook["Truck Detail"]
        center_align = Alignment(horizontal='center', vertical='center')
        for row in ws_detail.iter_rows(min_row=2, min_col=3, max_col=8):
            for cell in row:
                cell.alignment = center_align
    workbook.save(file_path)


def load_master_data(project_root, lokasi_code):
    """Memuat data master dan memfilternya berdasarkan kode lokasi."""
    master_path = os.path.join(project_root, 'master.json')
    try:
        with open(master_path, 'r', encoding='utf-8') as f:
            master_data = json.load(f)
            return {item['Email']: item['Driver'] for item in master_data if lokasi_code in item.get('Email', '')}
    except FileNotFoundError:
        messagebox.showerror("Error", "File master.json tidak ditemukan di direktori root.")
        return None
    except (json.JSONDecodeError, KeyError) as e:
        messagebox.showerror("Error", f"Gagal membaca atau memproses master.json: {e}")
        return None

def process_routing_data(date_formats, gui_instance):
    """
    Fungsi callback yang dipanggil oleh gui_utils untuk memproses data.
    """
    try:
        selected_date_dmy = date_formats['dmy']
        selected_date_for_filename = selected_date_dmy.replace('-', '.')
        
        try:
            date_obj = datetime.strptime(selected_date_dmy, '%d-%m-%Y')
            search_date_string = date_obj.strftime('%d/%m/%Y')
        except ValueError:
            messagebox.showerror("Error Format Tanggal", "Format tanggal dari pemilih tidak valid.")
            return
        
        config = load_config()
        constants = load_constants()
        lokasi_code = config.get('lokasi')
        
        master_data_map = load_master_data(project_root, lokasi_code)

        if not all([config, constants, master_data_map]):
            return

        api_token = constants.get('token')
        hub_id = constants.get('hub_ids', {}).get(lokasi_code)
        lokasi_mapping = constants.get('lokasi_mapping', {})
        lokasi_name = next((name for name, code in lokasi_mapping.items() if code == lokasi_code), lokasi_code)

        api_url = "https://apiweb.mile.app/api/v3/results"
        params = {'s': search_date_string, 'limit': 1000, 'hubId': hub_id}
        headers = {'Authorization': f'Bearer {api_token}'}

        gui_instance.update_status("Menghubungi API...")
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        response_data = response.json()
        
        processed_data = []
        first_tags_list = []
        processed_assignees_for_usage = set() # Set untuk melacak assignee yang sudah dihitung

        if 'data' in response_data and 'data' in response_data['data']:
            routing_results = response_data['data']['data']
            for item in routing_results:
                if 'result' in item and 'routing' in item['result']:
                    for route in item['result']['routing']:
                        assignee_email = route.get('assignee')
                        if not assignee_email or lokasi_code not in assignee_email:
                            continue

                        # --- LOGIKA BARU: HITUNG ASSIGNEE 1X UNTUK TRUCK USAGE ---
                        # Cek apakah assignee ini belum pernah dihitung untuk Truck Usage
                        if assignee_email not in processed_assignees_for_usage:
                            vehicle_tags = route.get('vehicleTags', [])
                            if vehicle_tags:
                                first_tags_list.append(vehicle_tags[0])
                            # Tandai assignee ini sudah dihitung
                            processed_assignees_for_usage.add(assignee_email)

                        # Data untuk Truck Detail tetap diproses untuk setiap rute agar bisa diakumulasi
                        driver_name = master_data_map.get(assignee_email, assignee_email)
                        try: weight_num = float(str(route.get('weightPercentage', '0')).replace('%', ''))
                        except: weight_num = 0.0
                        try: volume_num = float(str(route.get('volumePercentage', '0')).replace('%', ''))
                        except: volume_num = 0.0
                        try: time_num = int(route.get('totalSpentTime', 0))
                        except: time_num = 0
                        
                        processed_data.append({
                            'Assignee': driver_name, 'Vehicle Name': route.get('vehicleName'),
                            'Total Distance (m)': route.get('totalDistance', 0), 'Total Visits': None,
                            'Total Delivered': None, 'weight_numeric': weight_num,
                            'volume_numeric': volume_num, 'ship_duration_minutes': time_num
                        })

        if not processed_data:
            df_api_grouped = pd.DataFrame()
        else:
            df_api = pd.DataFrame(processed_data)
            agg_rules = {
                'Vehicle Name': lambda x: ', '.join(x.dropna().unique()), 'Total Distance (m)': 'sum',
                'Total Visits': 'first', 'Total Delivered': 'first', 'weight_numeric': 'sum',
                'volume_numeric': 'sum', 'ship_duration_minutes': 'sum'
            }
            df_api_grouped = df_api.groupby('Assignee', as_index=False).agg(agg_rules)
            df_api_grouped['Weight Percentage'] = df_api_grouped['weight_numeric'].apply(lambda x: f"{x:.1f}%")
            df_api_grouped['Volume Percentage'] = df_api_grouped['volume_numeric'].apply(lambda x: f"{x:.1f}%")
            df_api_grouped['Ship Duration'] = df_api_grouped['ship_duration_minutes'].apply(lambda x: f"'{divmod(x, 60)[0]:02d}:{divmod(x, 60)[1]:02d}")
            df_api_grouped = df_api_grouped.drop(columns=['weight_numeric', 'volume_numeric', 'ship_duration_minutes'])
        
        all_master_drivers = set(master_data_map.values())
        drivers_in_api = set(df_api_grouped['Assignee']) if not df_api_grouped.empty else set()
        missing_drivers = all_master_drivers - drivers_in_api
        df_missing = pd.DataFrame([{'Assignee': driver} for driver in missing_drivers])
        df_final = pd.concat([df_api_grouped, df_missing], ignore_index=True).sort_values(by='Assignee', ascending=True).reset_index(drop=True)
        column_order = ['Vehicle Name', 'Assignee', 'Weight Percentage', 'Volume Percentage', 'Total Distance (m)', 'Total Visits', 'Total Delivered', 'Ship Duration']
        df_final = df_final.reindex(columns=column_order)

        dry_dist_m = df_final[df_final['Assignee'].str.contains("DRY", na=False)]['Total Distance (m)'].sum()
        frz_dist_m = df_final[df_final['Assignee'].str.contains("FRZ", na=False)]['Total Distance (m)'].sum()
        df_summary = pd.DataFrame({'DRY': [round(dry_dist_m / 1000, 2)], 'FRZ': [round(frz_dist_m / 1000, 2)]})

        vehicle_types = ["L300", "CDE", "CDE-LONG", "CDD", "CDD-LONG", "FUSO", "FUSO-LONG"]
        usage_counts = {v_type: {'DRY': 0, 'FROZEN': 0} for v_type in vehicle_types}
        sorted_vehicle_types = sorted(vehicle_types, key=len, reverse=True)
        
        for tag in first_tags_list:
            if "KFC" in tag:
                if "FROZEN" in tag: usage_counts['CDD-LONG']['FROZEN'] += 1
                elif "DRY" in tag: usage_counts['CDD-LONG']['DRY'] += 1
                continue
            if "DRY-HAVI" in tag:
                usage_counts['FUSO']['DRY'] += 1
                continue
            
            category = None
            if "DRY" in tag: category = 'DRY'
            elif "FROZEN" in tag: category = 'FROZEN'
            
            if category:
                for v_type in sorted_vehicle_types:
                    if v_type in tag:
                        usage_counts[v_type][category] += 1
                        break
        
        usage_data_for_df = []
        for v_type, counts in usage_counts.items():
            dry_count = counts['DRY'] if counts['DRY'] > 0 else None
            frozen_count = counts['FROZEN'] if counts['FROZEN'] > 0 else None
            usage_data_for_df.append({
                'Tipe Kendaraan': v_type,
                'Jumlah (DRY)': dry_count,
                'Jumlah (FROZEN)': frozen_count
            })
        df_usage = pd.DataFrame(usage_data_for_df)

        gui_instance.update_status("Memilih direktori penyimpanan...")
        save_dir = get_save_directory()
        if not save_dir: return

        file_basename = f"Routing {lokasi_name} - {selected_date_for_filename}"
        save_path = get_unique_filepath(save_dir, file_basename, ".xlsx")

        gui_instance.update_status(f"Menyimpan ke {os.path.basename(save_path)}...")
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Truck Detail', index=False)
            df_summary.to_excel(writer, sheet_name='Total Distance Summary', index=False)
            df_usage.to_excel(writer, sheet_name='Truck Usage', index=False)
        
        gui_instance.update_status("Menerapkan style...")
        style_excel(save_path)
        
        gui_instance.update_status("Membuka file...")
        open_file(save_path)
        
        gui_instance.update_status("Selesai.")

    except Exception as e:
        import traceback
        messagebox.showerror("Error Tak Terduga", f"Terjadi kesalahan: {e}\n\n{traceback.format_exc()}")
    finally:
        if 'gui_instance' in locals() and gui_instance.winfo_exists():
             gui_instance.after(0, gui_instance.destroy)

def main():
    create_date_picker_window(
        title="Auto Routing Summary",
        process_callback=process_routing_data
    )

if __name__ == '__main__':
    main()