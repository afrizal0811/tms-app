# modules/Vehicles_Data/apps.py

import requests
import pandas as pd
import os
import sys
from tkinter import messagebox
from datetime import datetime
import openpyxl
from openpyxl.utils import get_column_letter

# Menambahkan project root ke sys.path agar impor shared_utils berfungsi
project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(project_root)

# Impor fungsi terpusat dari shared_utils
from modules.shared_utils import (
    load_config,
    load_constants,
    load_master_data,
    get_save_path,
    open_file_externally,
    load_secret 
)

def auto_size_columns(workbook):
    """Menyesuaikan lebar kolom agar sesuai dengan panjang teks maksimal."""
    for sheet_name in workbook.sheetnames:
        worksheet = workbook[sheet_name]
        for col in worksheet.columns:
            max_length = 0
            column = col[0].column_letter
            for cell in col:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except (ValueError, TypeError):
                    pass
            adjusted_width = (max_length + 2)
            if adjusted_width > 50:
                adjusted_width = 50
            if adjusted_width > 0:
                worksheet.column_dimensions[column].width = adjusted_width

def fetch_and_save_vehicles_data():
    """
    Mengambil data kendaraan dari API MileApp dan menyimpannya ke file Excel.
    """
    config = load_config()
    constants = load_constants()
    secrets = load_secret()

    if not config:
        messagebox.showerror("Error Konfigurasi", "File config.json tidak ditemukan atau kosong.")
        return False
    if not constants:
        messagebox.showerror("Error Konstanta", "File constant.json tidak ditemukan atau kosong.")
        return False
    if not secrets: 
        return False

    api_token = secrets.get('token')
    lokasi_code = config.get('lokasi')
    hub_ids = constants.get('hub_ids', {})
    lokasi_mapping = constants.get('lokasi_mapping', {})

    if not api_token:
        messagebox.showerror("Error Token API", "Token API tidak ditemukan.\n\nHubungi Admin.")
        return False
    if not lokasi_code:
        messagebox.showerror("Error Konfigurasi", "Lokasi cabang tidak diatur di config.json.")
        return False
    if lokasi_code not in hub_ids:
        messagebox.showerror("Error Hub ID", f"Hub ID untuk lokasi '{lokasi_code}' tidak ditemukan di constant.json.")
        return False

    hub_id = hub_ids.get(lokasi_code)
    api_url = "https://apiweb.mile.app/api/v3/vehicles"
    
    params = {
        "limit": 500,
        "hubId": hub_id
    }
    headers = {
        "Authorization": f"Bearer {api_token}",
        "Content-Type": "application/json"
    }

    try:
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        
        response_data = response.json()
        vehicles_data = response_data.get('data')

        if not vehicles_data:
            messagebox.showwarning("Data Kosong", "Tidak ada data kendaraan yang ditemukan dari API.")
            return False
        
        master_df = load_master_data(lokasi_cabang=lokasi_code)
        if master_df is None:
            return False
        driver_mapping = dict(zip(master_df['Email'].str.lower(), master_df['Driver']))
        
        template_data = []
        for vehicle in vehicles_data:
            working_time = vehicle.get('workingTime', {})
            break_time = vehicle.get('breakTime', {})
            capacity = vehicle.get('capacity', {})
            weight_cap = capacity.get('weight', {})
            volume_cap = capacity.get('volume', {})
            tags = vehicle.get('tags', [])

            template_data.append({
                "Name*": vehicle.get('name', ''),
                "Assignee": vehicle.get('assignee', ''),
                "Start Time": working_time.get('startTime', ''),
                "End Time": working_time.get('endTime', ''),
                "Break Start": break_time.get('startTime') if break_time.get('startTime') is not None else 0,
                "Break End": break_time.get('endTime') if break_time.get('endTime') is not None else 0,
                "Multiday": working_time.get('multiday') if working_time.get('multiday') is not None else 0,
                "Speed Km/h": vehicle.get('speed', 0),
                "Cost Factor": vehicle.get('fixedCost', 0),
                "Vehicle Tags": '; '.join(tags),
                "weight Min": weight_cap.get('min', ''),
                "weight Max": weight_cap.get('max', ''),
                "volume Min": volume_cap.get('min', ''),
                "volume Max": volume_cap.get('max', ''),
            })
            
        master_data = []
        for vehicle in vehicles_data:
            name = vehicle.get('name', '')
            assignee_email = (vehicle.get('assignee', '') or '').lower()
            tags = vehicle.get('tags', [])

            vehicle_type_raw = tags[0] if tags else ''
            
            if vehicle_type_raw == 'FROZEN-KFC':
                vehicle_type_raw = 'FROZEN-CDD-LONG-5000'
            elif vehicle_type_raw == 'DRY-HAVI':
                vehicle_type_raw = 'DRY-FUSO-LONG'

            driver_name = driver_mapping.get(assignee_email, assignee_email)
            
            master_data.append({
                "License Plat": name,
                "Type": vehicle_type_raw,
                "Email": assignee_email,
                "Name": driver_name,
            })
            
        df_template = pd.DataFrame(template_data)
        df_master = pd.DataFrame(master_data)
        
        # --- Logika untuk memindahkan data ke sheet Conditional Vehicle ---
        def clean_plat(plat_str):
            if not isinstance(plat_str, str): return ''
            parts = plat_str.split(' ')
            if len(parts) >= 3:
                return ' '.join(parts[:3])
            return plat_str

        df_master['base_plat'] = df_master['License Plat'].apply(clean_plat)
        
        duplicate_base_plats = df_master[df_master.duplicated(subset=['base_plat'], keep=False)]
        
        df_conditional = pd.DataFrame()
        if not duplicate_base_plats.empty:
            non_duplicate_master = df_master.drop(duplicate_base_plats.index)
            
            duplicate_groups = duplicate_base_plats.groupby('base_plat')
            
            master_rows_from_duplicates = []
            conditional_rows_from_duplicates = []
            
            for _, group in duplicate_groups:
                longest_plat_row = group.loc[group['License Plat'].str.len().idxmax()]
                
                conditional_rows_from_duplicates.append(longest_plat_row)
                
                shorter_plats_group = group.drop(longest_plat_row.name)
                if not shorter_plats_group.empty:
                    master_rows_from_duplicates.append(shorter_plats_group)
                    
            df_conditional = pd.DataFrame(conditional_rows_from_duplicates)
            
            if master_rows_from_duplicates:
                df_master = pd.concat([non_duplicate_master] + master_rows_from_duplicates, ignore_index=True)
            else:
                df_master = non_duplicate_master
        else:
            df_master = df_master.copy()
            
        # Hapus kolom helper
        df_master.drop(columns=['base_plat'], inplace=True)
        if not df_conditional.empty:
            df_conditional.drop(columns=['base_plat'], inplace=True)
            
        # --- Sinkronisasi nilai Type pada conditional dengan master ---
        if not df_conditional.empty and not df_master.empty:
            master_type_map = dict(zip(df_master['Email'], df_master['Type']))
            df_conditional['Type'] = df_conditional['Email'].map(master_type_map).fillna(df_conditional['Type'])
        
        # --- Sorting akhir ---
        df_master = df_master.sort_values(by="Email", ascending=True).reset_index(drop=True)
        if not df_conditional.empty:
            df_conditional = df_conditional.sort_values(by="Email", ascending=True).reset_index(drop=True)

        # --- Simpan ke file Excel ---
        lokasi_name = next((name for name, code in lokasi_mapping.items() if code == lokasi_code), lokasi_code)
        date_str = datetime.now().strftime("%d.%m.%Y")
        file_basename = f"Vehicle Data {lokasi_name} - {date_str}"
        save_path = get_save_path(file_basename)
        
        if not save_path:
            messagebox.showwarning("Dibatalkan", "Penyimpanan file dibatalkan.")
            return False

        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df_master.to_excel(writer, index=False, sheet_name='Master Vehicle')
            if not df_conditional.empty:
                df_conditional.to_excel(writer, index=False, sheet_name='Conditional Vehicle')
            df_template.to_excel(writer, index=False, sheet_name='Template Vehicle')
            
        workbook = openpyxl.load_workbook(save_path)
        auto_size_columns(workbook)
        workbook.save(save_path)
        
        open_file_externally(save_path)
        
        return True

    except requests.exceptions.HTTPError as errh:
        status_code = errh.response.status_code
        if status_code == 401:
            messagebox.showerror("Akses Ditolak (401)", "KESALAHAN: Unauthorized. Token API mungkin salah atau sudah kedaluwarsa.")
        elif status_code >= 500:
            messagebox.showerror("Masalah Server API", f"KESALAHAN: Terjadi masalah pada server API (Status Code: {status_code}). Coba lagi nanti.")
        else:
            messagebox.showerror("Kesalahan HTTP", f"KESALAHAN HTTP: {errh}")
        return False
    except requests.exceptions.ConnectionError:
        messagebox.showerror("Koneksi Gagal", "KESALAHAN: Tidak dapat terhubung ke server. Periksa koneksi internet Anda.")
        return False
    except requests.exceptions.Timeout:
        messagebox.showerror("Timeout", "Permintaan API melebihi batas waktu.")
        return False
    except requests.exceptions.RequestException as e:
        messagebox.showerror("Kesalahan API", f"Terjadi kesalahan saat melakukan permintaan API:\n{e}")
        return False
    except Exception as e:
        messagebox.showerror("Error Tak Terduga", f"Terjadi kesalahan tak terduga:\n{e}")
        return False

def main():
    """Fungsi utama untuk modul Vehicles Data."""
    fetch_and_save_vehicles_data()

if __name__ == "__main__":
    main()