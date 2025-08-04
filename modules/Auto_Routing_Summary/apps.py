from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Alignment
from openpyxl.utils import get_column_letter
import os
import pandas as pd
import requests
import sys
from utils.function import (
    get_save_path,
    load_config,
    load_constants,
    load_master_data as shared_load_master_data,
    load_secret,
    open_file_externally,
    show_error_message,
    show_info_message
)
from utils.gui import create_date_picker_window
from utils.messages import ERROR_MESSAGES, INFO_MESSAGES

project_root = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))
sys.path.append(project_root)

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
            show_error_message("Error Format Tanggal", "Format tanggal dari pemilih tidak valid.")
            return
        
        config = load_config()
        constants = load_constants()
        secrets = load_secret()
        lokasi_code = config.get('lokasi')
        
        master_data_df = shared_load_master_data(lokasi_cabang=lokasi_code)

        # Periksa semua konfigurasi yang diperlukan
        if not config:
            show_error_message("Gagal", ERROR_MESSAGES["CONFIG_FILE_ERROR"])
            return
        if not constants:
            show_error_message("Gagal", ERROR_MESSAGES["CONSTANT_FILE_ERROR"])
            return
        if not secrets:
            show_error_message("Gagal", ERROR_MESSAGES["SECRET_FILE_ERROR"])
            return
        if master_data_df is None or master_data_df.empty:
            show_error_message("Gagal", ERROR_MESSAGES["MASTER_DATA_MISSING"])
            return

        master_data_map = dict(zip(master_data_df['Email'], master_data_df['Driver']))
        
        api_token = secrets.get('token')
        hub_id = constants.get('hub_ids', {}).get(lokasi_code)
        lokasi_mapping = constants.get('lokasi_mapping', {})
        lokasi_name = next((name for name, code in lokasi_mapping.items() if code == lokasi_code), lokasi_code)

        if not api_token:
            show_error_message("Error Token API", ERROR_MESSAGES["API_TOKEN_MISSING"])
            return
        if not hub_id:
            show_error_message("Konfigurasi Salah", ERROR_MESSAGES["HUB_ID_MISSING"])
            return

        api_url = "https://apiweb.mile.app/api/v3/results"
        params = {'s': search_date_string, 'limit': 1000, 'hubId': hub_id}
        headers = {'Authorization': f'Bearer {api_token}'}

        gui_instance.update_status("Menghubungi API...")
        response = requests.get(api_url, headers=headers, params=params, timeout=30)
        response.raise_for_status()
        response_data = response.json()
        
        processed_data = []
        first_tags_list = []
        processed_assignees_for_usage = set() 

        if 'data' in response_data and 'data' in response_data['data']:
            routing_results = response_data['data']['data']
            for item in routing_results:
                if 'result' in item and 'routing' in item['result']:
                    for route in item['result']['routing']:
                        assignee_email = route.get('assignee')
                        if not assignee_email or lokasi_code not in assignee_email:
                            continue

                        if assignee_email not in processed_assignees_for_usage:
                            vehicle_tags = route.get('vehicleTags', [])
                            if vehicle_tags:
                                first_tags_list.append(vehicle_tags[0])
                            processed_assignees_for_usage.add(assignee_email)

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
            show_error_message("Data Tidak Ditemukan", ERROR_MESSAGES["DATA_NOT_FOUND"])
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
        
        file_basename = f"Routing Summary {lokasi_name} - {selected_date_for_filename}"
        save_path = get_save_path(base_name=file_basename, extension=".xlsx")
        
        if not save_path: 
            show_info_message("Dibatalkan", INFO_MESSAGES["CANCELLED_BY_USER"])
            return

        gui_instance.update_status(f"Menyimpan ke {os.path.basename(save_path)}...")
        with pd.ExcelWriter(save_path, engine='openpyxl') as writer:
            df_final.to_excel(writer, sheet_name='Truck Detail', index=False)
            df_summary.to_excel(writer, sheet_name='Total Distance Summary', index=False)
            df_usage.to_excel(writer, sheet_name='Truck Usage', index=False)
        
        gui_instance.update_status("Menerapkan style...")
        style_excel(save_path)
        
        gui_instance.update_status("Membuka file...")
        open_file_externally(save_path)
        
        gui_instance.update_status("Selesai.")

    except requests.exceptions.HTTPError as errh:
        status_code = errh.response.status_code
        if status_code == 401:
            show_error_message("Akses Ditolak (401)", ERROR_MESSAGES["API_TOKEN_MISSING"])
        elif status_code >= 500:
            show_error_message("Masalah Server API", ERROR_MESSAGES["SERVER_ERROR"].format(error_detail=status_code))
        else:
            show_error_message("Kesalahan HTTP", ERROR_MESSAGES["HTTP_ERROR_GENERIC"].format(status_code=status_code))
    except requests.exceptions.ConnectionError:
        show_error_message("Koneksi Gagal", ERROR_MESSAGES["CONNECTION_ERROR"].format(error_detail="Tidak dapat terhubung ke server. Periksa koneksi internet Anda."))
    except requests.exceptions.RequestException as e:
        show_error_message("Kesalahan API", ERROR_MESSAGES["API_REQUEST_FAILED"].format(error_detail=e))
    except Exception as e:
        import traceback
        show_error_message("Error Tak Terduga", ERROR_MESSAGES["UNKNOWN_ERROR"].format(error_detail=f"Terjadi kesalahan: {e}\n\n{traceback.format_exc()}"))


def main():
    """Fungsi utama untuk modul Auto Routing Summary."""
    
    def process_wrapper(dates, gui_instance):
        """Wrapper untuk menjalankan proses utama dan menangani penutupan GUI dengan aman."""
        
        def safe_close():
            if gui_instance and gui_instance.winfo_exists():
                gui_instance.destroy()

        try:
            process_routing_data(dates, gui_instance)
        finally:
            if gui_instance and gui_instance.winfo_exists():
                gui_instance.after(100, safe_close)
                
    create_date_picker_window(
        title="Auto Routing Summary",
        process_callback=process_wrapper
    )

if __name__ == '__main__':
    main()