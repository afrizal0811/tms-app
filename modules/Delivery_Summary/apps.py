from datetime import datetime
from openpyxl.styles import Alignment, PatternFill
from openpyxl.utils import get_column_letter
from tkinter import filedialog
import os
import pandas as pd
import re

from utils.function import (
    get_save_path,
    load_config,
    load_constants,
    load_master_data,
    open_file_externally,
    show_error_message,
    show_info_message
)
from utils.messages import ERROR_MESSAGES, INFO_MESSAGES

# =============================================================================
# HELPER FUNCTIONS
# =============================================================================

def apply_styles_and_formatting(writer):
    workbook = writer.book
    center_align = Alignment(horizontal='center', vertical='center')
    left_align = Alignment(horizontal='left', vertical='center')
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    cols_to_center = [
        'Open Time', 'Close Time', 'ETA', 'ETD', 'Actual Arrival',
        'Actual Departure', 'Visit Time', 'Actual Visit Time',
        'Customer ID', 'ET Sequence', 'Real Sequence', 'Temperature',
        'Total Visit', 'Total Delivered', 'Status Delivery', 'Is Same Sequence'
    ]

    for sheet_name in workbook.sheetnames:
        worksheet = writer.sheets[sheet_name]
        header_map = {cell.value: cell.column for cell in worksheet[1]}
        for col_name, col_idx in header_map.items():
            col_letter = get_column_letter(col_idx)
            align = center_align if col_name in cols_to_center else left_align
            for cell in worksheet[col_letter]:
                cell.alignment = align
                if col_name == ' ':
                    cell.fill = red_fill
        if ' ' in header_map:
            sep_col_idx = header_map[' ']
            worksheet.cell(row=1, column=sep_col_idx).value = ""
        for column_cells in worksheet.columns:
            try:
                max_length = max(len(str(cell.value)) for cell in column_cells if cell.value is not None)
                worksheet.column_dimensions[get_column_letter(column_cells[0].column)].width = min(max_length + 2, 50)
            except ValueError:
                pass

def convert_datetime_column(df, column_name, target_format='%H:%M'):
    def convert(val):
        if pd.isna(val) or val == '': return ''
        try:
            if isinstance(val, datetime):
                dt = val
            elif 'T' in str(val):
                dt = datetime.fromisoformat(str(val).replace('Z', '+00:00'))
            else:
                dt = pd.to_datetime(val)
            return dt.strftime(target_format)
        except Exception:
            return val
    df[column_name] = df[column_name].apply(convert)
    return df

def calculate_actual_visit(start, end):
    if start == '' or end == '' or pd.isna(start) or pd.isna(end): 
        return ""
    try:
        t1 = datetime.strptime(str(start), "%H:%M")
        t2 = datetime.strptime(str(end), "%H:%M")
        delta = (t2 - t1).total_seconds()
        if delta < 0: delta += 86400
        return int(delta // 60)
    except (ValueError, TypeError):
        return ""

# =============================================================================
# DATAFRAME PROCESSING
# =============================================================================

def process_total_delivered(df, master_driver_df):
    master_summary = master_driver_df[['Driver', 'Plat']].drop_duplicates().rename(columns={'Plat': 'License Plat'})
    if 'assignee' in df.columns and 'label' in df.columns:
        email_to_name = dict(zip(master_driver_df['Email'], master_driver_df['Driver']))
        df_proc = df.copy()
        df_proc['Driver'] = df_proc['assignee'].str.lower().map(email_to_name)
        df_proc.dropna(subset=['Driver'], inplace=True)
        visit_counts = df_proc['Driver'].value_counts().reset_index().rename(columns={'index':'Driver','Driver':'Total Visit'})
        delivered_counts = df_proc[df_proc['label'].str.upper() == 'SUKSES']['Driver'].value_counts().reset_index().rename(columns={'index':'Driver','Driver':'Total Delivered'})
        final_df = master_summary.merge(visit_counts, on='Driver', how='left').merge(delivered_counts, on='Driver', how='left')
    else:
        final_df = master_summary.assign(**{'Total Visit': pd.NA, 'Total Delivered': pd.NA})
    final_df[['Total Visit','Total Delivered']] = final_df[['Total Visit','Total Delivered']].astype('Int64')
    return final_df[['License Plat','Driver','Total Visit','Total Delivered']].sort_values('Driver')

def process_ro_vs_real(df, master_driver_df):
    df_proc = df.copy()
    email_to_name = dict(zip(master_driver_df['Email'], master_driver_df['Driver']))
    email_to_plat = dict(zip(master_driver_df['Email'], master_driver_df['Plat']))
    df_proc['assignee_email'] = df_proc['assignee'].str.lower()
    df_proc['assignee'] = df_proc['assignee_email'].map(email_to_name).fillna(df_proc['assignee'])
    df_proc['assignedVehicle'] = df_proc.apply(
        lambda r: email_to_plat.get(r['assignee_email'], r['assignedVehicle']) if not r['assignedVehicle'] else r['assignedVehicle'], axis=1
    )
    for col in ['Klik Jika Anda Sudah Sampai','doneTime']:
        if col in df_proc.columns: convert_datetime_column(df_proc,col)
    df_proc['Actual Visit Time'] = df_proc.apply(lambda r: calculate_actual_visit(r.get('Klik Jika Anda Sudah Sampai',''), r.get('doneTime','')),axis=1)
    df_proc['doneTime_parsed'] = pd.to_datetime(df_proc['doneTime'], format='%H:%M', errors='coerce')
    df_proc['Real Seq'] = df_proc.groupby('assignee')['doneTime_parsed'].rank(method='dense').astype('Int64')
    df_proc.rename(columns={
        'assignedVehicle':'License Plat','assignee':'Driver','title':'Customer','label':'Status Delivery',
        'Klik Jika Anda Sudah Sampai':'Actual Arrival','doneTime':'Actual Departure',
        'routePlannedOrder':'ET Sequence','Real Seq':'Real Sequence'
    }, inplace=True)
    df_proc['Is Same Sequence'] = (pd.to_numeric(df_proc['ET Sequence'], errors='coerce') == pd.to_numeric(df_proc['Real Sequence'], errors='coerce')).map({True:'SAMA',False:'TIDAK SAMA'})
    cols = ['License Plat','Driver','Customer','Status Delivery','Open Time','Close Time','Actual Arrival','Actual Departure','Visit Time','Actual Visit Time','ET Sequence','Real Sequence','Is Same Sequence']
    df_final = df_proc[cols].sort_values(['Driver','Real Sequence'])
    parts = []
    for _,g in df_final.groupby('Driver'):
        parts.append(g)
        parts.append(pd.DataFrame([[None]*len(df_final.columns)],columns=df_final.columns))
    return pd.concat(parts[:-1],ignore_index=True)

def process_pending_so(df, master_driver_df):
    df_proc = df.copy()
    email_to_name = dict(zip(master_driver_df['Email'], master_driver_df['Driver']))
    df_proc['Driver'] = df_proc['assignee'].str.lower().map(email_to_name).fillna(df_proc['assignee'])
    status_to_filter = ['BATAL','PENDING','TERIMA SEBAGIAN']
    df_filtered = df_proc[df_proc['label'].isin(status_to_filter)].copy()
    if df_filtered.empty: return None
    for col in ['Klik Jika Anda Sudah Sampai','doneTime','eta','etd']:
        if col in df_filtered.columns: convert_datetime_column(df_filtered,col)
    df_filtered['Actual Visit Time'] = df_filtered.apply(lambda r: calculate_actual_visit(r.get('Klik Jika Anda Sudah Sampai',''), r.get('doneTime','')),axis=1)
    df_filtered['Customer ID'] = df_filtered['title'].apply(lambda t: t.split('-')[1].strip() if isinstance(t,str) and '-' in t else '')
    df_filtered['Temperature'] = df_filtered['Driver'].str.split(' ').str[0].str.replace("'","")
    def reason(row):
        return row.get('Alasan Batal','') if row['label'] in ['PENDING','BATAL'] else row.get('Alasan Tolakan','') if row['label']=='TERIMA SEBAGIAN' else ''
    df_filtered['Reason'] = df_filtered.apply(reason,axis=1)
    def assign_faktur(row):
        return (row['title'],'','') if row['label']=='BATAL' else ('',row['title'],'') if row['label']=='TERIMA SEBAGIAN' else ('','','') if row['label']=='PENDING' else ('','','')
    (df_filtered['Faktur Batal/ Tolakan SO'],df_filtered['Terkirim Sebagian'],df_filtered['Pending']) = zip(*df_filtered.apply(assign_faktur,axis=1))
    cols = ['assignedVehicle','Driver','Faktur Batal/ Tolakan SO','Terkirim Sebagian','Pending','Reason','Open Time','Close Time','eta','etd','Klik Jika Anda Sudah Sampai','doneTime','Visit Time','Actual Visit Time','Customer ID','routePlannedOrder','Temperature']
    df_final = df_filtered[cols].rename(columns={
        'assignedVehicle':'License Plat','eta':'ETA','etd':'ETD','Klik Jika Anda Sudah Sampai':'Actual Arrival','doneTime':'Actual Departure','routePlannedOrder':'ET Sequence'
    })
    reason_loc = df_final.columns.get_loc('Reason')
    df_final.insert(reason_loc+1,' ', '')
    return df_final.sort_values('Driver')

def process_update_longlat(df):
    if 'title' not in df.columns or 'Klik Lokasi Client' not in df.columns:
        return pd.DataFrame(columns=["Customer ID","Customer Name","Location ID","New Longlat"])
    data = []
    for _,row in df.iterrows():
        longlat = str(row['Klik Lokasi Client']).strip()
        if longlat in ['','-']: continue
        parts = [p.strip() for p in str(row['title']).split('-')]
        if len(parts)>=3:
            customer_name,customer_id,location_id = parts[0],parts[1],parts[-1]
        else:
            match = re.search(r'(C0\d+)', str(row['title']))
            customer_id = match.group(1) if match else ''
            customer_name = parts[0] if parts else ''
            location_id = parts[-1] if len(parts)>1 else ''
        data.append({"Customer ID":customer_id,"Customer Name":customer_name,"Location ID":location_id,"New Longlat":longlat})
    return pd.DataFrame(data,columns=["Customer ID","Customer Name","Location ID","New Longlat"])

# =============================================================================
# MAIN
# =============================================================================

def main():
    config = load_config()
    constants = load_constants()
    if not config or "lokasi" not in config:
        show_error_message("Dibatalkan", ERROR_MESSAGES["LOCATION_CODE_MISSING"]); return
    lokasi_code = config["lokasi"]
    show_info_message("Upload File Task", INFO_MESSAGES["SELECT_FILE"].format(text="export task"))
    input_file = filedialog.askopenfilename(title="Pilih File Excel yang Akan Diproses", filetypes=[("Excel Files","*.xlsx *.xls")])
    if not input_file:
        show_info_message("Dibatalkan", INFO_MESSAGES["CANCELLED_BY_USER"]); return
    df_original = pd.read_excel(input_file)
    required_columns = ['assignedVehicle','assignee','Alasan Tidak Bisa Dikunjungi','Alasan Batal','Open Time','Close Time','eta','etd','Klik Jika Anda Sudah Sampai','doneTime','Visit Time','routePlannedOrder']
    if any(col not in df_original.columns for col in required_columns):
        show_error_message("Proses Gagal", ERROR_MESSAGES["INVALID_FILE"].format(details="Upload file Export Task dengan benar!")); return
    email_prefixes = df_original["assignee"].dropna().astype(str).str.extract(r'kendaraan\.([^.@]+)',expand=False).dropna().str.lower().unique()
    if not any(lokasi_code.lower() in prefix for prefix in email_prefixes):
        show_error_message("Proses Gagal", ERROR_MESSAGES["LOCATION_CODE_MISSING"]); return
    master_df = load_master_data(lokasi_code)
    if master_df is None:
        show_error_message("Proses Gagal", ERROR_MESSAGES["MASTER_DATA_MISSING"]); return
    results_to_save = {
        'Total Delivered': process_total_delivered(df_original, master_df),
        'Hasil Pending SO': process_pending_so(df_original, master_df),
        'Hasil RO vs Real': process_ro_vs_real(df_original, master_df),
        'Update Longlat': process_update_longlat(df_original)
    }
    if results_to_save['Update Longlat'].empty:
        results_to_save['Update Longlat'] = pd.DataFrame([{"Customer ID":"Tidak Ada Update Longlat","Customer Name":"","Location ID":"","New Longlat":""}])
    lokasi_mapping = constants.get('lokasi_mapping', {})
    lokasi_name = next((n for n,c in lokasi_mapping.items() if c == lokasi_code), lokasi_code)
    input_filename = os.path.basename(input_file)
    date_match = re.search(r'(\d{2}-\d{2}-\d{4})', input_filename)
    date_str = date_match.group(1).replace('-', '.') if date_match else datetime.now().strftime('%d.%m.%Y')
    file_basename = f"Delivery Summary {lokasi_name} - {date_str}"
    save_file_path = get_save_path(file_basename)
    if not save_file_path: show_error_message("Proses Gagal", INFO_MESSAGES["CANCELLED_BY_USER"]); return
    with pd.ExcelWriter(save_file_path, engine='openpyxl') as writer:
        for sheet in ['Total Delivered','Hasil Pending SO','Hasil RO vs Real','Update Longlat']:
            if sheet in results_to_save and results_to_save[sheet] is not None:
                results_to_save[sheet].to_excel(writer, sheet_name=sheet, index=False)
        apply_styles_and_formatting(writer)
    open_file_externally(save_file_path)

if __name__ == "__main__":
    main()
