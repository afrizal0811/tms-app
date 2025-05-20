import pandas as pd
import os
from tkinter import Tk, filedialog, messagebox
from datetime import datetime
import subprocess
import sys
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter
from geopy.distance import geodesic
from openpyxl.styles import Alignment

def convert_datetime_column(df, column_name, source_format=None, target_format='%H:%M'):
    def convert(val):
        if pd.isna(val):
            return ''
        try:
            if source_format:
                dt = datetime.strptime(val, source_format)
            elif 'T' in val:
                dt = datetime.fromisoformat(val.replace('Z', '+00:00'))
            else:
                dt = pd.to_datetime(val)
            return dt.strftime(target_format)
        except:
            return val
    return df[column_name].astype(str).apply(convert)

def insert_blank_rows(df, column):
    new_rows = []
    prev_value = None
    for _, row in df.iterrows():
        current_value = row[column]
        if prev_value is not None and current_value != prev_value:
            new_rows.append(pd.Series([None]*len(df.columns), index=df.columns))
        new_rows.append(row)
        prev_value = current_value
    return pd.DataFrame(new_rows)

def get_save_path():
    base_name = "Hasil RO vs Real"
    folder = filedialog.askdirectory(title="Pilih Lokasi Untuk Menyimpan File")
    if not folder:
        return None
    save_path = os.path.join(folder, base_name + ".xlsx")  # <- tambahkan ekstensi di sini
    counter = 1
    while os.path.exists(save_path):
        save_path = os.path.join(folder, f"{base_name} - {counter}.xlsx")
        counter += 1
    return save_path

def autosize_columns(filename, max_width=15):
    wb = load_workbook(filename)
    ws = wb.active
    for col in ws.columns:
        column = col[0].column
        ws.column_dimensions[get_column_letter(column)].width = max_width
    wb.save(filename)

def baca_file_data():
    try:
        data_df = pd.read_excel('Master_Driver.xlsx')
        data_df.columns = [col.strip() for col in data_df.columns]
        if 'Email' not in data_df.columns or 'Driver' not in data_df.columns:
            raise ValueError("Kolom 'Email' dan/atau 'Driver' tidak ditemukan di file 'Master Driver.xlsx'")
        data_df['Email'] = data_df['Email'].astype(str).str.strip().str.lower()
        data_df['Driver'] = data_df['Driver'].astype(str).str.strip()
        return data_df
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan saat membaca file 'Master Driver.xlsx': {e}")
        return None

def parse_coordinates(coord_str):
    try:
        lat, lon = map(float, str(coord_str).split(','))
        return (lat, lon)
    except:
        return None

def is_in_range(expected_coord, done_coord):
    expected_point = parse_coordinates(expected_coord)
    done_point = parse_coordinates(done_coord)
    if expected_point and done_point:
        distance = geodesic(expected_point, done_point).meters
        return round(distance, 1)
    return ""


def main():
    root = Tk()
    root.withdraw()
    messagebox.showinfo("Informasi", "Pilih Export Task (Custom 'RO vs Real')")

    try:
        file_path = filedialog.askopenfilename(title="Pilih File Excel", filetypes=[("Excel Files", "*.xlsx")])
        if not file_path:
            messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")
            return
        df = pd.read_excel(file_path)


        # Menggunakan fungsi baca_file_data() untuk membaca Master Driver.xlsx
        data_df = baca_file_data()
        if data_df is None:
            return

        email_to_name = dict(zip(data_df['Email'], data_df['Driver']))
        email_to_plat = dict(zip(data_df['Email'], data_df['Plat']))

        df['assignee_email'] = df['assignee']
        df['assignee'] = df['assignee_email'].map(email_to_name).fillna(df['assignee_email'])

        if 'assignedVehicle' in df.columns:
            df['assignedVehicle'] = df.apply(
                lambda row: email_to_plat.get(row['assignee_email'], row['assignedVehicle'])
                if pd.isna(row['assignedVehicle']) or str(row['assignedVehicle']).strip() == '-'
                else row['assignedVehicle'],
                axis=1
            )

        columns_to_drop = ['eta', 'etd', 'Alasan Tolakan', 'Alasan Batal', 'assignedVehicleId']
        df = df.drop(columns=[col for col in columns_to_drop if col in df.columns], errors='ignore')

        if 'Klik Jika Anda Sudah Sampai' in df.columns:
            df['Klik Jika Anda Sudah Sampai'] = convert_datetime_column(df, 'Klik Jika Anda Sudah Sampai')
        if 'doneTime' in df.columns:
            df['doneTime'] = convert_datetime_column(df, 'doneTime')

        if 'Klik Jika Anda Sudah Sampai' in df.columns and 'doneTime' in df.columns:
            time_format = "%H:%M"
            def calculate_actual_visit(start, end):
                try:
                    t1 = datetime.strptime(str(start), time_format)
                    t2 = datetime.strptime(str(end), time_format)
                    delta = (t2 - t1).total_seconds()
                    if delta < 0:
                        delta += 86400
                    return int(delta // 60)
                except:
                    return ""
            df['Actual Visit Time'] = df.apply(
                lambda row: calculate_actual_visit(row.get('Klik Jika Anda Sudah Sampai', ''), row.get('doneTime', '')),
                axis=1
            )
            if 'Visit Time' in df.columns:
                visit_time_index = df.columns.get_loc('Visit Time')
                cols = df.columns.tolist()
                cols.insert(visit_time_index + 1, cols.pop(cols.index('Actual Visit Time')))
                df = df[cols]

        if 'assignee' in df.columns:
            df = df.sort_values(by='assignee', ascending=True)

        if 'assignee' in df.columns:
            df = insert_blank_rows(df, 'assignee')

        if 'doneTime' in df.columns and 'assignee' in df.columns:
            df['doneTime_parsed'] = pd.to_datetime(df['doneTime'], format='%H:%M', errors='coerce')
            df['Real Seq'] = df.groupby('assignee')['doneTime_parsed'].rank(method='dense').astype('Int64')
            df.drop(columns=['doneTime_parsed'], inplace=True)
        else:
            df['Real Seq'] = pd.NA

        df.drop(columns=['assignee_email'], inplace=True)

        # Rename kolom sesuai kebutuhan
        rename_dict = {
            'assignedVehicle': 'License Plat',
            'assignee': 'Driver',
            'title': 'Customer',
            'label': 'Status Delivery',
            'Open Time': 'Open Time',
            'Close Time': 'Close Time',
            'Klik Jika Anda Sudah Sampai': 'Actual Time Arrived',
            'doneTime': 'Actual Time Depatured',
            'Visit Time': 'Planned Visit Time',
            'Actual Visit Time': 'Actual Visit Time',
            'routePlannedOrder': 'Planned Sequence',
            'Real Seq': 'Actual Sequence'
        }
        df.rename(columns=rename_dict, inplace=True)

        # Urutkan berdasarkan Driver dan Actual Sequence (di akhir proses)
        if 'Driver' in df.columns and 'Actual Sequence' in df.columns:
            df = df.sort_values(by=['Driver', 'Actual Sequence'], ascending=[True, True])

        # Sisipkan baris kosong antar grup Driver
        if 'Driver' in df.columns:
            df = insert_blank_rows(df, 'Driver')

        if 'expectedCoordinate' in df.columns and 'doneCoordinate' in df.columns:
            df['Actual vs Expected Distance (m)'] = df.apply(
                lambda row: is_in_range(row['expectedCoordinate'], row['doneCoordinate']),
                axis=1
            )

        # Tambahkan kolom 'Is Same Sequence' di akhir
        if 'Planned Sequence' in df.columns and 'Actual Sequence' in df.columns:
            df['Is Same Sequence'] = df.apply(
                lambda row: row['Planned Sequence'] == row['Actual Sequence']
                if pd.notna(row['Planned Sequence']) and pd.notna(row['Actual Sequence'])
                else pd.NA,
                axis=1
            )

        # Atur urutan kolom sesuai keinginan
        desired_columns = [
            'License Plat', 'Driver', 'Customer', 'Status Delivery',
            'Open Time', 'Close Time', 'Actual Time Arrived', 'Actual Time Depatured',
            'Planned Visit Time', 'Actual Visit Time',
            'Planned Sequence', 'Actual Sequence', 'Actual vs Expected Distance (m)',
            'Is Same Sequence'
        ]

        df = df[[col for col in desired_columns if col in df.columns]]

        # Simpan ke Excel
        save_path = get_save_path()
        if not save_path:
            messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")
            return
        df.to_excel(save_path, index=False)

        # Buka workbook dan worksheet
        wb = load_workbook(save_path)
        ws = wb.active

        # Kolom yang ingin dirata tengah
        center_columns = [
            'Status Delivery', 'Open Time', 'Close Time',
            'Actual Time Arrived', 'Actual Time Depatured',
            'Planned Visit Time', 'Actual Visit Time',
            'Planned Sequence', 'Actual Sequence'
        ]

        # Buat mapping kolom ke huruf kolom
        header_to_letter = {cell.value: cell.column_letter for cell in ws[1] if cell.value in center_columns}

        # Terapkan alignment center untuk kolom-kolom tersebut
        for col_letter in header_to_letter.values():
            for row in range(2, ws.max_row + 1):
                ws[f"{col_letter}{row}"].alignment = Alignment(horizontal="center", vertical="center")

        wb.save(save_path)

        autosize_columns(save_path)

        # Buka file otomatis setelah selesai
        if os.name == 'nt':
            os.startfile(save_path)
        elif os.name == 'posix':
            subprocess.call(['open' if sys.platform == 'darwin' else 'xdg-open', save_path])


    except Exception as e:
        messagebox.showerror("Terjadi Error", "Kesalahan pada proses: " + str(e))

if __name__ == "__main__":
    main()
