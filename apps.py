import pandas as pd
import os
import tkinter as tk
from tkinter import filedialog, messagebox
import subprocess
import re
from openpyxl.styles import PatternFill

def pilih_file_excel():
    root = tk.Tk()
    root.withdraw()
    messagebox.showinfo("Info", "Pilih Export Task (Custom 'Pending SO')")
    return filedialog.askopenfilename(title="Pilih file Excel", filetypes=[("Excel files", "*.xlsx *.xls")])

def generate_nama_file_simpan(folder):
    base_name = "Hasil Pending SO"
    ext = ".xlsx"
    file_path = os.path.join(folder, base_name + ext)
    counter = 1

    while os.path.exists(file_path):
        file_path = os.path.join(folder, f"{base_name} - {counter}{ext}")
        counter += 1

    return file_path

def buka_file(filepath):
    try:
        os.startfile(filepath)
    except AttributeError:
        subprocess.call(['open', filepath])
    except Exception:
        subprocess.call(['xdg-open', filepath])

def format_waktu(waktu):
    try:
        return pd.to_datetime(waktu).strftime('%H:%M')
    except:
        return None

def extract_customer_id(title):
    if pd.isna(title):
        return None
    match = re.search(r'(C0\d+)', str(title))
    return match.group(1) if match else None

def hitung_actual_visit_time(row):
    if pd.notna(row['doneTime']) and pd.notna(row['Klik Jika Anda Sudah Sampai']):
        try:
            done_time = pd.to_datetime(row['doneTime'], format='%H:%M')
            arrival_time = pd.to_datetime(row['Klik Jika Anda Sudah Sampai'], format='%H:%M')
            return (done_time - arrival_time).total_seconds() / 60
        except:
            return None
    return None

def baca_file_data():
    try:
        data_df = pd.read_excel('Master_Driver.xlsx')
        data_df.columns = [col.strip() for col in data_df.columns]
        if 'Email' not in data_df.columns or 'Driver' not in data_df.columns:
            raise ValueError("Kolom 'Email' dan/atau 'Driver' tidak ditemukan di file 'data.xlsx'")
        data_df['Email'] = data_df['Email'].astype(str).str.strip().str.lower()
        data_df['Driver'] = data_df['Driver'].astype(str).str.strip()
        return data_df
    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan saat membaca file 'data.xlsx': {e}")
        return None

def main():
    file_path = pilih_file_excel()
    if not file_path:
        messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")
        return

    try:
        data_df = baca_file_data()
        if data_df is None:
            return

        df = pd.read_excel(file_path)
        if 'assignee' not in df.columns:
            messagebox.showerror("Error", "Kolom 'assignee' tidak ditemukan.")
            return

        if 'Klik Jika Anda Sudah Sampai' in df.columns:
            df['Klik Jika Anda Sudah Sampai'] = df['Klik Jika Anda Sudah Sampai'].apply(format_waktu)
        if 'doneTime' in df.columns:
            df['doneTime'] = df['doneTime'].apply(format_waktu)
            df['doneTime_rank'] = pd.to_datetime(df['doneTime'], format='%H:%M', errors='coerce')

        df['Real Seq'] = df.groupby('assignee')['doneTime_rank'].rank(method='dense').astype('Int64')
        df.drop(columns=['doneTime_rank'], inplace=True)

        df['Customer ID'] = df['title'].apply(extract_customer_id) if 'title' in df.columns else None
        df['Actual Visit Time (minute)'] = df.apply(hitung_actual_visit_time, axis=1)

        df['Driver'] = df['assignee'].astype(str).str.strip().str.lower()
        email_to_name = dict(zip(data_df['Email'], data_df['Driver']))
        df['Driver'] = df['Driver'].map(email_to_name)

        df_comp = df[df['label'].isin(['PENDING', 'TERIMA SEBAGIAN', 'BATAL'])].copy()
        df_comp['Driver'] = df_comp['assignee'].astype(str).str.strip().str.lower()
        df_comp['Driver'] = df_comp['Driver'].map(email_to_name)

        df_comp['Customer Order'] = df_comp['title']
        df_comp['Status'] = df_comp['label']
        df_comp['Reason'] = df_comp['Alasan Batal'] if 'Alasan Batal' in df_comp.columns else None

        df_comp['Faktur Batal/ Tolakan SO'] = df_comp.apply(lambda x: x['Customer Order'] if x['Status'] == 'BATAL' else None, axis=1)
        df_comp['Terkirim Sebagian'] = df_comp.apply(lambda x: x['Customer Order'] if x['Status'] == 'TERIMA SEBAGIAN' else None, axis=1)
        df_comp['Pending'] = df_comp.apply(lambda x: x['Customer Order'] if x['Status'] == 'PENDING' else None, axis=1)

        tambahan_kolom = ['Open Time', 'Close Time', 'eta', 'etd',
                        'Klik Jika Anda Sudah Sampai', 'doneTime',
                        'Visit Time', 'Actual Visit Time (minute)',
                        'Customer ID', 'routePlannedOrder', 'Real Seq']

        for kolom in tambahan_kolom:
            if kolom not in df_comp.columns:
                df_comp[kolom] = None

        df_comp = df_comp.sort_values(by='Driver', ascending=False)

        kolom_final = [
            'Driver', 'Faktur Batal/ Tolakan SO', 'Terkirim Sebagian', 'Pending', 'Reason',
            'Open Time', 'Close Time', 'eta', 'etd',
            'Klik Jika Anda Sudah Sampai', 'doneTime',
            'Visit Time', 'Actual Visit Time (minute)',
            'Customer ID', 'routePlannedOrder', 'Real Seq'
        ]

        rename_kolom = {
            'eta': 'ETA',
            'etd': 'ETD',
            'Klik Jika Anda Sudah Sampai': 'Actual Arrival',
            'doneTime': 'Actual Departure',
            'routePlannedOrder': 'ET Sequence',
            'Real Seq': 'Real Sequence'
        }

        compilation_df = df_comp[kolom_final].rename(columns=rename_kolom)

        # Kolom kosong warna merah
        compilation_df.insert(compilation_df.columns.get_loc('Open Time'), '', '')

        # Kolom Assigned Vehicle
        if 'assignedVehicle' in df_comp.columns:
            compilation_df.insert(0, 'Assigned Vehicle', df_comp['assignedVehicle'])
        else:
            compilation_df.insert(0, 'Assigned Vehicle', '-')

        root = tk.Tk()
        root.withdraw()
        folder_path = filedialog.askdirectory(title="Pilih lokasi untuk menyimpan file")
        if not folder_path:
            messagebox.showwarning("Proses Gagal", "Proses Dibatalkan")
            return

        final_path = generate_nama_file_simpan(folder_path)

        with pd.ExcelWriter(final_path, engine='openpyxl') as writer:
            compilation_df.to_excel(writer, index=False, sheet_name='Compilation')
            workbook = writer.book
            sheet = workbook['Compilation']

            for col in sheet.columns:
                max_length = 0
                column = col[0].column_letter
                for cell in col:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                sheet.column_dimensions[column].width = (max_length + 2)

            red_fill = PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
            empty_column = sheet['G']  # Karena ada kolom baru di awal
            for cell in empty_column:
                cell.fill = red_fill

        buka_file(final_path)

    except Exception as e:
        messagebox.showerror("Error", f"Terjadi kesalahan: {e}")

if __name__ == '__main__':
    main()
