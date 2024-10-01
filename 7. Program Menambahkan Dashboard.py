#!/usr/bin/env python
# coding: utf-8

# In[4]:


import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import datetime
import os
# Mendapatkan tanggal dan bulan saat ini
today_date = datetime.datetime.today()
month = datetime.datetime.now().strftime("%Y-%m")  # Mendapatkan bulan saat ini dengan format YYYY-MM
tgl = datetime.datetime.now().strftime("%Y-%m-%d")  # Mendapatkan tanggal saat ini dengan format YYYY-MM-DD

# Membuat path direktori dengan tanggal
dir_path = rf'C:\Users\lenovo\Desktop\Gilland Ganesha\3. Data Hasil Cleaning\{month}\{tgl}'
# Nama file yang akan dibaca
file1 = f'{tgl}_P2K_data_hasil.xlsx'
file2 = f'{tgl}_P2R_data_hasil.xlsx'
file3 = f'{tgl}_RPL_data_hasil.xlsx'

# Membuat path penuh ke file
file_path1 = os.path.join(dir_path, file1)
file_path2 = os.path.join(dir_path, file2)
file_path3 = os.path.join(dir_path, file3)

# Membaca file Excel menjadi DataFrame
df1 = pd.read_excel(file_path1)
df2 = pd.read_excel(file_path2)
df3 = pd.read_excel(file_path3)

# Menggabungkan semua DataFrame
merged_df = pd.concat([df1, df2, df3], ignore_index=True)

# Menambahkan kolom urutan "NO"
merged_df["NO"] = range(1, len(merged_df) + 1)
dashboard_df = merged_df.copy()

# Mendefinisikan lingkup akses untuk Google Sheets dan Google Drive, serta melakukan otentikasi menggunakan file kredensial
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Credential\Faiqul cred.json', scope)
client = gspread.authorize(creds)

# Membuka Google Sheet menggunakan URL atau ID sheet
spreadsheet_url = 'https://docs.google.com/spreadsheets/d/1onJjQpA3lfBy7UCN8liF9xR1Rs2Ga89HFeKj3CifWMM/edit?gid=0#gid=0'
sheet = client.open_by_url(spreadsheet_url).sheet1  # Pilih sheet pertama atau bisa gunakan nama sheet spesifik

# Mengambil data yang sudah ada di sheet untuk menentukan nilai terakhir kolom "NO"
existing_data = sheet.get_all_values()

# Mengecek apakah ada data yang sudah ada (lebih dari header)
if len(existing_data) > 1:  # Jika ada lebih dari satu baris data
    # Diasumsikan baris pertama adalah header, dan kolom "NO" ada di kolom pertama
    last_no = int(existing_data[-1][0])  # Mengambil nilai terakhir dari kolom "NO" dari baris terakhir
else:
    # Jika belum ada data, nomor dimulai dari 0
    last_no = 0

# Mengganti nilai NaN di DataFrame dengan string kosong
dashboard_df = dashboard_df.fillna('')

# Menambahkan nomor baru di kolom "NO" mulai dari nomor setelah nomor terakhir
dashboard_df['NO'] = range(last_no + 1, last_no + 1 + len(dashboard_df))

# Memastikan semua baris memiliki jumlah kolom yang konsisten dengan struktur header di sheet
columns = ['NO', 'SUMBER DATA', 'TANGGAL INPUT', 'TANGGAL PROSES', 'TANGGAL UPLOAD', 'PICK UP', 
           'NAMA', 'NO TELEPON', 'EMAIL', 'DOMISILI', 'NAMA PERUSAHAAN', 
           'ALAMAT PERUSAHAAN', 'KOTA', 'KAMPUS', 'PRODI', 'EVENT']

# Mengurutkan DataFrame berdasarkan kolom yang sesuai dengan struktur kolom yang diinginkan
dashboard_df = dashboard_df.reindex(columns=columns)

# Mengubah semua kolom datetime menjadi format string (jika ada kolom datetime)
dashboard_df = dashboard_df.applymap(lambda x: x.strftime('%Y-%m-%d %H:%M:%S') if isinstance(x, pd.Timestamp) else x)

# Mengubah DataFrame menjadi list of lists (hanya mengambil nilai tanpa index)
data_to_append = dashboard_df.values.tolist()

# Mencari baris pertama yang kosong (tempat di mana data baru akan ditambahkan)
next_row = len(existing_data) + 1  # +1 untuk menambahkan data setelah baris terakhir

# Mendefinisikan sel awal untuk menambahkan data (mulai dari kolom A, baris kosong pertama)
start_cell = f"A{next_row}"

# Menambahkan data baru dimulai dari kolom A
sheet.update(start_cell, data_to_append, value_input_option='RAW')

print(f"Data berhasil ditambahkan ke Google Sheet dimulai dari {start_cell}.")


# In[ ]:




