#!/usr/bin/env python
# coding: utf-8

# In[1]:


import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import requests
import os
import warnings as wr

# Mengabaikan peringatan
wr.filterwarnings('ignore')

# Mengatur opsi tampilan pandas (opsional)
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

# Informasi program
print("=================================================")
print("Program Mendownload File Dictionary dan Template BDC")
print("=================================================")

# Definisikan scope dan lakukan otentikasi menggunakan file kredensial untuk Google Sheets (dictionary)
scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
creds = ServiceAccountCredentials.from_json_keyfile_name(r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Credential\Faiqul cred.json', scope)
client = gspread.authorize(creds)

# Step 1: Download Dictionary from Google Sheets
url1 = 'https://docs.google.com/spreadsheets/d/1yQhIOExgbMXWj4dxu2lufSKv0eTKGAXj7SG66pjFzg4/export?format=xlsx'
output_path1 = r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\dictionary.xlsx'

# Mengecek apakah file sudah ada dan mencetak pesan jika file akan diganti
if os.path.exists(output_path1):
    print(f"File {output_path1} sudah ada dan akan diganti.")

# Mengirim permintaan GET untuk mengunduh file dari Google Sheets
response1 = requests.get(url1)

# Mengecek apakah permintaan berhasil
if response1.status_code == 200:
    # Menyimpan konten ke file lokal, menimpa jika sudah ada
    with open(output_path1, 'wb') as file:
        file.write(response1.content)
    print(f"File Dictionary berhasil diunduh dan disimpan di: {output_path1}")
else:
    print(f"Gagal mengunduh file Dictionary. Kode status: {response1.status_code}")

# Step 2: Download Template BDC from Web BDC
url_bdc = 'https://daftarkuliah.my.id/bdcv2/assets/excel/Template Data Prospek.xlsx'
output_path_bdc = r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\Template BDC Upload.xlsx'

# Mengecek apakah file sudah ada dan mencetak pesan jika file akan diganti
if os.path.exists(output_path_bdc):
    print(f"File {output_path_bdc} sudah ada dan akan diganti.")

# Mengirim permintaan GET untuk mengunduh file dari Web BDC
response_bdc = requests.get(url_bdc)

# Mengecek apakah permintaan berhasil
if response_bdc.status_code == 200:
    # Menyimpan konten ke file lokal, menimpa jika sudah ada
    with open(output_path_bdc, 'wb') as file:
        file.write(response_bdc.content)
    print(f"File Template BDC berhasil diunduh dan disimpan di: {output_path_bdc}")
else:
    print(f"Gagal mengunduh file Template BDC. Kode status: {response_bdc.status_code}")


# In[ ]:




