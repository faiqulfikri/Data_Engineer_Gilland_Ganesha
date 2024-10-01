#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import shutil
import datetime
import zipfile

# Informasi program
print("=================================================")
print("Program Mengkopi file ke arsip data harian dan rar data harian")
print("=================================================")

# Definisikan folder-folder yang diperlukan
source_folder = r"C:\Users\lenovo\Desktop\Gilland Ganesha\New Files"
destination_folder = r"C:\Users\lenovo\Desktop\Gilland Ganesha\1. Data Hasil Tarik Perhari"
rar_folder = r"C:\Users\lenovo\Desktop\Gilland Ganesha\2. Data Hasil Tarik RAR"

# Mendapatkan tanggal hari ini untuk penamaan folder
today = datetime.datetime.today()
this_month = today.strftime('%Y-%m')  # Format bulan saat ini
this_date = today.strftime('%Y-%m-%d')  # Format tanggal saat ini

# Membuat folder tujuan untuk file yang disalin
destination_path = os.path.join(destination_folder, this_month, this_date)
os.makedirs(destination_path, exist_ok=True)  # Membuat folder jika belum ada

# Menyalin file CSV ke folder tujuan
for file_name in os.listdir(source_folder):
    if file_name.endswith('.csv'):  # Hanya memproses file dengan ekstensi .csv
        shutil.copy(os.path.join(source_folder, file_name), destination_path)

# Membuat file RAR/ZIP di direktori baru
rar_path = os.path.join(rar_folder, this_month)
os.makedirs(rar_path, exist_ok=True)  # Membuat folder untuk ZIP jika belum ada
zip_file_name = f"New Files {this_date}.zip"  # Nama file ZIP
zip_file_path = os.path.join(rar_path, zip_file_name)

# Membuat arsip ZIP untuk semua file yang telah disalin
with zipfile.ZipFile(zip_file_path, 'w') as zip_file:
    for file_name in os.listdir(destination_path):
        file_path = os.path.join(destination_path, file_name)
        zip_file.write(file_path, arcname=file_name)  # Tambahkan file ke ZIP

# Jika ingin menghapus file asli setelah ZIP dibuat, hilangkan komentar pada baris di bawah ini
# shutil.rmtree(destination_path)  # Hapus file asli jika sudah dimasukkan ke ZIP


# In[ ]:




