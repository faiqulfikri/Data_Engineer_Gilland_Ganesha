#!/usr/bin/env python
# coding: utf-8

# In[1]:


import os
import shutil
import datetime

# Mendapatkan tanggal hari ini dalam format 'YYYY-MM-DD'
today_str = datetime.datetime.now().strftime('%Y-%m-%d')
current_month_str = datetime.datetime.now().strftime('%Y-%m')

# Mendefinisikan path sumber secara dinamis berdasarkan bulan dan tanggal hari ini
source_folders = [
    fr'C:\Users\lenovo\Desktop\Gilland Ganesha\1. Data Hasil Tarik Perhari\{current_month_str}\{today_str}',
    fr'C:\Users\lenovo\Desktop\Gilland Ganesha\2. Data Hasil Tarik RAR\{current_month_str}',
    fr'C:\Users\lenovo\Desktop\Gilland Ganesha\3. Data Hasil Cleaning\{current_month_str}\{today_str}',
    fr'C:\Users\lenovo\Desktop\Gilland Ganesha\4. Data Upload BDC\{current_month_str}',
    fr'C:\Users\lenovo\Desktop\Gilland Ganesha\5. Data Dashboard\{current_month_str}'
]

# Mendefinisikan path tujuan di GDrive secara dinamis
destination_folders = [
    fr'G:\My Drive\Gilland Ganesha\1. Data Hasil Tarik Perhari\{current_month_str}\{today_str}',
    fr'G:\My Drive\Gilland Ganesha\2. Data Hasil Tarik RAR\{current_month_str}',
    fr'G:\My Drive\Gilland Ganesha\3. Data Hasil Cleaning\{current_month_str}\{today_str}',
    fr'G:\My Drive\Gilland Ganesha\4. Data Upload BDC\{current_month_str}',
    fr'G:\My Drive\Gilland Ganesha\5. Data Dashboard\{current_month_str}'
]

# Mendapatkan tanggal hari ini untuk pemeriksaan file
today = datetime.datetime.now().date()

def copy_files(src, dest):
    # Memastikan folder tujuan ada, jika tidak, buat folder tersebut
    if not os.path.exists(dest):
        os.makedirs(dest)

    # Menelusuri seluruh file dan subdirektori di dalam folder sumber
    for root, dirs, files in os.walk(src):
        # Membuat subdirektori yang sesuai di folder tujuan
        rel_path = os.path.relpath(root, src)  # Mendapatkan path relatif
        dest_path = os.path.join(dest, rel_path)  # Membuat path untuk tujuan

        # Membuat direktori di tujuan jika belum ada
        if not os.path.exists(dest_path):
            os.makedirs(dest_path)

        # Menyalin file ke folder tujuan, melewati jika file sudah ada
        for file in files:
            src_file = os.path.join(root, file)  # Path lengkap file sumber
            dest_file = os.path.join(dest_path, file)  # Path lengkap file tujuan

            # Mendapatkan waktu modifikasi terakhir file
            file_mtime = datetime.datetime.fromtimestamp(os.path.getmtime(src_file)).date()

            # Mengecek apakah file dimodifikasi atau dibuat hari ini
            if file_mtime == today:
                if os.path.exists(dest_file):
                    print(f"File sudah ada, melewati: {dest_file}")
                else:
                    shutil.copy2(src_file, dest_file)  # Menyalin file beserta metadata
                    print(f"File disalin dari {src_file} ke {dest_file}")
            else:
                print(f"File {src_file} tidak diubah hari ini, melewati.")

# Memanggil fungsi untuk setiap folder sumber dan folder tujuan
for src_folder, dest_folder in zip(source_folders, destination_folders):
    copy_files(src_folder, dest_folder)


# In[ ]:




