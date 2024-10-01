#!/usr/bin/env python
# coding: utf-8

# In[25]:


import pandas as pd
import datetime
import os
import shutil
import subprocess
import pyautogui
import pygetwindow as gw
import time

# Mendapatkan tanggal dan bulan saat ini
today_date = datetime.datetime.today()
month = datetime.datetime.now().strftime("%Y-%m")  # Mendapatkan bulan saat ini dengan format YYYY-MM
tgl = datetime.datetime.now().strftime("%Y-%m-%d")  # Mendapatkan tanggal saat ini dengan format YYYY-MM-DD

# Membuat path direktori dengan tanggal
dir_path = rf'C:\Users\lenovo\Desktop\Gilland Ganesha\3. Data Hasil Cleaning\{month}\{tgl}'

# Mendefinisikan path sumber file (template kosong)
source_path = r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\Template BDC Upload.xlsx'

# Membuat path untuk direktori tujuan (upload) dan file dengan format tanggal
destination_dir = rf'C:\Users\lenovo\Desktop\Gilland Ganesha\4. Data Upload BDC\{month}'
os.makedirs(destination_dir, exist_ok=True)

# Membuat nama file draft dan file template kosong
draft_file_path = os.path.join(destination_dir, f'Draft Template Data Prospek-{tgl}.xlsx')
template_file_path = os.path.join(destination_dir, f'Template Data Prospek-{tgl}.xlsx')

# Menyalin file template kosong
shutil.copy(source_path, template_file_path)
print(f'File template kosong disalin ke: {template_file_path}')

# Nama file yang akan dibaca
file1 = f'{tgl}_P2K_data_hasil.xlsx'
file2 = f'{tgl}_P2R_data_hasil.xlsx'
file3 = f'{tgl}_RPL_data_hasil.xlsx'

# Membuat path penuh ke file hasil cleaning
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

# Menambahkan kolom "ALAMAT PERUSAHAAN" dari "DOMISILI"
merged_df['ALAMAT PERUSAHAAN'] = merged_df['DOMISILI']

# Menyeleksi kolom yang ingin disimpan dalam draft
merged_df = merged_df[['NO', 'SUMBER DATA', 'PICK UP', 'PRODI', 'NO TELEPON', 'KOTA', 'NAMA PERUSAHAAN', 'ALAMAT PERUSAHAAN', 'NAMA', 'EMAIL', 'KAMPUS', 'TANGGAL INPUT', 'EVENT']]

# Menyimpan DataFrame ke file draft
merged_df.to_excel(draft_file_path, index=False)
print(f'Data draft disimpan ke: {draft_file_path}')

# Membuka kedua file (draft dan template kosong) secara otomatis
subprocess.Popen([r'explorer', draft_file_path])
time.sleep(2)  # Memberi waktu untuk membuka file

subprocess.Popen([r'explorer', template_file_path])
time.sleep(2)  # Memberi waktu untuk membuka file template

# Menemukan jendela Excel berdasarkan judul file
def find_window_by_title(title):
    windows = gw.getWindowsWithTitle(title)
    if len(windows) > 0:
        return windows[0]  # Mengembalikan jendela pertama yang cocok
    return None

# Fungsi untuk memindahkan jendela
def move_window(window, position):
    if window is not None:
        window.activate()  # Mengaktifkan jendela
        if position == "left":
            pyautogui.hotkey('win', 'left')  # Pindahkan ke kiri
        elif position == "right":
            pyautogui.hotkey('win', 'right')  # Pindahkan ke kanan
    else:
        print(f"Jendela dengan judul '{title}' tidak ditemukan.")

# Pindahkan jendela draft ke kiri dan template ke kanan
time.sleep(1)  # Beri waktu tambahan agar file benar-benar terbuka
draft_window = find_window_by_title(f"Draft Template Data Prospek-{tgl}.xlsx")
template_window = find_window_by_title(f"Template Data Prospek-{tgl}.xlsx")

# Memindahkan jendela ke kiri dan kanan
move_window(draft_window, "left")
time.sleep(1)  # Jeda
move_window(template_window, "right")

print("Jendela file sudah dibuka dan diatur berdampingan. Silakan salin data dari draft ke template secara manual.")

