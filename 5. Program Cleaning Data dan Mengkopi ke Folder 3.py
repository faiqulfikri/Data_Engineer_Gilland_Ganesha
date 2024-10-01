#!/usr/bin/env python
# coding: utf-8

# In[8]:


import gspread
from oauth2client.service_account import ServiceAccountCredentials
import pandas as pd
import glob
import chardet
import pandas as pd
import re
import pytz
import os
import shutil
import requests
from datetime import datetime
import unicodedata
from unidecode import unidecode
import xml.etree.ElementTree as ET
import warnings as wr
wr.filterwarnings('ignore')
pd.set_option('display.max_rows', 500)
pd.set_option('display.max_columns', 500)
pd.set_option('display.width', 1000)

# Tentukan path folder sumber
source_folder = r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files'
# Daftar kategori dan subfolder yang sesuai
categories = {
    'Kita Kuliah': r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\Kita Kuliah',
    'Affi'    : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\Affiliator',
    'RPL'     : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\RPL',
    'Kar&Reg' : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R&P2K',
    'Reg'     : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R',
    'Kar'     : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2K',
    'P2R&P2K' : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R&P2K',
    'P2R'     : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R',
    'P2K'     : r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2K'
}
# Buat subfolder jika belum ada
for subfolder in categories.values():
    if not os.path.exists(subfolder):
        os.makedirs(subfolder)
# Dapatkan daftar file dari direktori dengan ekstensi yang sesuai
def get_files_from_directory(directory, extensions=['*.csv', '*.xml', '*.xls']):
    file_paths = []
    for ext in extensions:
        file_paths.extend(glob.glob(os.path.join(directory, ext)))
    return file_paths
# Fungsi untuk mendeteksi encoding file
def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        print(f"Detected encoding for {file_path}: {encoding}")
    return encoding
# Fungsi untuk mendeteksi delimiter file CSV
def detect_delimiter(file_path, encoding):
    with open(file_path, 'r', encoding=encoding) as file:
        first_line = file.readline()
        if '\t' in first_line:
            return '\t'
        elif ',' in first_line:
            return ','
        else:
            return ','  # Default delimiter
# Fungsi untuk membaca CSV dengan encoding yang terdeteksi
def read_csv_with_detected_encoding(file_path):
    encoding = detect_encoding(file_path)
    delimiter = detect_delimiter(file_path, encoding)
    df = pd.read_csv(file_path, encoding=encoding, sep=delimiter)
    return df
# Fungsi untuk memproses DataFrame dengan satu kolom
def process_single_column_df(df, delimiter=','):
    if df.shape[1] == 1:
        split_columns = df.iloc[:, 0].str.split(delimiter, expand=True)
        return split_columns
    return df
# Fungsi untuk mengubah XML ke DataFrame
def parse_xml_to_df(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    namespaces = {'ss': 'urn:schemas-microsoft-com:office:spreadsheet'}
    data = []
    columns = []
    worksheet = root.find('.//ss:Worksheet', namespaces)
    if worksheet is not None:
        table = worksheet.find('ss:Table', namespaces)
        if table is not None:
            for i, row in enumerate(table.findall('ss:Row', namespaces)):
                row_data = []
                for cell in row.findall('ss:Cell', namespaces):
                    data_elem = cell.find('ss:Data', namespaces)
                    if data_elem is not None:
                        if i == 0:
                            columns.append(data_elem.text)
                        else:
                            row_data.append(data_elem.text)
                    else:
                        row_data.append('')
                if i > 0:
                    data.append(row_data)
    df = pd.DataFrame(data, columns=columns)
    return df
# Fungsi untuk mengubah file XLS ke XML
def rename_xls_to_xml(directory):
    xls_files = glob.glob(os.path.join(directory, '*.xls'))
    for file_path in xls_files:
        new_file_path = file_path.replace('.xls', '.xml')
        os.rename(file_path, new_file_path)
        print(f"Renamed {file_path} to {new_file_path}")
# Fungsi untuk mengubah file XML ke XLS
def rename_xml_to_xls(directory):
    xml_files = glob.glob(os.path.join(directory, '*.xml'))
    for file_path in xml_files:
        new_file_path = file_path.replace('.xml', '.xls')
        os.rename(file_path, new_file_path)
        print(f"Renamed {file_path} to {new_file_path}")
# Dapatkan daftar file dari direktori
files = get_files_from_directory(source_folder)
# Fungsi untuk memindahkan file berdasarkan kata kunci
def move_files_based_on_keywords():
    for file_path in files:
        filename = os.path.basename(file_path)
        # Abaikan jika itu adalah folder
        if os.path.isdir(file_path):
            continue
        # Tentukan ke kategori mana file tersebut sesuai berdasarkan nama file
        moved = False
        if "tanpa judul" in filename.lower():
            print(f"Skipping file: {filename}")
            continue
        else:
            for category, subfolder in categories.items():
                if category.lower() in filename.lower():
                    destination_file = os.path.join(subfolder, filename)
                    shutil.move(file_path, destination_file)
                    print(f"Moved {filename} to {subfolder}")
                    moved = True
                    break
        if moved:
            continue
        # Jika tidak ada kata kunci dalam nama file, periksa header file
        try:
            if filename.lower().endswith('.csv'):
                df = read_csv_with_detected_encoding(file_path)
            elif filename.lower().endswith('.xml'):
                df = parse_xml_to_df(file_path)
            elif filename.lower().endswith('.xls'):
                rename_xls_to_xml(source_folder)
                new_file_path = file_path.replace('.xls', '.xml')
                df = parse_xml_to_df(new_file_path)
                rename_xml_to_xls(source_folder)
            else:
                continue

            headers = df.columns.str.lower()
            if any("tanpa judul" in header for header in headers):
                print(f"Skipping file: {filename}")
                continue
            elif any("Untitled" in header for header in headers):
                print(f"Skipping file: {filename}")
                continue
            elif any("kita kuliah" in header for header in headers):
                destination_file = os.path.join(categories['Kita Kuliah'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['Kita Kuliah']}")
            elif any("rpl" in header for header in headers):
                destination_file = os.path.join(categories['RPL'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['RPL']}")
            elif any("affi" in header for header in headers):
                destination_file = os.path.join(categories['Affiliator'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['Affiliator']}")
            elif any("Kar&Reg" in header for header in headers):
                destination_file = os.path.join(categories['P2R&P2K'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2R&P2K']}")
            elif any("Kar" in header for header in headers):
                destination_file = os.path.join(categories['P2K'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2K']}")
            elif any("Reg" in header for header in headers):
                destination_file = os.path.join(categories['P2R'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2R']}")
            elif any("P2R&P2K" in header for header in headers):
                destination_file = os.path.join(categories['P2R&P2K'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2R&P2K']}")
            elif any("P2K" in header for header in headers):
                destination_file = os.path.join(categories['P2K'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2K']}")
            elif any("P2R" in header for header in headers):
                destination_file = os.path.join(categories['P2R'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2R']}")
            else:
                destination_file = os.path.join(categories['P2K'], filename)
                shutil.move(file_path, destination_file)
                print(f"Moved {filename} to {categories['P2K']}")
        except Exception as e:
            print(f"Failed to process file {filename}: {e}")
# Eksekusi fungsi
move_files_based_on_keywords()
#============================================================================================================================================================#
#============================================================================================================================================================#
# List of directories kuliah to process
directories_kuliah = [
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R&P2K',
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R',
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2K',
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\Kita Kuliah',
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\RPL'
]
directories_affiliator = [
    r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\Affiliator'
]
directory_output = r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files'
#============================================================================================================================================================#
#============================================================================================================================================================#
def get_files_from_directory(directory, extensions=['*.csv', '*.xml', '*.xls']):
    file_paths = []
    for ext in extensions:
        file_paths.extend(glob.glob(os.path.join(directory, ext)))
    return file_paths
def detect_encoding(file_path):
    with open(file_path, 'rb') as file:
        raw_data = file.read()
        result = chardet.detect(raw_data)
        encoding = result['encoding']
        print(f"Detected encoding for {file_path}: {encoding}")
    return encoding
def detect_delimiter(file_path, encoding):
    with open(file_path, 'r', encoding=encoding) as file:
        first_line = file.readline()
        if '\t' in first_line:
            return '\t'
        elif ',' in first_line:
            return ','
        else:
            return ','  # Default delimiter
def read_csv_with_detected_encoding(file_path):
    encoding = detect_encoding(file_path)
    delimiter = detect_delimiter(file_path, encoding)
    df = pd.read_csv(file_path, encoding=encoding, sep=delimiter)
    return df
def process_single_column_df(df, delimiter=','):
    if df.shape[1] == 1:
        split_columns = df.iloc[:, 0].str.split(delimiter, expand=True)
        return split_columns
    return df
def parse_xml_to_df(file_path):
    tree = ET.parse(file_path)
    root = tree.getroot()
    namespaces = {
        'ss': 'urn:schemas-microsoft-com:office:spreadsheet'
    }
    data = []
    columns = []
    worksheet = root.find('.//ss:Worksheet', namespaces)
    if worksheet is not None:
        table = worksheet.find('ss:Table', namespaces)
        if table is not None:
            for i, row in enumerate(table.findall('ss:Row', namespaces)):
                row_data = []
                for cell in row.findall('ss:Cell', namespaces):
                    data_elem = cell.find('ss:Data', namespaces)
                    if data_elem is not None:
                        if i == 0:
                            columns.append(data_elem.text)
                        else:
                            row_data.append(data_elem.text)
                    else:
                        row_data.append('')
                if i > 0:
                    data.append(row_data)

    df = pd.DataFrame(data, columns=columns)
    return df
def rename_xls_to_xml(directory):
    xls_files = glob.glob(os.path.join(directory, '*.xls'))
    for file_path in xls_files:
        new_file_path = file_path.replace('.xls', '.xml')
        os.rename(file_path, new_file_path)
        print(f"Renamed {file_path} to {new_file_path}")
def rename_xml_to_xls(directory):
    xml_files = glob.glob(os.path.join(directory, '*.xml'))
    for file_path in xml_files:
        new_file_path = file_path.replace('.xml', '.xls')
        os.rename(file_path, new_file_path)
        print(f"Renamed {file_path} to {new_file_path}")
# Function to convert datetime string to target time zone format
def format_datetime(value):
    try:
        # Convert ISO 8601 format to 'YYYY-MM-DD HH:MM:SS'
        return pd.to_datetime(value).strftime('%Y-%m-%d %H:%M:%S')
    except (ValueError, TypeError):
        return value  # If it cannot be parsed, return as is
#============================================================================================================================================================#
#============================================================================================================================================================#
def process_directory_kuliah(directory):
    print(f"Processing directory: {directory}")
    # Rename all XLS files to XML
    rename_xls_to_xml(directory)
    # Get file paths after renaming
    file_paths = get_files_from_directory(directory)
    # Store DataFrames in a dictionary
    dataframes = {}
    for idx, file_path in enumerate(file_paths, start=1):
        try:
            df1_name = f"df1{idx}"
            if file_path.endswith('.csv'):
                df1 = read_csv_with_detected_encoding(file_path)
                df1 = process_single_column_df(df1)
            elif file_path.endswith('.xml'):
                df1 = parse_xml_to_df(file_path)
            dataframes[df1_name] = df1
            print(f"Data from {file_path} has been loaded into {df1_name}")
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
    # Load the dictionary from Excel
    dict_df1 = pd.read_excel(r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\dictionary.xlsx')
    # Convert the dictionary DataFrame to a dict of dicts
    header_dict = {}
    for column in dict_df1.columns:
        if column.endswith('_dict'):
            key = column.replace('_dict', '')
            header_dict[key] = dict(zip(dict_df1[f'{key}_input'].dropna(), dict_df1[column].dropna()))
    # Standard column names for the merged DataFrame
    standard_columns_1 = ['NO', 'SUMBER DATA', 'TANGGAL INPUT', 'TANGGAL PROSES', 'TANGGAL UPLOAD', 'PICK UP', 'NAMA', 'NO TELEPON', 'EMAIL', 'DOMISILI', 'NAMA PERUSAHAAN', 'ALAMAT PERUSAHAAN', 'KOTA', 'KAMPUS', 'PRODI', 'EVENT']
    # Create an empty DataFrame to store the merged data
    merged_df1 = pd.DataFrame(columns=standard_columns_1)
    # Function to rename columns using the dictionary
    def rename_columns(df1, header_dict):
        temp_df1 = pd.DataFrame(columns=standard_columns_1)
        for key, mapping in header_dict.items():
            for input_col, dict_col in mapping.items():
                if input_col in df1.columns:
                    temp_df1[dict_col] = df1[input_col]
                    break
        return temp_df1
    # Loop through each DataFrame in the dictionary
    for key, df1 in dataframes.items():
        temp_df1 = rename_columns(df1, header_dict)
        merged_df1 = pd.concat([merged_df1, temp_df1], ignore_index=True)
    if merged_df1.empty:
        print("data kosong")
    else:
        merged_df1= merged_df1.replace(r'[\n\t\r]+', ' ', regex=True)
        # Further processing and saving the merged DataFrame
        merged_df1 = merged_df1.applymap(str)
        # Replace source data values
        merged_df1['SUMBER DATA'] = merged_df1['SUMBER DATA'].str.replace('fb', 'Facebook ADS')
        merged_df1['SUMBER DATA'] = merged_df1['SUMBER DATA'].str.replace('ig', 'Instagram ADS')
        merged_df1['PICK UP'] = merged_df1['PICK UP'].replace('', pd.NA).replace('nan', pd.NA, regex=True)
        # Process phone numbers
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].astype(str)
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].apply(lambda x: re.sub(r'[^\x00-\x7F]+', '', x))
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].str.replace(' ', '')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].str.replace('-', '')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].str.lstrip('p:')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].str.lstrip('62')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].str.lstrip('+62')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].str.lstrip('0')
        merged_df1['NO TELEPON'] = '0' + merged_df1['NO TELEPON']
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].apply(lambda x: x if len(x) >= 9 else '')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].apply(lambda x: x if len(x) <= 13 else '')
        merged_df1['NO TELEPON'] = merged_df1['NO TELEPON'].apply(lambda x: x if str(x).isdigit() else '')
        # Process city names
        merged_df1['KOTA'] = merged_df1['KOTA'].str.title()
        # Process company names
        merged_df1['NAMA PERUSAHAAN'] = merged_df1['NAMA PERUSAHAAN'].astype(str).str.strip().replace('nan', '', regex=True)
        merged_df1['NAMA PERUSAHAAN'] = merged_df1['NAMA PERUSAHAAN'].apply(lambda x: x if len(x) >= 3 else '')
        # Process company addresses
        merged_df1['ALAMAT PERUSAHAAN'] = merged_df1['ALAMAT PERUSAHAAN'].str.title()
        merged_df1['ALAMAT PERUSAHAAN'] = merged_df1['ALAMAT PERUSAHAAN'].replace('', pd.NA).replace('Nan', pd.NA, regex=True)
        merged_df1['ALAMAT PERUSAHAAN'] = merged_df1['ALAMAT PERUSAHAAN'].str.title()
        merged_df1['DOMISILI'] = merged_df1['KOTA']
        # Process names
        def normalize_and_capitalize(name):
            normalized_name = unicodedata.normalize('NFKD', name)
            ascii_name = ''.join(c for c in normalized_name if not unicodedata.combining(c))
            proper_name = ascii_name.title()
            return proper_name
        merged_df1['NAMA'] = merged_df1['NAMA'].apply(normalize_and_capitalize)
        merged_df1['NAMA'] = merged_df1['NAMA'].apply(
            lambda x: '' if re.search(r'[^\x00-\x7F]', x) or len(x) < 3 or len(x) > 50 else x
        )
        def extract_name_from_email(email):
            username = email.split('@')[0]
            name_parts = re.split(r'[._-]', username)
            proper_name = ' '.join(part.capitalize() for part in name_parts)
            return proper_name
        merged_df1['NAMA'] = merged_df1.apply(
            lambda row: extract_name_from_email(row['EMAIL']) if row['NAMA'] == '' else row['NAMA'], axis=1
        )
        merged_df1['NAMA'] = merged_df1['NAMA'].str.replace(r'[^a-zA-Z\s]', '', regex=True)
        
        merged_df1['EMAIL'] = merged_df1['EMAIL'].apply(lambda x: x if len(x) >= 10 else '')
        merged_df1['EMAIL'] = merged_df1['EMAIL'].apply(lambda x: x if '@' in str(x) else '')

        # Special case
        merged_df1.loc[merged_df1['KAMPUS'].str.contains('STIH', na=False), 'PRODI'] = 'S1 Ilmu Hukum'
        merged_df1.loc[merged_df1['PRODI'].str.contains('(stai)', na=False), 'KAMPUS'] = 'STAI Al-Muhajirin Purwakarta'
        merged_df1.loc[merged_df1['PRODI'].str.contains('(unismu)', na=False), 'KAMPUS'] = 'UNISMU Purwakarta'
        merged_df1.loc[merged_df1['PRODI'].str.contains('(itm)', na=False), 'KAMPUS'] = 'ITM Purwakarta'
        # Load dictionaries from Excel
        excel_file = 'dictionary.xlsx'
        dict_prodi = pd.read_excel(r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\dictionary.xlsx', sheet_name='dict_prodi')
        dict_kampus = pd.read_excel(r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\dictionary.xlsx', sheet_name='dict_kampus')
        dict_kota = pd.read_excel(r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\dictionary.xlsx', sheet_name='dict_kota')
        prodi_dict = dict(zip(dict_prodi['input'], dict_prodi['dictionary']))
        kampus_dict = dict(zip(dict_kampus['input'], dict_kampus['dictionary']))
        kota_dict = dict(zip(dict_kota['input'], dict_kota['dictionary']))
        # Map values based on dictionaries
        merged_df1['PRODI'] = merged_df1['PRODI'].map(prodi_dict)
        merged_df1['KAMPUS'] = merged_df1['KAMPUS'].map(kampus_dict)
        merged_df1['KOTA'] = merged_df1['KAMPUS'].map(kota_dict)
        # Define the original and target time zones
        merged_df1['TANGGAL INPUT'] = merged_df1['TANGGAL INPUT'].apply(format_datetime)
        merged_df1['TANGGAL PROSES'] = pd.Timestamp.today()
        merged_df1['TANGGAL UPLOAD'] = merged_df1['TANGGAL UPLOAD'].astype(str).str.strip().replace('nan', '', regex=True)
        merged_df1 = merged_df1.sort_values(by='TANGGAL INPUT', ascending=True)
        merged_df1['NO'] = range(1, len(merged_df1) + 1)
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('program_kelas_karyawan','P2K')
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('program_kelas_reguler','P2R')
    
    # Assuming 'directory' contains the path you're checking
    if directory == r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2K':
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('nan','P2K')
    elif directory == r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R':
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('nan','P2R')
    elif directory == r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\P2R&P2K':
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('nan','P2R&P2K')
    elif directory == r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files\RPL':
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('nan','RPL')
    else:
        merged_df1['EVENT'] = merged_df1['EVENT'].str.replace('nan','P2R&P2K') # Optional: handles cases where the directory doesn't match any condition
    
    merged_df_data_clean = merged_df1
    # Display the merged DataFrame
    print(merged_df_data_clean)
    folder_name = os.path.basename(os.path.normpath(directory))
    # datenow 
    tgl=pd.Timestamp("today").strftime("%Y-%m-%d_")
    
    # Save the merged DataFrame to an Excel file
    output_file = os.path.join(directory_output, f'{tgl}{folder_name}_data_hasil.xlsx')
    merged_df_data_clean.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Merged DataFrame saved to {output_file}")
    # Create a list to hold the combined data with headers
    combined_list = []
    # Iterate through each DataFrame in the dictionary
    for df1_name, df1 in dataframes.items():
        # Add the header row as a DataFrame
        header_df1 = pd.DataFrame([df1.columns.tolist()], columns=df1.columns)
        combined_list.append(header_df1)
        # Add the DataFrame's data with reset index
        df1 = df1.reset_index(drop=True)
        combined_list.append(df1)
    # Concatenate all DataFrames into one
    combined_df1 = pd.concat(combined_list, ignore_index=True)
    # Save the combined DataFrame to an Excel file
    output_file_data_asli = os.path.join(directory_output, f'{tgl}{folder_name}_data_asli.xlsx')
    with pd.ExcelWriter(output_file_data_asli, engine='openpyxl') as writer:
        combined_df1.to_excel(writer, index=False, header=False)
    print(f"Combined DataFrame saved to {output_file_data_asli}")
    # Rename all XML files back to XLS
    rename_xml_to_xls(directory)    
#====================================================================================================================================#
def process_directory_affiliator(directory):
    print(f"Processing directory: {directory}")
    # Rename all XLS files to XML
    rename_xls_to_xml(directory)
    # Get file paths after renaming
    file_paths = get_files_from_directory(directory)
    # Store DataFrames in a dictionary
    dataframes = {}
    for idx, file_path in enumerate(file_paths, start=1):
        try:
            df2_name = f"df2{idx}"
            if file_path.endswith('.csv'):
                df2 = read_csv_with_detected_encoding(file_path)
                df2 = process_single_column_df(df2)
            elif file_path.endswith('.xml'):
                df2 = parse_xml_to_df(file_path)
            dataframes[df2_name] = df2
            print(f"Data from {file_path} has been loaded into {df2_name}")
        except Exception as e:
            print(f"Error reading {file_path}: {e}")
    # Load the dictionary from Excel
    dict_df2 = pd.read_excel(r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Excel\dictionary.xlsx')
    # Convert the dictionary DataFrame to a dict of dicts
    header_dict = {}
    for column in dict_df2.columns:
        if column.endswith('_dict'):
            key = column.replace('_dict', '')
            header_dict[key] = dict(zip(dict_df2[f'{key}_input'].dropna(), dict_df2[column].dropna()))
    # Standard column names for the merged DataFrame
    standard_columns_2 = ['NO', 'SUMBER DATA', 'NAMA', 'NO TELEPON', 'EMAIL', 'KOTA', 'NAMA PERUSAHAAN', 'METODE PRESENTASI', 'ALAMAT PRESENTASI', 'JAM PRESENTASI YANG DIINGINKAN', 'CALON MAHASISWA', 'TANGGAL INPUT', 'TANGGAL PROSES', 'TANGGAL UPLOAD']
    # Create an empty DataFrame to store the merged data
    merged_df2 = pd.DataFrame(columns=standard_columns_2)
    # Function to rename columns using the dictionary
    def rename_columns(df2, header_dict):
        temp_df2 = pd.DataFrame(columns=standard_columns_2)
        for key, mapping in header_dict.items():
            for input_col, dict_col in mapping.items():
                if input_col in df2.columns:
                    temp_df2[dict_col] = df2[input_col]
                    break
        return temp_df2
    # Loop through each DataFrame in the dictionary
    for key, df2 in dataframes.items():
        temp_df2 = rename_columns(df2, header_dict)
        merged_df2 = pd.concat([merged_df2, temp_df2], ignore_index=True)
    if merged_df2.empty:
        print("data kosong")
    else:
        # Further processing and saving the merged DataFrame
        merged_df2 = merged_df2.applymap(str)
        # Replace source data values
        merged_df2['SUMBER DATA'] = merged_df2['SUMBER DATA'].str.replace('fb', 'Facebook ADS')
        merged_df2['SUMBER DATA'] = merged_df2['SUMBER DATA'].str.replace('ig', 'Instagram ADS')
        # Process phone numbers
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].astype(str)
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].apply(lambda x: re.sub(r'[^\x00-\x7F]+', '', x))
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].str.replace(' ', '')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].str.replace('-', '')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].str.lstrip('+p')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].str.lstrip('62')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].str.lstrip('+62')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].str.lstrip('0')
        merged_df2['NO TELEPON'] = '0' + merged_df2['NO TELEPON']
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].apply(lambda x: x if len(x) >= 9 else '')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].apply(lambda x: x if len(x) <= 13 else '')
        merged_df2['NO TELEPON'] = merged_df2['NO TELEPON'].apply(lambda x: x if str(x).isdigit() else '')
        # Process city names
        merged_df2['KOTA'] = merged_df2['KOTA'].str.title()
        # Process company names
        merged_df2['NAMA PERUSAHAAN'] = merged_df2['NAMA PERUSAHAAN'].astype(str).str.strip().replace('nan', '', regex=True)
        merged_df2['NAMA PERUSAHAAN'] = merged_df2['NAMA PERUSAHAAN'].apply(lambda x: x if len(x) >= 3 else '')
        # Process company addresses
        merged_df2['ALAMAT PRESENTASI'] = merged_df2['ALAMAT PRESENTASI'].str.title()
        merged_df2['ALAMAT PRESENTASI'] = merged_df2['ALAMAT PRESENTASI'].replace('', pd.NA).replace('Nan', pd.NA, regex=True)
        # Process names
        def normalize_and_capitalize(name):
            normalized_name = unicodedata.normalize('NFKD', name)
            ascii_name = ''.join(c for c in normalized_name if not unicodedata.combining(c))
            proper_name = ascii_name.title()
            return proper_name
        merged_df2['NAMA'] = merged_df2['NAMA'].apply(normalize_and_capitalize)
        merged_df2['NAMA'] = merged_df2['NAMA'].apply(
            lambda x: '' if re.search(r'[^\x00-\x7F]', x) or len(x) < 3 or len(x) > 50 else x
        )
        def extract_name_from_email(email):
            username = email.split('@')[0]
            name_parts = re.split(r'[._-]', username)
            proper_name = ' '.join(part.capitalize() for part in name_parts)
            return proper_name
        merged_df2['NAMA'] = merged_df2.apply(
            lambda row: extract_name_from_email(row['EMAIL']) if row['NAMA'] == '' else row['NAMA'], axis=1
        )
        merged_df2['NAMA'] = merged_df2['NAMA'].str.replace(r'[^a-zA-Z\s]', '', regex=True)
        merged_df2['EMAIL'] = merged_df2['EMAIL'].apply(lambda x: x if len(x) >= 10 else '')
        merged_df2['EMAIL'] = merged_df2['EMAIL'].apply(lambda x: x if '@' in str(x) else '')
                # Define the original and target time zones
        merged_df2['TANGGAL INPUT'] = merged_df2['TANGGAL INPUT'].apply(format_datetime)
        merged_df2['TANGGAL PROSES'] = pd.Timestamp.today()
        merged_df2['TANGGAL UPLOAD'] = merged_df2['TANGGAL UPLOAD'].astype(str).str.strip().replace('nan', '', regex=True)
        merged_df2 = merged_df2.sort_values(by='TANGGAL INPUT', ascending=True)
        merged_df2['NO'] = range(1, len(merged_df2) + 1)
    merged_df2 = merged_df2[['NO', 'SUMBER DATA', 'NAMA', 'NO TELEPON', 'EMAIL', 'KOTA', 'NAMA PERUSAHAAN', 'METODE PRESENTASI', 'ALAMAT PRESENTASI', 'JAM PRESENTASI YANG DIINGINKAN', 'CALON MAHASISWA', 'TANGGAL INPUT', 'TANGGAL PROSES', 'TANGGAL UPLOAD']]
    # Display the merged DataFrame
    print(merged_df2)
    folder_name = os.path.basename(os.path.normpath(directory))
    # datenow 
    tgl=pd.Timestamp("today").strftime("%Y-%m-%d_")
    
    # Save the merged DataFrame to an Excel file
    output_file = os.path.join(directory_output, f'{tgl}{folder_name}_data_hasil.xlsx')
    merged_df2.to_excel(output_file, index=False, engine='openpyxl')
    print(f"Merged DataFrame saved to {output_file}")
    # Create a list to hold the combined data with headers
    combined_list = []
    # Iterate through each DataFrame in the dictionary
    for df2_name, df2 in dataframes.items():
        # Add the header row as a DataFrame
        header_df2 = pd.DataFrame([df2.columns.tolist()], columns=df2.columns)
        combined_list.append(header_df2)
        # Add the DataFrame's data with reset index
        df2 = df2.reset_index(drop=True)
        combined_list.append(df2)
    # Concatenate all DataFrames into one
    combined_df2 = pd.concat(combined_list, ignore_index=True)
    # Save the combined DataFrame to an Excel file
    output_file_data_asli = os.path.join(directory_output, f'{tgl}{folder_name}_data_asli.xlsx')
    with pd.ExcelWriter(output_file_data_asli, engine='openpyxl') as writer:
        combined_df2.to_excel(writer, index=False, header=False)
    print(f"Combined DataFrame saved to {output_file_data_asli}")
    # Rename all XML files back to XLS
    rename_xml_to_xls(directory)
#============================================================================================================================================================#
#============================================================================================================================================================#
# Process each directory
for directory in directories_kuliah:
    if len(os.listdir(directory)) == 0:
        continue
    else:    
        process_directory_kuliah(directory)
for directory in directories_affiliator:
    if len(os.listdir(directory)) == 0:
        continue
    else:    
        process_directory_affiliator(directory)
#============================================================================================================================================================#
#============================================================================================================================================================#
# Get current date and month in the desired format
tgl = datetime.now().strftime("%Y-%m-%d")  # Current date as YYYY-MM-DD
month = datetime.now().strftime("%Y-%m")  # Current month as YYYY-MM

# Define source and base directories
file_source = r'C:\Users\lenovo\Desktop\Gilland Ganesha\New Files'
base_dir = r'C:\Users\lenovo\Desktop\Gilland Ganesha\3. Data Hasil Cleaning'

# Define the paths for the month and date directories
month_dir = os.path.join(base_dir, month)
date_dir = os.path.join(month_dir, tgl)

# Create the directories if they don't exist
os.makedirs(date_dir, exist_ok=True)

# List all files in the source directory
all_files = os.listdir(file_source)

# Move each file from the source directory to the destination directory
for file_name in all_files:
    src_path = os.path.join(file_source, file_name)
    dst_path = os.path.join(date_dir, file_name)
    
    # Move the file
    shutil.move(src_path, dst_path)

print(f'All files moved to: {date_dir}')


# In[ ]:




