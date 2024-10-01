#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import os

# Define the folder where Python programs are located
program_folder = r'C:\Users\lenovo\Desktop\Gilland Ganesha\Program\Python'

# Function to list programs and menu options
def list_programs():
    programs = {
        1: "1. Program Auto Scroll (6 chrome).py",
        2: "2. Program Auto Click (6 chrome).py",
        3: "3. Program Mengkopi Data Mentah ke Folder 1 dan 2.py",
        4: "4. Program Mendownload File Dictionary dan Template BDC.py",
        5: "5. Program Cleaning Data dan Mengkopi ke Folder 3.py",
        6: "6. Program Membuat File BDC dan Dashboard.py",
        7: "7. Program Menambahkan Dashboard.py",
        8: "8. Program Mengkopi Semua File Baru ke Gdrive.py",
    }
    return programs

# Function to display the menu to the user
def display_menu():
    programs = list_programs()
    print("\nDaftar Program:")
    for index, program in programs.items():
        print(f"{program.replace('.py', '').capitalize()}")
    print("0. Keluar")

# Function to run the selected program
def run_program(selection):
    programs = list_programs()
    program = programs.get(selection)
    if program:
        program_path = os.path.join(program_folder, program)
        print(f"Menjalankan {program} di {program_path}...")
        os.system(f'python "{program_path}"')  # Run the selected program
        input("\nTekan Enter untuk kembali ke menu utama...")  # Wait before returning to the menu

# Main function to run the program with the menu
def main():
    while True:
        display_menu()
        try:
            # Prompt user to select a program
            choice = int(input("\nPilih program (dengan nomor): "))
            if choice == 0:
                print("Keluar...")
                break
            elif choice in list_programs():
                run_program(choice)
            else:
                print("Pilihan tidak valid. Coba lagi.")
        except ValueError:
            print("Masukkan angka yang valid.")

if __name__ == "__main__":
    main()

