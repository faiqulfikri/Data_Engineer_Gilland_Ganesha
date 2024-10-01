#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pyautogui
import time
import keyboard

# Informasi program
print("=================================================")
print("Program auto click")
print("Pastikan ukuran jendela 640x320 jika resolusi layar 1920x1080")
print("dan posisi tampilan ada di pojok kanan bawah.")
print("=================================================")

# Definisi urutan tindakan
actions = [
    {"type": "click", "pos": (320, 256), "interval": 1},   {"type": "click", "pos": (480, 256), "interval": 4},
    {"type": "click", "pos": (320, 200), "interval": 0.5}, {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},        {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},        {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},        {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "click", "pos": (595, 200), "interval": 0.5}, {"type": "click", "pos": (595, 200), "interval": 0.5},
    
    {"type": "click", "pos": (960, 256), "interval": 1},   {"type": "click", "pos": (1120, 256), "interval": 4},
    {"type": "click", "pos": (960, 200), "interval": 0.5}, {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},        {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},        {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},        {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "click", "pos": (1235, 200), "interval": 0.5},{"type": "click", "pos": (1235, 200), "interval": 0.5},
    
    {"type": "click", "pos": (1600, 256), "interval": 1},   {"type": "click", "pos": (1760, 256), "interval": 4},
    {"type": "click", "pos": (1600, 200), "interval": 0.5}, {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "click", "pos": (1875, 200), "interval": 0.5}, {"type": "click", "pos": (1875, 200), "interval": 0.5},
    
    {"type": "click", "pos": (320, 578), "interval": 1},    {"type": "click", "pos": (480, 578), "interval": 4},
    {"type": "click", "pos": (320, 520), "interval": 0.5},  {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "click", "pos": (595, 520), "interval": 0.5},  {"type": "click", "pos": (595, 520), "interval": 0.5},
    
    {"type": "click", "pos": (960, 578), "interval": 1},    {"type": "click", "pos": (1120, 578), "interval": 4},
    {"type": "click", "pos": (960, 520), "interval": 0.5},  {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "click", "pos": (1235, 520), "interval": 0.5}, {"type": "click", "pos": (1235, 520), "interval": 0.5},
    
    {"type": "click", "pos": (1600, 578), "interval": 1},   {"type": "click", "pos": (1760, 578), "interval": 4},
    {"type": "click", "pos": (1600, 520), "interval": 0.5}, {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "tab", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "key", "key": "tab", "interval": 0.5},         {"type": "key", "key": "enter", "interval": 0.5},
    {"type": "click", "pos": (1875, 520), "interval": 0.5}, {"type": "click", "pos": (1875, 520), "interval": 0.5},
]

# Fungsi untuk menghentikan otomatisasi
def stop_actions():
    global interrupt
    interrupt = True
    print("Perintah berhenti diterima!")

# Fungsi untuk melakukan tindakan
def perform_actions():
    global interrupt
    interrupt = False
    
    print("Memulai aksi dalam 5 detik...")
    time.sleep(5)  # Tambahkan jeda 5 detik di sini
    
    for cycle in range(repetition_times):
        for action in actions:
            if interrupt:
                print("Menghentikan...")
                return
            if action["type"] == "click":
                x, y = action["pos"]
                if action.get("double_click"):
                    pyautogui.doubleClick(x, y)
                else:
                    pyautogui.click(x, y)
            elif action["type"] == "key":
                pyautogui.press(action["key"])
            time.sleep(action["interval"])
        print(f"Siklus ke-{cycle + 1} selesai.")

# Meminta jumlah pengulangan dari pengguna
while True:
    try:
        repetition_times = int(input("Masukkan jumlah pengulangan: "))
        if repetition_times > 0:
            break
        else:
            print("Masukkan bilangan bulat positif.")
    except ValueError:
        print("Input tidak valid. Masukkan angka.")

# Tombol untuk memulai/menghentikan aksi
start_key = 'f3'
stop_key = 'f4'
interrupt = False

# Monitor tombol mulai
keyboard.add_hotkey(start_key, perform_actions)

# Monitor tombol berhenti
keyboard.add_hotkey(stop_key, stop_actions)

# Menjaga program tetap berjalan
print(f"Tekan {start_key} untuk mulai dan {stop_key} untuk berhenti.")
keyboard.wait(stop_key)

