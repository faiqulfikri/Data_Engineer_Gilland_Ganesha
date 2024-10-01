#!/usr/bin/env python
# coding: utf-8

# In[ ]:


import pyautogui
import time
import keyboard

# Informasi program
print("=================================================")
print("Program auto scroll individual")
print("Pastikan ukuran jendela 640x320 jika resolusi layar 1920x1080")
print("dan posisi tampilan ada di pojok kanan bawah.")
print("=================================================")

# Fungsi untuk menghentikan otomatisasi
def stop_actions():
    global interrupt
    interrupt = True
    print("Perintah berhenti diterima!")

# Fungsi untuk melakukan aksi
def perform_actions():
    global interrupt, repetition_times
    interrupt = False
    
    print("Memulai aksi dalam 5 detik...")
    time.sleep(5)  # Jeda 5 detik sebelum memulai

    for cycle in range(repetition_times):
        for window in action_windows:
            if interrupt:
                print("Menghentikan...")
                return
            for action in window:
                if interrupt:
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

# Fungsi untuk meminta input jumlah klik tambahan per jendela
def get_click_input():
    window_clicks = {}
    print("Masukkan 0 jika tidak ingin ada gerakan pada jendela tersebut.")
    window_names = [
        "atas kiri", 
        "atas tengah", 
        "atas kanan", 
        "bawah kiri", 
        "bawah tengah", 
        "bawah kanan"
    ]
    for i, name in enumerate(window_names, 1):
        while True:
            try:
                clicks = int(input(f"Masukkan jumlah klik tambahan untuk jendela {name}: "))
                if clicks >= 0:
                    break
                else:
                    print("Masukkan bilangan bulat positif.")
            except ValueError:
                print("Input tidak valid. Masukkan angka.")
        window_clicks[i] = clicks  # Simpan jumlah klik tambahan untuk tiap jendela
    return window_clicks

# Meminta input jumlah klik tambahan untuk tiap jendela
window_clicks = get_click_input()

# Posisi klik yang sudah ditentukan untuk tiap jendela
action_windows = []

# Aksi untuk jendela atas kiri
actions_window_atas_kiri = [{"type": "click", "pos": (320, 120), "interval": 0.5}]
actions_window_atas_kiri += [{"type": "click", "pos": (595, 205), "interval": 0.2} for _ in range(window_clicks[1])]
action_windows.append(actions_window_atas_kiri)

# Aksi untuk jendela atas tengah
actions_window_atas_tengah = [{"type": "click", "pos": (960, 120), "interval": 0.5}]
actions_window_atas_tengah += [{"type": "click", "pos": (1235, 205), "interval": 0.2} for _ in range(window_clicks[2])]
action_windows.append(actions_window_atas_tengah)

# Aksi untuk jendela atas kanan
actions_window_atas_kanan = [{"type": "click", "pos": (1600, 120), "interval": 0.5}]
actions_window_atas_kanan += [{"type": "click", "pos": (1875, 205), "interval": 0.2} for _ in range(window_clicks[3])]
action_windows.append(actions_window_atas_kanan)

# Aksi untuk jendela bawah kiri
actions_window_bawah_kiri = [{"type": "click", "pos": (320, 450), "interval": 0.5}]
actions_window_bawah_kiri += [{"type": "click", "pos": (595, 525), "interval": 0.2} for _ in range(window_clicks[4])]
action_windows.append(actions_window_bawah_kiri)

# Aksi untuk jendela bawah tengah
actions_window_bawah_tengah = [{"type": "click", "pos": (960, 450), "interval": 0.5}]
actions_window_bawah_tengah += [{"type": "click", "pos": (1235, 525), "interval": 0.2} for _ in range(window_clicks[5])]
action_windows.append(actions_window_bawah_tengah)

# Aksi untuk jendela bawah kanan
actions_window_bawah_kanan = [{"type": "click", "pos": (1600, 450), "interval": 0.5}]
actions_window_bawah_kanan += [{"type": "click", "pos": (1875, 525), "interval": 0.2} for _ in range(window_clicks[6])]
action_windows.append(actions_window_bawah_kanan)

# Input jumlah pengulangan dari user
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
start_key = 'f1'
stop_key = 'f2'
interrupt = False

# Monitoring tombol start
keyboard.add_hotkey(start_key, perform_actions)

# Monitoring tombol stop
keyboard.add_hotkey(stop_key, stop_actions)

# Menjaga program tetap berjalan
print(f"Tekan {start_key} untuk mulai dan {stop_key} untuk berhenti.")
keyboard.wait(stop_key)


# In[ ]:




