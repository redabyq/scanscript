import time
import os
from PIL import Image
import win32com.client
import pyautogui
import threading
from convert import convertJPG
def move():
    while True:
        pyautogui.moveTo(1106, 749, duration=0.5)
        pyautogui.click()
        time.sleep(20)

scanner = win32com.client.Dispatch("WIA.CommonDialog")
scan_folder = "scans"
if not os.path.exists(scan_folder):
    os.makedirs(scan_folder)

scan_count = 7
scan_button_x, scan_button_y = 1106, 749
move_thread = threading.Thread(target=move, daemon=True)
move_thread.start()

while True:
    try:
        image = scanner.ShowAcquireImage()
        pyautogui.moveTo(scan_button_x, scan_button_y, duration=0.5)
        pyautogui.click()
        if image:
            filename = os.path.join(scan_folder, f"scan_{scan_count}.jpg")
            image.SaveFile(filename)
            print(f"Sacn saved: {filename}\nWaiting...")
            scan_count += 1
            print(f"20 s")
            time.sleep(5)
            print(f"15 s")
            time.sleep(5)
            print(f"10 s")
            time.sleep(5)
            print(f"5 s")
            time.sleep(5)

    except Exception as e:
        print(f"Error: {e}")
        time.sleep(5)
