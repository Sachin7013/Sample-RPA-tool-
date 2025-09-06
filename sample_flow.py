import pyautogui
import time
import subprocess

# 1. Open Notepad
subprocess.Popen(["notepad.exe"])
time.sleep(2)  # wait for Notepad to open

# 2. Type text
pyautogui.typewrite("Hello! This is my first RPA bot 🎉", interval=0.2)

# 3. Press Ctrl+S (Save dialog)
pyautogui.hotkey("ctrl", "s")
time.sleep(2)

# 4. Type the file path
pyautogui.typewrite(r"C:\RPA\MyFirstBot.txt", interval=0.2)

# 5. Press Enter (save file)
pyautogui.press("enter")
time.sleep(2)

# 6. Press Alt+F4 (close Notepad)
pyautogui.hotkey("alt", "f4")
