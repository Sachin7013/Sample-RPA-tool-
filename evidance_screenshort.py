import pyautogui
import time
import subprocess
import os

# 1. Open File Explorer
subprocess.Popen(["explorer"])
time.sleep(2)

# 2. Type the search folder path (e.g. C:\Users\Sachin\Documents)
pyautogui.typewrite(r"C:\Users\Manoj\Documents")
pyautogui.press("enter")
time.sleep(2)

# 3. Press Ctrl+F to focus on search box
pyautogui.hotkey("ctrl", "f")
time.sleep(1)

# 4. Type the file/folder name you want to search
file_to_search = "response.txt"
pyautogui.typewrite(file_to_search)
pyautogui.press("enter")

# Wait for search results
time.sleep(5)

# 5. Take screenshot as evidence
screenshot_path = r"C:\RPA\Evidence\SearchResult.png"

# Create the Evidence directory if it doesn't exist
evidence_dir = os.path.dirname(screenshot_path)
os.makedirs(evidence_dir, exist_ok=True)

pyautogui.screenshot(screenshot_path)

print(f"Screenshot saved to {screenshot_path}")
