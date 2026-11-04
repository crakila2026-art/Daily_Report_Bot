import pyautogui
import pyperclip
import pygetwindow as gw
import time

import os
from datetime import datetime

# Project configuration
project_folder = os.path.dirname(os.path.abspath(__file__))

today = datetime.now().strftime("%Y-%m-%d")

current_datetime = datetime.now().strftime(
    "%Y-%m-%d %H:%M:%S"
)

excel_filename = f"daily_report_{today}.xlsx"

excel_path = os.path.join(
    project_folder,
    excel_filename
)

screenshot_path = os.path.join(
    project_folder,
    f"daily_report_{today}.png"
)

comment = "Weather information captured successfully."


pyautogui.FAILSAFE = True
pyautogui.PAUSE = 1

# Open Chrome
print("STEP 1: Opening Chrome")

pyautogui.hotkey("win", "r")
time.sleep(2)

pyautogui.write("chrome", interval=0.1)
pyautogui.press("enter")

# Wait for Chrome window
print("STEP 2: Waiting for Chrome window")

chrome_window = None

for attempt in range(20):
    windows = gw.getWindowsWithTitle("Chrome")

    if windows:
        chrome_window = windows[0]
        break

    time.sleep(1)

if chrome_window is None:
    raise Exception("Chrome window was not found.")

# Activate Chrome explicitly
print("STEP 3: Activating Chrome")

if chrome_window.isMinimized:
    chrome_window.restore()

chrome_window.activate()
time.sleep(3)

active_window = gw.getActiveWindow()

if active_window is None or "Chrome" not in active_window.title:
    raise Exception(
        "Chrome is not active. Stopping before copy operation."
    )

print("Active window:", active_window.title)

# Navigate to weather website
print("STEP 4: Opening weather website")

pyautogui.hotkey("ctrl", "l")
time.sleep(1)

pyautogui.write(
    "https://wttr.in/Chennai?format=3",
    interval=0.05
)

pyautogui.press("enter")

time.sleep(15)

# Copy webpage information
print("STEP 5: Selecting webpage information")

pyautogui.hotkey("ctrl", "a")
time.sleep(2)

print("STEP 6: Copying information")

# Confirm Chrome is still active before Ctrl+C
active_window = gw.getActiveWindow()

if active_window is None or "Chrome" not in active_window.title:
    raise Exception("Chrome lost focus. Copy cancelled.")

pyautogui.hotkey("ctrl", "c")
time.sleep(3)

# Read copied information
weather_data = pyperclip.paste().strip()

print("STEP 7: Clipboard verification")
print("Fetched information:", weather_data)

if not weather_data:
    raise Exception("Clipboard is empty.")

print("SUCCESS: Chrome data copied successfully!")


# --------------------------------
# 8. Completion
# --------------------------------

print("\nAutomation completed successfully!")
print("Excel file will be saved at:")
print(excel_path)
print("Excel file:", excel_path)
print("Screenshot:", screenshot_path)
print("Checking Excel file...")

if os.path.exists(excel_path):
    print("SUCCESS: Excel file created!")
else:
    print("ERROR: Excel file was not created!")

    # --------------------------------

print("Opening Microsoft Excel...")

pyautogui.hotkey("win", "r")
time.sleep(2)

pyautogui.write("excel", interval=0.1)
pyautogui.press("enter")

time.sleep(10)

# Create new workbook
pyautogui.hotkey("ctrl", "n")
time.sleep(4)

# --------------------------------
# STEP 9: Enter report information
# --------------------------------

print("Entering report information...")

excel_data = (
    "Date & Time\tFetched Information\tComment\n"
    f"{current_datetime}\t{weather_data}\t{comment}"
)

pyperclip.copy(excel_data)

pyautogui.hotkey("ctrl", "v")
time.sleep(3)

# --------------------------------
# STEP 10: Save Excel file
# --------------------------------

print("Saving Excel file...")

pyautogui.press("f12")
time.sleep(5)

pyperclip.copy(excel_path)

pyautogui.hotkey("ctrl", "a")
pyautogui.hotkey("ctrl", "v")

time.sleep(2)

pyautogui.press("enter")
time.sleep(7)

# Verify Excel output
if os.path.exists(excel_path):
    print("SUCCESS: Excel file created!")
    print(excel_path)
else:
    raise Exception("Excel file was not saved.")

# --------------------------------
# STEP 11: Screenshot
# --------------------------------

print("Taking screenshot...")

pyautogui.screenshot().save(screenshot_path)

print("Screenshot saved:", screenshot_path)

print("AUTOMATION COMPLETED SUCCESSFULLY!")

# --------------------------------
# CREATE EXCEL WORKBOOK
# --------------------------------

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime
import os

# Get project folder
project_folder = os.path.dirname(os.path.abspath(__file__))

# Generate current date and time
today = datetime.now().strftime("%Y-%m-%d")
current_datetime = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

# Create Excel workbook
workbook = Workbook()

# Select worksheet
sheet = workbook.active
sheet.title = "Daily Status Report"

# Write column headers
sheet["A1"] = "Date & Time"
sheet["B1"] = "Fetched Information"
sheet["C1"] = "Comment"

# Write data into second row
sheet["A2"] = current_datetime
sheet["B2"] = weather_data
sheet["C2"] = "Weather information captured successfully."

# Format column headers
for cell in sheet[1]:
    cell.font = Font(bold=True, color="FFFFFF")
    cell.fill = PatternFill(
        start_color="4472C4",
        end_color="4472C4",
        fill_type="solid"
    )
    cell.alignment = Alignment(horizontal="center")

# Adjust column widths
sheet.column_dimensions["A"].width = 22
sheet.column_dimensions["B"].width = 35
sheet.column_dimensions["C"].width = 45

# Save Excel workbook
excel_filename = f"daily_report_{today}.xlsx"
excel_path = os.path.join(project_folder, excel_filename)

workbook.save(excel_path)

print("Excel workbook created successfully!")
print("Saved location:", excel_path)
chrome
