from datetime import datetime
import pyperclip
import pyautogui

current_date = datetime.now() # Get the current date

def getLastMonthYear():
    if current_date.month == 1:
        return current_date.year - 1
    return current_date.year

def getLastMonth():
    if current_date.month == 1:
        return 12
    return current_date.month - 1

def workaroundWrite(text):
    pyperclip.copy(text)
    pyautogui.hotkey('ctrl', 'v')
    pyperclip.copy('')