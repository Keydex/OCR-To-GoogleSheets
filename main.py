import time
from datetime import datetime
import cv2
import numpy as nm
import pyautogui
import pytesseract
import win32gui
from PIL import Image
import gspread

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']

# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1Sx-8Tq5b1F8hvlE43iBVKS4NrybPAmqUcw-vwPzloc4'
SAMPLE_RANGE_NAME = 'Class Data!A2:E'

# Allows us to debug ocr and load an existing image into the code
DEBUG_IMAGE = False

# Name to log in google sheets
USER = 'kaitohyodo'
GOOGLE_SHEET_ID = '1Sx-8Tq5b1F8hvlE43iBVKS4NrybPAmqUcw-vwPzloc4'


def cropImage(original, scale=1.0):
    width, height = original.size  # Get dimensions
    left = width * scale
    top = height * scale
    right = 3 * width * scale
    bottom = 3 * height * scale
    return original.crop((left, top, right, bottom))


def screenshot(window_title=""):
    if window_title:
        hwnd = win32gui.FindWindow(None, window_title)
        while not hwnd:
            hwnd = win32gui.FindWindow(None, window_title)
            print("INFO: FFXIV Window not found!")
        # win32gui.SetForegroundWindow(hwnd)
        x, y, x1, y1 = win32gui.GetClientRect(hwnd)
        x, y = win32gui.ClientToScreen(hwnd, (x, y))
        x1, y1 = win32gui.ClientToScreen(hwnd, (x1 - x, y1 - y))
        im = pyautogui.screenshot(region=(x, y, x1, y1))
        return im
    else:
        im = pyautogui.screenshot()
        return im

def logToGoogleSheets(message):
    print("INFO: Logging to google sheets")
    gc = gspread.service_account()
    sheet = gc.open_by_key(GOOGLE_SHEET_ID).sheet1
    sheet.append_row([datetime.now().isoformat(), message, USER])


def parseFFXIV(message):
    if "The lobby server connection has encountered an error" in message:
        logToGoogleSheets("The lobby server connection has encountered an error")
        raise Exception("The lobby disconnected!")
    elif "Players in queue:" in message:
        chunks = message.split('\n')
        for i, x in enumerate(chunks):
            if "Players in queue:" in x:
                queue_length = x.split(' ')[-1][:-1]
                print("INFO: Queue Length of: ", queue_length)
                logToGoogleSheets(queue_length)
                break
        print("INFO: We detected something")
    else:
        print("Info: Nothing to parse!")


def imToString():
    # Path of tesseract executable
    pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract"

    if DEBUG_IMAGE:
        cropped_image = Image.open("./testqueueimage.png")
        cropped_image.load()
        print("DEBUG: Using debug image")
    else:
        # Capture ffxiv client once
        cap = screenshot("Final Fantasy XIV")

        # Crop image to attempt to just capture queue info
        cropped_image = cropImage(cap, 0.25)
        # cropped_image.show()

    # Convert image to monochrome then parse for text
    parsed_strings = pytesseract.image_to_string(
        cv2.cvtColor(nm.array(cropped_image), cv2.COLOR_BGR2GRAY),
        lang='eng')
    return parsed_strings


# Time between parsing
delay_time = 30
start_time = time.time()

try:
    while True:
        print("\nINFO: Checking queue info", datetime.now().isoformat(' ', 'seconds'))
        ffxivMessage = imToString()
        parseFFXIV(ffxivMessage)
        time.sleep(delay_time - ((time.time() - start_time) % delay_time))
except Exception as e:
    # Log error here
    print(e)
    pass
