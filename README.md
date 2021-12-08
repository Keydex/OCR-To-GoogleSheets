# Screenshot to OCR to Google Sheets (Python)

This project was built to screenshot a Windows application (Final Fantasy FXIV) and parse its content in order to log queue data into google sheets. This can be modified to work for your use.

## Requirements -
- Windows Computer (You can use others, but win32 needs to be replaced)
- Python 3.6+
- Install [Tesseract OCR](https://github.com/tesseract-ocr/tesseract)
- [Create google oauth credentials](https://docs.gspread.org/en/latest/oauth2.html) - Note on step 7, you can paste your credentials in the code instead

## FFXIV Specifics Notes
This has only really been tested in windowed mode for FFXIV, and this is only supported if the game is on your primary monitor

## Supported Constants
There are three constants,
- `DEBUG_IMAGE` (BOOL) - When enabled, this uses the test image in the repository to parse and test insteading of screenshotting
- `USER` (STRING) - This is the user that is logged in the google sheets for pooled information
- `GOOGLE_SHEET_ID` (STRING) - This is the google sheets identifier to write to. You must add your google secret client email (which is setup when you create a service account) as an editor to your sheet. 
- `DELAY_TIME` (INT) - This is the time in seconds between Screenshots/OCR attempts

## Running Instructions
- Run `pip install -r requirements.txt`
- Run the application `python main.py`

### Troubleshooting
#### `pytesseract.pytesseract.TesseractNotFoundError: tesseract is not installed or it's not in your path`
Your tesseract installation location is not default to `"C:\Program Files\Tesseract-OCR\tesseract"`, so change the link to the path of where you installed it in the code