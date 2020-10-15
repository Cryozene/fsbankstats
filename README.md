# fsbankstats
Small script for automated reading of bank stats from Firestone Idle RPG. Feel free to do whatever you want with the code

# Setup
1. Python 3.8 Interpreter
2. Script uses following libraries outside of OCR and Google Sheets: `numpy`, `pytz`, `tkinter`, `pyautogui`
3. OCR:  
   - `cv2` library for searching & locating member symbols  
   - `pytesseract`  library for OCR  
   - Note that you need to install [Google’s Tesseract-OCR Engine](https://github.com/tesseract-ocr/tesseract) *and specify the path to the executable* at the start of `main2.py` as variable where it says `path to pytesseract executable`

4. For Google Sheet Access:
   - `gspread` and `oauth2client` libraries
   - needs a `credentials.json` next to the script containing your private key -> [Google OAuth2](https://cloud.google.com/docs/authentication/getting-started)
   - example sheet similar to the one used for uploading data in the script: [Link](https://docs.google.com/spreadsheets/d/1S492o3-cTvhJo6w4keBw-_byEWjEFRa7GWkRYyMgc_E/edit?usp=sharing)
