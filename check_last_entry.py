import os
import gspread
from google.oauth2.service_account import Credentials
from dotenv import load_dotenv

load_dotenv()

def check_last_entry():
    print("Checking sheet...")
    SHEET_ID = os.getenv("SHEET_ID")
    scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
    creds = Credentials.from_service_account_file("credentials.json", scopes=scope)
    client = gspread.authorize(creds)
    
    sheet = client.open_by_key(SHEET_ID).sheet1
    all_values = sheet.get_all_values()
    
    if len(all_values) > 1:
        last_row = all_values[-1]
        print(f"✅ Last Record: {last_row}")
    else:
        print("⚠️ Sheet is empty (except headers).")

if __name__ == "__main__":
    check_last_entry()
