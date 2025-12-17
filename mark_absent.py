import os
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
from dotenv import load_dotenv
from sheets_handler import SheetsHandler

load_dotenv()

def mark_absent_employees(target_date=None):
    print(f"Checking for absentees (Date: {target_date or 'Today'})...")
    
    # Initialize Handler
    SHEET_ID = os.getenv("SHEET_ID")
    if not SHEET_ID:
        print("❌ Error: SHEET_ID not found in .env")
        return

    handler = SheetsHandler("credentials.json", SHEET_ID)
    if not handler.sheet:
        print("❌ Error: Could not connect to sheet.")
        return

    # 1. Get all employees
    try:
        emp_sheet = handler.client.open_by_key(SHEET_ID).worksheet("Employees")
        employees = emp_sheet.get_all_records()
        all_names = [e['Name'] for e in employees if e['Name']]
    except Exception as e:
        print(f"❌ Error fetching employees: {e}")
        return

    # 2. Get attendance for target date
    import pytz
    tz = pytz.timezone('Asia/Kolkata')
    
    if target_date:
        current_date_str = target_date
    else:
        current_date_str = datetime.now(tz).strftime("%Y-%m-%d")
        
    print(f"DEBUG: Checking attendance for Date: {current_date_str}")
    
    try:
        attendance_records = handler.sheet.get_all_records()
        # Debug: Print first few records
        # print(f"DEBUG: First record date: {attendance_records[0]['Date']}")
        
        present_names = [
            r['Name'] for r in attendance_records 
            if str(r['Date']).strip() == current_date_str
        ]
        print(f"DEBUG: Found {len(present_names)} present: {present_names}")
    except Exception as e:
        print(f"❌ Error fetching attendance: {e}")
        return []

    # 3. Find missing people
    # Improved Fuzzy Matching Logic
    absent_names = []
    
    # Pre-process lists for comparison
    present_map = {name.lower().strip(): name for name in present_names}
    
    for emp_name in all_names:
        emp_clean = emp_name.lower().strip()
        is_present = False
        
        # Check against all present names
        for present_clean in present_map.keys():
            if emp_clean == present_clean:
                is_present = True
                break
            
            if len(emp_clean) > 3 and len(present_clean) > 3: 
                if emp_clean in present_clean or present_clean in emp_clean:
                    is_present = True
                    break
        
        if not is_present:
            absent_names.append(emp_name)
    
    if not absent_names:
        print("✅ Everyone is present today!")
        return []

    print(f"⚠️ Marking {len(absent_names)} people as Absent: {', '.join(absent_names)}")

    # 4. Mark them as Absent (Red)
    for name in absent_names:
        phone = next((e['Phone'] for e in employees if e['Name'] == name), "")
        
        handler.mark_attendance(
            date=current_date_str, 
            name=name, 
            phone=phone, 
            time="--", 
            status="Absent", 
            color="red"
        )
        print(f"   - Marked {name} as Absent")
        
    return absent_names

if __name__ == "__main__":
    mark_absent_employees()
