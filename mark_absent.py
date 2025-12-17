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
    if target_date:
        current_date_str = target_date
    else:
        current_date_str = datetime.now().strftime("%Y-%m-%d")
    try:
        attendance_records = handler.sheet.get_all_records()
        present_names = [
            r['Name'] for r in attendance_records 
        if str(r['Date']) == current_date_str
        ]
    except Exception as e:
        print(f"❌ Error fetching attendance: {e}")
        return

    # 3. Find missing people
    # Improved Fuzzy Matching Logic
    absent_names = []
    
    # Pre-process lists for comparison
    # Structure: {'clean_name': 'Original Name'}
    present_map = {name.lower().strip(): name for name in present_names}
    
    for emp_name in all_names:
        emp_clean = emp_name.lower().strip()
        is_present = False
        
        # Check against all present names
        for present_clean in present_map.keys():
            # Match 1: Exact Match
            if emp_clean == present_clean:
                is_present = True
                break
            
            # Match 2: Substring Match (e.g., "Akash" in "Akash Das")
            # We check both ways: if Slack name is part of Sheet name OR Sheet name is part of Slack name
            if len(emp_clean) > 3 and len(present_clean) > 3: # Safety to prevent "Al" matching "Alan" falsely too easily
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
        # Find phone number if available (optional)
        phone = next((e['Phone'] for e in employees if e['Name'] == name), "")
        
        # Log as Absent
        # We use a special time or leave it empty
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
