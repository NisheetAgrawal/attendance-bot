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
        return [], {}

    handler = SheetsHandler("credentials.json", SHEET_ID)
    if not handler.sheet:
        print("❌ Error: Could not connect to sheet.")
        return [], {}

    # 1. Get all employees
    try:
        emp_sheet = handler.client.open_by_key(SHEET_ID).worksheet("Employees")
        emp_rows = emp_sheet.get_all_values()
        if not emp_rows: return [], {}
        
        emp_headers = emp_rows[0]
        norm_emp_headers = [str(h).strip().lower() for h in emp_headers]
        
        try:
            name_idx = next(i for i, h in enumerate(norm_emp_headers) if h in ["name", "employee name", "employee"])
        except StopIteration:
             print(f"❌ Error: 'Name' column missing in Employees tab. Found: {emp_headers}")
             return [], {}
             
        all_names = [row[name_idx] for row in emp_rows[1:] if len(row) > name_idx and row[name_idx]]
    except Exception as e:
        print(f"❌ Error fetching employees: {e}")
        return [], {}

    # 2. Get attendance for target date
    import pytz
    tz = pytz.timezone('Asia/Kolkata')
    
    if target_date:
        current_date_str = target_date
    else:
        current_date_str = datetime.now(tz).strftime("%Y-%m-%d")
        
    print(f"DEBUG: Checking attendance for Date: {current_date_str}")
    
    try:
        # Use get_all_values() to avoid "Duplicate Header" errors from gspread
        all_rows = handler.sheet.get_all_values()
        if not all_rows:
            print("❌ Error: Sheet is empty")
            return [], {}
            
        headers = all_rows[0]
        # Normalize headers for finding (strip whitespace, lowercase)
        norm_headers = [str(h).strip().lower() for h in headers]
        
        try:
            # Allow common aliases
            date_col_aliases = ["date", "timestamp", "time in"]
            name_col_aliases = ["name", "employee name", "employee", "slack name"]
            
            date_idx = next(i for i, h in enumerate(norm_headers) if h in date_col_aliases)
            name_idx = next(i for i, h in enumerate(norm_headers) if h in name_col_aliases)
        except StopIteration:
            print(f"❌ Error: Could not find Date/Name columns. Found headers: {headers}")
            return [], {}

        attendance_records = all_rows[1:]
        
        present_names = []
        for row in attendance_records:
            # Safety check for row length
            if len(row) <= max(date_idx, name_idx): continue
            
            row_date = row[date_idx]
            row_name = row[name_idx]
            
            if str(row_date).strip() == current_date_str:
                present_names.append(row_name)

        print(f"DEBUG: Found {len(present_names)} present: {present_names}")
    except Exception as e:
        print(f"❌ Error fetching attendance: {e}")
        return [], {}

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
    
    stats = {
        "total_employees": len(all_names),
        "present_count": len(present_names),
        "absent_count": len(absent_names),
        "checked_date": current_date_str,
        "sample_employee": all_names[0] if all_names else "None",
        "sample_present": present_names[0] if present_names else "None"
    }

    if not absent_names:
        print("✅ Everyone is present today!")
        return [], stats

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
        
    return absent_names, stats

if __name__ == "__main__":
    mark_absent_employees()

if __name__ == "__main__":
    mark_absent_employees()
