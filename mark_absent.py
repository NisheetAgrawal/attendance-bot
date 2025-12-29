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
        return [], {"error": "Could not connect to Google Sheet (Check ID)"}

    # 1. Get all employees
    try:
        emp_sheet = handler.client.open_by_key(SHEET_ID).worksheet("Employees")
        emp_rows = emp_sheet.get_all_values()
        if not emp_rows: return [], {}
        
        # Scan first 20 rows for "Name" column in Employees sheet
        emp_header_idx = -1
        name_idx = -1
        
        for i, row in enumerate(emp_rows[:20]):
            norm_row = [str(cell).strip().lower() for cell in row]
            if any(h in ["name", "employee name", "employee"] for h in norm_row):
                emp_header_idx = i
                # Find exact index
                for col_i, cell_val in enumerate(norm_row):
                    if cell_val in ["name", "employee name", "employee"]:
                        name_idx = col_i
                        break
                break
        
        if emp_header_idx == -1:
             msg = f"'Name' col missing in Employees (Checked top 20 rows). Top row: {emp_rows[0]}"
             print(f"❌ {msg}")
             return [], {"error": msg}
             
        all_names = [row[name_idx] for row in emp_rows[emp_header_idx+1:] if len(row) > name_idx and row[name_idx]]
        print(f"DEBUG: Found {len(all_names)} employees: {all_names}")
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
            return [], {"error": "Main Sheet is empty/blank"}
            
        # Scan first 20 rows for "Date" and "Name" columns
        header_row_idx = -1
        date_idx = -1
        name_idx = -1
        
        # Allow common aliases
        date_col_aliases = ["date", "timestamp", "time in"]
        name_col_aliases = ["name", "employee name", "employee", "slack name"]
        
        for i, row in enumerate(all_rows[:20]): # Check top 20 rows only
            norm_row = [str(cell).strip().lower() for cell in row]
            
            # Check if this row has BOTH target columns
            found_date = any(h in date_col_aliases for h in norm_row)
            found_name = any(h in name_col_aliases for h in norm_row)
            
            if found_date and found_name:
                header_row_idx = i
                # Find exact indices
                # We need to be careful to pick the *first* match in the row
                for col_i, cell_val in enumerate(norm_row):
                    if date_idx == -1 and cell_val in date_col_aliases:
                        date_idx = col_i
                    if name_idx == -1 and cell_val in name_col_aliases:
                        name_idx = col_i
                break
        
        if header_row_idx == -1:
             print("⚠️ Warning: Could not find headers. Defaulting to Col A (Date) and Col B (Name).")
             # Fallback: Assume Data starts at Row 1 (after separator?) or just scan ALL rows safely
             # standard Google Sheet structure is Row 1 = Header, but if shifted, we just search data.
             header_row_idx = 0 
             date_idx = 0
             name_idx = 1
             # We assume data is "Date, Name, ..."

        # Data starts after the header (or what we think is header)
        attendance_records = all_rows[header_row_idx+1:]
        
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
        phone = ""
        
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
