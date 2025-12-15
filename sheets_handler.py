import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import os

class SheetsHandler:
    def __init__(self, creds_file, sheet_id):
        self.scope = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive"
        ]
        try:
            # Check for Cloud Env Variable first (Render/Heroku/etc)
            import json
            if os.environ.get("GOOGLE_CREDENTIALS_JSON"):
                creds_dict = json.loads(os.environ.get("GOOGLE_CREDENTIALS_JSON"))
                self.creds = Credentials.from_service_account_info(creds_dict, scopes=self.scope)
            else:
                # Fallback to local file
                self.creds = Credentials.from_service_account_file(creds_file, scopes=self.scope)
                
            self.client = gspread.authorize(self.creds)
            self.sheet = self.client.open_by_key(sheet_id).sheet1
        except Exception as e:
            print(f"Warning: Could not connect to Google Sheets. Error: {e}")
            self.sheet = None

    def mark_attendance(self, date, name, phone, time, status, color):
        """
        Logs attendance to the main sheet.
        Columns: Date, Name, Phone, Time, Status
        """
        if not self.sheet: return

        # Check for date change to add separators
        try:
            all_values = self.sheet.get_all_values()
            if len(all_values) > 1: # Header exists
                last_row = all_values[-1]
                last_date_str = last_row[0]
                
                # Check if it's a valid date format (skip if it's a separator or header)
                if last_date_str and last_date_str != "Date" and "---" not in last_date_str:
                    last_date = datetime.strptime(last_date_str, "%Y-%m-%d")
                    curr_date = datetime.strptime(date, "%Y-%m-%d") # Format must match 'current_date' passed in
                    
                    if curr_date > last_date:
                        # Day changed
                        
                        # Check Month Change
                        if curr_date.month != last_date.month:
                            # Month Separator
                            month_name = last_date.strftime("%B %Y")
                            sep_row = [f"--- END OF {month_name} ---", "", "", "", ""]
                            self.sheet.append_row(sep_row)
                            # Formatting for month separator (Bold, Dark Gray)
                            row_idx = len(all_values) + 1
                            self.sheet.format(f"A{row_idx}:E{row_idx}", {
                                "backgroundColor": {"red": 0.4, "green": 0.4, "blue": 0.4},
                                "textFormat": {"foregroundColor": {"red": 1.0, "green": 1.0, "blue": 1.0}, "bold": True},
                                "horizontalAlignment": "CENTER"
                            })
                            self.sheet.merge_cells(f"A{row_idx}:E{row_idx}") # Optional: merge cells
                        else:
                            # Day Separator
                            sep_row = [f"--- End of {last_date_str} ---", "", "", "", ""]
                            self.sheet.append_row(sep_row)
                            # Formatting for day separator (Light Gray)
                            row_idx = len(all_values) + 1
                            self.sheet.format(f"A{row_idx}:E{row_idx}", {
                                "backgroundColor": {"red": 0.9, "green": 0.9, "blue": 0.9},
                                "textFormat": {"italic": True},
                                "horizontalAlignment": "CENTER"
                            })
                            self.sheet.merge_cells(f"A{row_idx}:E{row_idx}")
                            
        except Exception as e:
            print(f"Error checking separators: {e}")

        row = [date, name, phone, time, status]
        self.sheet.append_row(row)
        
        # Apply formatting
        # Get the last row number
        # Note: This might be slow if sheet is huge, but fine for now
        row_idx = len(self.sheet.get_all_values())
        
        # Define colors (RGB)
        colors = {
            "green": {"red": 0.0, "green": 1.0, "blue": 0.0},
            "yellow": {"red": 1.0, "green": 1.0, "blue": 0.0},
            "orange": {"red": 1.0, "green": 0.6, "blue": 0.0},
            "red": {"red": 1.0, "green": 0.0, "blue": 0.0}
        }
        
        target_color = colors.get(color, {"red": 1.0, "green": 1.0, "blue": 1.0})
        
        fmt = {
            "backgroundColor": target_color,
            "textFormat": {"bold": True}
        }
        
        cell_range = f"E{row_idx}"
        self.sheet.format(cell_range, fmt)

