import os
from datetime import datetime
from slack_bolt import App
from slack_bolt.adapter.flask import SlackRequestHandler
from flask import Flask, request
from sheets_handler import SheetsHandler
from dotenv import load_dotenv

load_dotenv()

# Initialize Slack App
app = App(
    token=os.environ.get("SLACK_BOT_TOKEN"),
    signing_secret=os.environ.get("SLACK_SIGNING_SECRET")
)
handler = SlackRequestHandler(app)

# Initialize Flask (to host the slack bolt handler)
flask_app = Flask(__name__)

# Initialize Sheets
SHEET_ID = os.getenv("SHEET_ID")
sheets_handler = SheetsHandler("credentials.json", SHEET_ID)

# Initialize Scheduler for Absenteeism
from apscheduler.schedulers.background import BackgroundScheduler
from mark_absent import mark_absent_employees
import logging
import pytz

# Enable Scheduler Logging
logging.basicConfig()
logging.getLogger('apscheduler').setLevel(logging.DEBUG)

scheduler = BackgroundScheduler()
timezone = pytz.timezone('Asia/Kolkata')

# Run every day at 4:00 PM (16:00) IST
scheduler.add_job(func=mark_absent_employees, trigger="cron", hour=16, minute=0, timezone=timezone)
scheduler.start()
print("⏰ Scheduler started: Will mark absentees everyday at 16:00 IST")

@flask_app.route("/trigger-absent", methods=["GET"])
def manual_absent_trigger():
    mark_absent_employees()
    return "✅ Absentee check triggered manually!", 200

import re

# Match "in", "In", "IN", "iN" (case insensitive, allowing whitespace around)
# Match "in", "In", "IN", "iN" (case insensitive, allowing whitespace around)
# Strict Match: Start of string ^, "in" (case-insensitive), End of string $
# Allows whitespace logic like " in " or "In" but REJECTS "I am in"
@app.message(re.compile(r"^\s*in\s*$", re.IGNORECASE))
def handle_attendance(message, say, logger):
    user_id = message['user']
    
    # Get user info for real name
    try:
        result = app.client.users_info(user=user_id)
        user_name = result['user']['real_name'] or result['user']['name']
    except Exception as e:
        user_name = f"Unknown ({user_id})"
        logger.error(f"Error fetching user info: {e}")

    now = datetime.now()
    current_time = now.strftime("%H:%M:%S")
    current_date = now.strftime("%Y-%m-%d")
    # Logic:
    # < 11:45 AM: Present (Green)
    # 11:45 AM - 12:59 PM: Late (Yellow)
    # >= 1:00 PM: Half Day (Orange)
    
    cutoff_present = now.replace(hour=11, minute=45, second=0, microsecond=0)
    cutoff_late = now.replace(hour=13, minute=0, second=0, microsecond=0) # 1 PM
    
    if now < cutoff_present:
        status = "Present"
        color = "green"
    elif now < cutoff_late:
        status = "Late"
        color = "yellow"
    else:
        status = "Half Day"
        color = "orange"

    # Log to Sheets
    try:
        # We pass user_name directly, no phone number matching needed
        sheets_handler.mark_attendance(current_date, user_name, "Slack", current_time, status, color)
        
        # Threaded reply
        say(
            text=f"✅ Attendance marked for *{user_name}* at {current_time} ({status})",
            thread_ts=message['ts']
        )
    except Exception as e:
        say(f"❌ Error logging to sheet: {str(e)}", thread_ts=message['ts'])

# Match "mark absent" (case insensitive, strict match but tolerant of internal spaces)
@app.message(re.compile(r"mark\s+absent", re.IGNORECASE))
def handle_absent_trigger(message, say, logger):
    logger.info(f"received 'mark absent' trigger from user {message['user']}")
    try:
        say(f"🕵️‍♂️ Checking for absentees...", thread_ts=message['ts'])
        # Expecting tuple (list, dict) now
        result = mark_absent_employees()
        
        # Handle backward compatibility if it returns just list (safety)
        if isinstance(result, tuple):
            absent_list, stats = result
        else:
            absent_list, stats = result, {}

        debug_msg = f"\n(Debug: Checked {stats.get('total_employees', '?')} fail-safes against {stats.get('present_count', '?')} present on {stats.get('checked_date', '?')})"

        if not absent_list:
            say(f"✅ Everyone is present today! {debug_msg}", thread_ts=message['ts'])
        else:
            names_str = ", ".join(absent_list)
            say(f"🔴 Marked {len(absent_list)} people as Absent:\n{names_str} {debug_msg}", thread_ts=message['ts'])
    except Exception as e:
        logger.error(f"Failed to handle absent trigger: {e}")
        say(f"❌ Error running absentee check: {e}", thread_ts=message['ts'])

# Fallback listener for debugging
@app.message(re.compile(r"absent", re.IGNORECASE))
def debug_absent(message, logger):
    logger.info(f"Ignored broad 'absent' message: {message['text']}")

@flask_app.route("/slack/events", methods=["POST"])
def slack_events():
    return handler.handle(request)

if __name__ == "__main__":
    from datetime import timedelta
    # Startup Catch-up: Check yesterday's attendance just in case
    print("🔄 Checking if we missed yesterday's attendance...")
    yesterday = (datetime.now() - timedelta(days=1)).strftime("%Y-%m-%d")
    try:
        mark_absent_employees(yesterday)
    except Exception as e:
        print(f"⚠️ Catch-up failed: {e}")

    # Run Flask
    port = int(os.environ.get("PORT", 5000))
    flask_app.run(host='0.0.0.0', port=port)
