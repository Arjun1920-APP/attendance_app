import gspread
import pandas as pd
from google.oauth2.credentials import Credentials
from datetime import datetime
import random
import string
import os
import json
import time
from tenacity import retry, stop_after_attempt, wait_exponential

# ✅ Google Sheet ID
SPREADSHEET_ID = "16j_H3ND9BrBGucTxv5PIyvI22P5Q7xSCHsAelQbpOyY"

# -------------------- Google Auth --------------------
def get_gsheet_client():
    """
    Create a Google Sheets client using OAuth.
    Works on both Render (via GOOGLE_TOKEN_JSON env var)
    and locally (via token.json file).
    """
    creds = None

    try:
        # ✅ When running on Render — read token from environment
        if os.getenv("GOOGLE_TOKEN_JSON"):
            print("✅ Using OAuth token from Render environment")
            token_info = json.loads(os.getenv("GOOGLE_TOKEN_JSON"))
            creds = Credentials.from_authorized_user_info(token_info)

        # ✅ Local fallback — use token.json file
        else:
            print("⚠️ Using local token.json file")
            creds = Credentials.from_authorized_user_file("token.json")

        client = gspread.authorize(creds)
        return client

    except Exception as e:
        print("❌ Failed to load Google OAuth token:", str(e))
        raise
    # 2. Fallback for local testing (Development)
    file_path = "token.json"
    
    if os.path.exists(file_path):
        print("Authenticating via local file (token.json)...")
        # Original method for reading the local file:
        creds = Credentials.from_authorized_user_file(file_path)
        return gspread.authorize(creds)
    else:
        # If neither credential source exists, fail.
        raise IOError(f"Authentication failed: token.json not found locally, and GOOGLE_TOKEN_JSON env var is missing.")

# -------------------- Upload Session Excel --------------------
@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=2, max=10))
def upload_session_from_excel(file_path, session_name, session_date):
    """Upload session Excel data into Master_Attendance tab"""
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)

    # ✅ Try to open Master_Attendance tab
    try:
        ws = sh.worksheet("Master_Attendance")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title="Master_Attendance", rows="100", cols="20")
        ws.append_row([
            "Session ID", "Session Name", "Session Date",
            "Employee Code", "Employee Name", "Official Email", "Business",
            "Attendance", "Timestamp"
        ])

    # ✅ Read Excel
    df = pd.read_excel(file_path)
    df.columns = [str(c).strip() for c in df.columns]

    # ✅ Generate unique Session ID (avoid conflicts)
    rand_suffix = ''.join(random.choices(string.ascii_uppercase + string.digits, k=4))
    session_id = f"{session_name.strip().replace(' ', '_')}_{session_date}_{rand_suffix}"

    # ✅ Add session details automatically
    df.insert(0, "Session Date", session_date)
    df.insert(0, "Session Name", session_name)
    df.insert(0, "Session ID", session_id)
    df["Attendance"] = ""
    df["Timestamp"] = ""

    # ✅ Append data into Master_Attendance
    ws.append_rows(df.values.tolist())

    print(f"✅ Uploaded {len(df)} employees to Master_Attendance ({session_id})")
    return session_id

# -------------------- Mark Attendance (for QR scan/morning check-in) --------------------
@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=2, max=10))
def mark_present(session_id, email):
    """Mark 'Present' for given email in Master_Attendance if record exists and not marked."""
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    
    try:
        ws = sh.worksheet("Master_Attendance")
    except gspread.exceptions.WorksheetNotFound:
        print("Master_Attendance sheet not found.")
        return False

    # 1. Get header row to find column indices dynamically
    header = ws.row_values(1) # This returns a list of values from row 1 (the headers)
    
    try:
        attendance_col = header.index("Attendance") + 1
        email_col = header.index("Official Email") + 1
        timestamp_col = header.index("Timestamp") + 1
        session_id_col = header.index("Session ID") + 1
    except ValueError as e:
        print(f"Error: Missing column in Master_Attendance sheet: {e}")
        return False

    all_values = ws.get_all_values() 
    
    for i, row in enumerate(all_values[1:], start=2): 
        session_match = row[session_id_col - 1] == session_id 
        email_match = row[email_col - 1].strip().lower() == email.strip().lower()
        
        if session_match and email_match:
            current_status = row[attendance_col - 1].strip().lower()
            
            if current_status == "present":
                print(f"ℹ️ Attendance already marked for: {email}")
                return True

            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # --- FIXED: 1 API Call instead of 2 ---
            cell_range = f"{gspread.utils.rowcol_to_a1(i, attendance_col)}:{gspread.utils.rowcol_to_a1(i, timestamp_col)}"
            ws.update(cell_range, [["Present", timestamp]])
            
            print(f"✅ Attendance marked for: {email} (Row {i})")
            return True
            
    print(f"❌ Email '{email}' not found for Session ID '{session_id}' in attendance list.")
    return False

@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=2, max=10))
def check_email_exists_for_feedback(session_id, email):
    """Checks if the email exists on the Master_Attendance list for the given session."""
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet("Master_Attendance")
    except gspread.exceptions.WorksheetNotFound:
        return False # Treat as not found if sheet is missing

    header = ws.row_values(1)
    
    try:
        email_col = header.index("Official Email") + 1
        session_id_col = header.index("Session ID") + 1
    except ValueError:
        print("Error: Missing required columns for email check.")
        return False
        
    all_values = ws.get_all_values()
    
    for row in all_values[1:]: # Start from row 2
        if (row[session_id_col - 1] == session_id and
            row[email_col - 1].strip().lower() == email.strip().lower()):
            return True # Email found!

    return False # Email not found

# -------------------- Mark Attendance (for Feedback check-in) --------------------
@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=2, max=10))
def check_and_mark_attendance_from_feedback(session_id, email, name, phone, session_name, session_date):
    """
    Checks Master_Attendance:
    1. If Email is found AND Attendance is empty, mark 'Present'.
    2. If Email is found AND Attendance is 'Present', do nothing.
    """
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet("Master_Attendance")
    except gspread.exceptions.WorksheetNotFound:
        return {'marked_now': False, 'status': 'Sheet not found'}

    header = ws.row_values(1)
    
    try:
        attendance_col = header.index("Attendance") + 1
        email_col = header.index("Official Email") + 1
        timestamp_col = header.index("Timestamp") + 1
        session_id_col = header.index("Session ID") + 1
    except ValueError:
        print("Error: Missing required columns in Master_Attendance sheet.")
        return {'marked_now': False, 'status': 'Missing columns'}
        
    all_values = ws.get_all_values()
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    for i, row in enumerate(all_values[1:], start=2): 
        if (row[session_id_col - 1] == session_id and
            row[email_col - 1].strip().lower() == email.strip().lower()):
            
            current_status = row[attendance_col - 1].strip()
            
            if current_status == "":
                # --- FIXED: 1 API Call instead of 2 ---
                cell_range = f"{gspread.utils.rowcol_to_a1(i, attendance_col)}:{gspread.utils.rowcol_to_a1(i, timestamp_col)}"
                ws.update(cell_range, [["Present", timestamp]])
                
                print(f"✅ Attendance marked late via feedback for: {email}")
                return {'marked_now': True, 'status': 'Marked Present'}
            else:
                print(f"ℹ️ Attendance already marked for: {email}")
                return {'marked_now': False, 'status': 'Already Present'}
                
    return {'marked_now': False, 'status': 'Email not on master list'}

# -------------------- Append Feedback --------------------
@retry(stop=stop_after_attempt(5), wait=wait_exponential(multiplier=1, min=2, max=10))
def append_feedback(session_id, session_name, session_date, data):
    """Append feedback row into Master_Feedback"""
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)

    try:
        ws = sh.worksheet("Master_Feedback")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title="Master_Feedback", rows="100", cols="20")
        ws.append_row([
            "Timestamp", "Session ID", "Session Name", "Session Date",
            "Employee Name", "Email", "Phone",
            "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "Q10"
        ])

    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        session_id, session_name, session_date,
        data.get("name", ""), data.get("email", ""), data.get("phone", "")
    ]

    for i in range(1, 11):
        row.append(data.get(f"Q{i}", ""))

    ws.append_row(row)
    print(f"✅ Feedback added for {session_name} ({session_id})")