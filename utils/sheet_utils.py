import gspread
import pandas as pd
from google.oauth2.credentials import Credentials
from datetime import datetime
import random
import string

# ✅ Google Sheet ID
SPREADSHEET_ID = "16j_H3ND9BrBGucTxv5PIyvI22P5Q7xSCHsAelQbpOyY"


# -------------------- Google Auth --------------------
def get_gsheet_client():
    """Authorize using OAuth token.json"""
    creds = Credentials.from_authorized_user_file("GOOGLE_TOKEN")
    return gspread.authorize(creds)


# -------------------- Upload Session Excel --------------------
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


# -------------------- Mark Attendance --------------------
# -------------------- Mark Attendance (for QR scan/morning check-in) --------------------
# -------------------- Mark Attendance (for QR scan/morning check-in) --------------------
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
        # Find 0-based index of columns and convert to 1-based index for gspread update_cell
        # Index() returns the 0-based list index, so we add 1 for gspread's 1-based column index.
        attendance_col = header.index("Attendance") + 1
        email_col = header.index("Official Email") + 1
        timestamp_col = header.index("Timestamp") + 1
        session_id_col = header.index("Session ID") + 1
    except ValueError as e:
        # This catches if a column name is not found (e.g., a typo)
        print(f"Error: Missing column in Master_Attendance sheet: {e}")
        return False

    # 2. Get all row values for comparison
    # We use get_all_values() to get rows as lists, which is better for direct indexing.
    all_values = ws.get_all_values() 
    
    # Iterate over rows starting from the second row (index 1 in all_values list)
    # The actual GSheet row number (i) starts at 2 (since header is row 1)
    for i, row in enumerate(all_values[1:], start=2): 
        
        # Check Session ID and Email
        # We use the 0-based list index here: [col_index - 1]
        session_match = row[session_id_col - 1] == session_id 
        email_match = row[email_col - 1].strip().lower() == email.strip().lower()
        
        if session_match and email_match:
            
            # Check current attendance status using 0-based list index
            current_status = row[attendance_col - 1].strip().lower()
            
            if current_status == "present":
                print(f"ℹ️ Attendance already marked for: {email}")
                return True # Already present, no need to update

            # 3. Mark Present and Timestamp (using 1-based column index)
            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            
            # Update Attendance column
            ws.update_cell(i, attendance_col, "Present") 
            
            # Update Timestamp column
            ws.update_cell(i, timestamp_col, timestamp)
            
            print(f"✅ Attendance marked for: {email} (Row {i})")
            return True
            
    print(f"❌ Email '{email}' not found for Session ID '{session_id}' in attendance list.")
    return False
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
        # Check if the record matches Session ID and Email
        if (row[session_id_col - 1] == session_id and
            row[email_col - 1].strip().lower() == email.strip().lower()):
            
            return True # Email found!

    return False # Email not found

# -------------------- Mark Attendance (for Feedback check-in) --------------------
def check_and_mark_attendance_from_feedback(session_id, email, name, phone, session_name, session_date):
    """
    Checks Master_Attendance:
    1. If Email is found AND Attendance is empty, mark 'Present'.
    2. If Email is found AND Attendance is 'Present', do nothing.
    3. If Email is NOT found, you mentioned: 'new row (rare case)' in app.py logic, but that
       is usually only if the email wasn't in the original uploaded list. For safety, 
       we will stick to updating ONLY employees in the original list.
    """
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)
    try:
        ws = sh.worksheet("Master_Attendance")
    except gspread.exceptions.WorksheetNotFound:
        return {'marked_now': False, 'status': 'Sheet not found'}

    header = ws.row_values(1)
    
    # Safely get column indices
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
    
    for i, row in enumerate(all_values[1:], start=2): # Start from row 2 (index 1 in all_values[1:])
        
        # Check if the record matches Session ID and Email
        if (row[session_id_col - 1] == session_id and
            row[email_col - 1].strip().lower() == email.strip().lower()):
            
            current_status = row[attendance_col - 1].strip()
            
            if current_status == "":
                # Found the row, attendance is EMPTY -> Mark Present!
                ws.update_cell(i, attendance_col, "Present")
                ws.update_cell(i, timestamp_col, timestamp)
                print(f"✅ Attendance marked late via feedback for: {email}")
                return {'marked_now': True, 'status': 'Marked Present'}
            else:
                # Found the row, attendance is ALREADY MARKED -> Do nothing
                print(f"ℹ️ Attendance already marked for: {email}")
                return {'marked_now': False, 'status': 'Already Present'}
            print(f"❌ Email '{email}' not found on the master attendance list for Session ID '{session_id}'.")
    
    # Return a specific error status for app.py to handle (STRICT VALIDATION)
    return {'marked_now': False, 'status': 'Email not on master list'}

    # If loop finishes, email was not on the original list for this session
    # Based on your request, we should not add a new row here as the initial 
    # list is loaded from Excel. The logic in app.py has a 'rare case' to append,
    # but the goal is to only mark attendance for people *expected* to attend.
    # To satisfy the "Not found -> new row" part of your app.py logic, 
    # we'll add the new row here.
    
    # New row for a person not on the original list but giving feedback
    new_row = [
        session_id, session_name, session_date,
        "",  # Employee Code
        name,
        email,
        "",  # Business
        "Present",
        timestamp
    ]
    ws.append_row(new_row, value_input_option="USER_ENTERED")
    print(f"⚠️ New employee added and marked present via feedback: {email}")
    return {'marked_now': True, 'status': 'New Entry Marked Present'}


# -------------------- Append Feedback --------------------
def append_feedback(session_id, session_name, session_date, data):
    """Append feedback row into Master_Feedback"""
    client = get_gsheet_client()
    sh = client.open_by_key(SPREADSHEET_ID)

    # ✅ Try to open Master_Feedback tab
    try:
        ws = sh.worksheet("Master_Feedback")
    except gspread.exceptions.WorksheetNotFound:
        ws = sh.add_worksheet(title="Master_Feedback", rows="100", cols="20")
        ws.append_row([
            "Timestamp", "Session ID", "Session Name", "Session Date",
            "Employee Name", "Email", "Phone",
            "Q1", "Q2", "Q3", "Q4", "Q5", "Q6", "Q7", "Q8", "Q9", "Q10"
        ])

    # ✅ Build feedback row
    row = [
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        session_id, session_name, session_date,
        data.get("name", ""), data.get("email", ""), data.get("phone", "")
    ]

    # ✅ Append Q1–Q10
    for i in range(1, 11):
        row.append(data.get(f"Q{i}", ""))

    ws.append_row(row)
    print(f"✅ Feedback added for {session_name} ({session_id})")
