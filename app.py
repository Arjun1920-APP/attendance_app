from flask import Flask, render_template, request, jsonify, session, url_for
import os
from werkzeug.utils import secure_filename
import re
# Import all necessary GSheet functions from sheet_utils
from utils.sheet_utils import (
    get_gsheet_client, 
    upload_session_from_excel, 
    mark_present, # This is the GSheet version!
    check_and_mark_attendance_from_feedback,
    check_email_exists_for_feedback,
    append_feedback as gsheet_append_feedback # Use this alias to avoid conflict if you ever define a local one
)
import pandas as pd
from datetime import datetime

app = Flask(__name__)
app.secret_key = os.getenv("SECRET_KEY", "supersecretkey123")

UPLOAD_FOLDER = "uploads"
QR_FOLDER = "static/qr"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(QR_FOLDER, exist_ok=True)

# NOTE: The Excel-based 'append_feedback' and 'mark_present' functions 
# have been REMOVED from this file to eliminate the PermissionError and conflict.
# All data operations now use the GSheet functions imported above.

# ======================================================
# 🔹 Admin Upload Route
# ======================================================
# app.py

# ======================================================
# 🔹 Admin Upload Route
# ======================================================
@app.route("/admin/upload_session", methods=["GET", "POST"])
def upload_session():
    if request.method == "GET":
        # Renders the styled uploads.html page
        return render_template("uploads.html") 

    # --- POST Handling (Submission Logic) ---
    session_name = request.form.get("session_name")
    session_date = request.form.get("session_date")
    f = request.files.get("file")

    if not (session_name and session_date and f):
        return render_template("uploads.html", message="⚠️ Missing session name, date, or file."), 400

    fname = secure_filename(f.filename)
    path = os.path.join(UPLOAD_FOLDER, fname)
    
    try:
        f.save(path)
        # Calls the GSheet utility function
        session_id = upload_session_from_excel(path, session_name, session_date) 
    except Exception as e:
        # Handle file saving or GSheet upload errors gracefully
        return render_template("uploads.html", message=f"❌ Upload failed: {str(e)}"), 500

    # Build the URLs
    attendance_url = url_for("attendance_form", session_id=session_id, _external=True)
    feedback_url = url_for("index", session_id=session_id, _external=True)

    # --- SUCCESS MESSAGE: Styled with Buttons ---
    success_html = f"""
    <div style="text-align: center; margin-top: 20px;">
        <h3>✅ Session uploaded successfully!</h3>
        <p><b>Session Name:</b> {session_name}</p>
        <p><b>Session Date:</b> {session_date}</p>
        <p><b>Session ID:</b> {session_id}</p>
        <hr style="border-top: 1px solid #ccc; margin: 20px 0;">
        
        <p>
            <a href="{attendance_url}" target="_blank" class="submit-btn" 
               style="display: inline-block; margin: 5px; background: #2980B9; width: 45%;">
                Attendance QR/Link
            </a>
            
            <a href="{feedback_url}" target="_blank" class="submit-btn" 
               style="display: inline-block; margin: 5px; background: #C0392B; width: 45%;">
                Feedback QR/Link
            </a>
        </p>
        
        <p style="margin-top: 20px; font-size: 0.85rem;">
            (Links opened in a new tab)
        </p>
    </div>
    """
    
    # Render the uploads.html template, passing the styled message
    return render_template("uploads.html", message=success_html)

# ======================================================
# 🔹 Attendance Form & Submit
# ======================================================
@app.route("/attendance")
def attendance_form():
    session_id = request.args.get("session_id")
    session_name = ""
    session_date = ""
    if session_id:
        # Find the index of the first four-digit number (the Year)
        import re
        year_match = re.search(r"(\d{4})", session_id)
        
        if year_match:
            # The name part is everything up to the underscore *before* the year
            # Subtract 1 to include the underscore that separates the name and date
            year_start_index = year_match.start() - 1 
            
            # Extract the raw name part 
            raw_session_name = session_id[:year_start_index]
            
            # Replace underscores with spaces for clean display
            session_name = raw_session_name.replace("_", " ").strip()
        
        else:
            # Fallback if no year/date structure is found
            session_name = session_id.replace("_", " ")

    return render_template("attendance.html", session=session_name, session_id=session_id)
# NOTE: Removed 'session_date' from being passed to the template.

@app.route("/submit_attendance", methods=["POST"])
def submit_attendance():
    data = request.get_json()
    email = data.get("email")
    name = data.get("name", "")
    session_id = data.get("session_id")
    if not (email and session_id):
        return jsonify({"status":"error","message":"Missing email or session_id"}), 400

    # Calls the GSheet utility function from sheet_utils.py
    # Removed name=name, as the GSheet utility function only needs session_id and email
    ok = mark_present(session_id, email) 
    
    if ok:
        session['user_name'] = name or email.split('@')[0]
        return jsonify({"status":"success", "redirect": url_for("thankyou_attendance")})
    else:
        # Improved error message clarity
        return jsonify({"status":"error","message":"Email not found for this session. Check for typos."}), 404
# app.py

# ... (after submit_attendance route)

@app.route("/validate_email", methods=["POST"])
def validate_email():
    """Checks if the email exists on the Master_Attendance list for the session."""
    data = request.get_json()
    email = data.get("email", "")
    session_id = data.get("session_id")
    
    if not (email and session_id):
        return jsonify({"status": "error", "message": "Missing email or session ID."}), 400

    # Import the necessary utility function
    from utils.sheet_utils import check_email_exists_for_feedback
    
    # Check email status using a new utility function (defined in step B)
    exists = check_email_exists_for_feedback(session_id, email)
    
    if exists:
        return jsonify({"status": "success", "message": "Email found."})
    else:
        return jsonify({"status": "error", "message": "Email not found on the session's registration list. Please check for typos."}), 404

# ... (before index route)

@app.route("/thankyou_attendance")
def thankyou_attendance():
    user_name = session.get('user_name', 'Participant')
    session_name = session.get('session_name', 'L&D Session')
    return render_template("thankyou_attendance.html", name=user_name, session_name=session_name)

# ======================================================
# 🔹 Feedback Form & Submit
# ======================================================
@app.route("/")
def index():
    session_id = request.args.get("session_id")
    session_name = ""
    session_date = ""
    if session_id:
        # Find the index of the first four-digit number (the Year)
        import re
        year_match = re.search(r"(\d{4})", session_id)
        
        if year_match:
            # The name part is everything up to the underscore *before* the year
            year_start_index = year_match.start() - 1 
            
            # Extract the raw name part
            raw_session_name = session_id[:year_start_index]
            
            # Replace underscores with spaces for clean display
            session_name = raw_session_name.replace("_", " ").strip()
            
            # Optional: Extract date for the hidden field
            # The date part starts at the year and is 10 characters long (YYYY_MM_DD)
            date_part_raw = session_id[year_start_index + 1 : year_start_index + 11]
            session_date = date_part_raw.replace("_", "-")
        
        else:
            # Fallback
            session_name = session_id.replace("_", " ")

    return render_template("index.html", session=session_name, session_id=session_id, session_date=session_date)

@app.route("/submit_feedback", methods=["POST"])
def submit_feedback_route():
    # Remove local imports as they are now at the top
    # from utils.sheet_utils import get_gsheet_service 
    # import gspread 
    # from datetime import datetime 
    
    data = request.get_json()
    session_id = data.get("session_id") or data.get("session")
    session_name = data.get("session_name") or data.get("session")
    session_date = data.get("session_date", "")
    name = data.get("name", "")
    email = data.get("email", "")
    phone = data.get("phone", "")

    if not session_id or not email:
        return jsonify({"status": "error", "message": "Missing session_id or email"}), 400

    # --- STEP 1: Check/Mark attendance via utility function (Corrected call) ---
    attendance_info = check_and_mark_attendance_from_feedback(
        session_id, email, name, phone, session_name, session_date
    )
    attendance_marked = attendance_info['marked_now']

    # --- STEP 2: Append Feedback (Using the imported GSheet version) ---
    
    # Collect Q1-Q10 into the data dictionary for the utility function
    # NOTE: The client-side (index.html) sends q1, q2, etc., but the GSheet utility 
    # expects Q1, Q2, etc. (which matches the GSheet headers).
    feedback_data = data.copy() # Use a copy to avoid modifying request data unnecessarily
    for i in range(1, 11):
        feedback_data[f"Q{i}"] = data.get(f"q{i}", "") # Map 'q1' to 'Q1' in the data structure
        
    gsheet_append_feedback(session_id, session_name, session_date, feedback_data)

    # --- STEP 3: Thank-you redirect ---
    session['user_name'] = name or email.split('@')[0]
    session['attendance_marked'] = attendance_marked

    return jsonify({"status": "success", "redirect": url_for("thankyou")})
@app.route("/thankyou")
def thankyou():
    user_name = session.get('user_name', 'Participant')
    # session_name or other context can be retrieved here if needed
    return render_template("thankyou.html", name=user_name)

# ======================================================
# 🔹 Run App
# ======================================================
if __name__ == "__main__":
    from waitress import serve
    print("🚀 Server started. Ready for high traffic (60+ users).")
    # threads=10 allows 10 users to be processed at the exact same time
    serve(app, host="0.0.0.0", port=5000, threads=10, channel_timeout=60)