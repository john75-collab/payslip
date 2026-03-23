import os
os.environ["OAUTHLIB_INSECURE_TRANSPORT"] = "1"
import time, calendar, smtplib, pdfkit
from datetime import date, datetime
from flask import Flask, render_template, request, redirect, session, send_file, url_for
import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as UserCredentials
from google_auth_oauthlib.flow import Flow
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
#from email.mime.application import MIMEApplication
from googleapiclient.errors import HttpError
import calendar
import imgkit
from email.mime.image import MIMEImage
import base64
from num2words import num2words

# ================= APP =================
app = Flask(__name__)
app.secret_key = "change_this_secret"

# ================= CONFIG =================
SENDER_EMAIL = "dioriderjohn75@gmail.com"
SENDER_PASSWORD = "zgdp uixj rjes evlc"

DRIVE_FOLDER_ID = "1K6wistYGpv2YgSM4aBvxpMQNqAUmfaaq"
PDF_FOLDER = "payslips"
os.makedirs(PDF_FOLDER, exist_ok=True)

WKHTML_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe"
config = pdfkit.configuration(wkhtmltopdf=WKHTML_PATH)

WKHTML_IMAGE_PATH = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltoimage.exe"
img_config = imgkit.config(wkhtmltoimage=WKHTML_IMAGE_PATH)

options = {
    "enable-local-file-access": "",
}

IMAGE_FOLDER = "payslip_images"
os.makedirs(IMAGE_FOLDER, exist_ok=True)

SPREADSHEET_ID = "1A5pcalEwbRPKmOynBNFf5Y2CXSTq70WToV0VSZ21G9k"
MONTHS = list(calendar.month_name)[1:]

# ================= GOOGLE SHEETS =================
sheet_creds = Credentials.from_service_account_file(
    "gsjs.json",
    scopes=["https://www.googleapis.com/auth/spreadsheets"]
)
client = gspread.authorize(sheet_creds)
spreadsheet = client.open_by_key(SPREADSHEET_ID)

employee_sheet = spreadsheet.get_worksheet(0)
attendance_sheet = spreadsheet.get_worksheet(1)
history_sheet = spreadsheet.get_worksheet(2)

# ================= GOOGLE DRIVE =================
flow = Flow.from_client_secrets_file(
    "client_secret.json",
    scopes=["https://www.googleapis.com/auth/drive.file"],
    redirect_uri="http://127.0.0.1:5000/oauth2callback"
)

def get_drive_service():
    if "drive_creds" not in session:
        return None
    creds = UserCredentials(**session["drive_creds"])
    return build("drive", "v3", credentials=creds)

# ================= HELPERS =================
def safe_get_all(sheet):
    for _ in range(3):
        try:
            return sheet.get_all_records()
        except:
            time.sleep(2)
    return []

def get_logo_base64():
    logo_file = r"C:\Users\diori\OneDrive\Desktop\payslip\static\logo.jpg"
    with open(logo_file, "rb") as f:
        return base64.b64encode(f.read()).decode("utf-8")

def payslip_sent(emp_id, month):

    history = safe_get_all(history_sheet)

    for record in history:
        if (
            str(record.get("emp_id", "")).strip() == str(emp_id).strip()
            and record.get("month", "").strip() == month.strip()
        ):
            return True

    return False

def get_employee(emp_id):
    return next((e for e in safe_get_all(employee_sheet) if str(e["emp_id"]) == str(emp_id)), None)

def get_attendance(emp_id):
    return next((a for a in safe_get_all(attendance_sheet) if str(a["emp_id"]) == str(emp_id)), None)

def safe_filename(name):
    return "".join(c for c in name if c.isalnum() or c in (" ", "_")).replace(" ", "_")

def calc_value(enabled, value, basic):
    if not enabled:
        return 0
    v = float( value or  0)
    return round((basic * v) / 100, 2) if v <= 100 else round(v, 2)
def get_or_create_month_folder(folder_name):
    drive = get_drive_service()
    if not drive:
        raise Exception("Drive not authenticated")

    query = (
        f"'{DRIVE_FOLDER_ID}' in parents and "
        f"name='{folder_name}' and "
        f"mimeType='application/vnd.google-apps.folder' and trashed=false"
    )

    results = drive.files().list(q=query, fields="files(id, name)").execute()
    files = results.get("files", [])

    if files:
        return files[0]["id"]

    folder_metadata = {
        "name": folder_name,
        "mimeType": "application/vnd.google-apps.folder",
        "parents": [DRIVE_FOLDER_ID]
    }

    folder = drive.files().create(body=folder_metadata, fields="id").execute()
    return folder["id"]
def upload_to_drive(filepath, filename, period):
    drive = get_drive_service()
    if not drive:
        raise Exception("Drive not authenticated")

    month_folder_id = get_or_create_month_folder(period)

    media = MediaFileUpload(filepath, mimetype="image/png")
    file = drive.files().create(
        body={"name": filename, "parents": [month_folder_id]},
        media_body=media,
        fields="id"
    ).execute()

    file_id = file["id"]
    link = f"https://drive.google.com/file/d/{file_id}/view"
    return file_id, link
def send_email(to_email, emp_name, month, filepath, filename):

    msg = MIMEMultipart()
    msg["From"] = SENDER_EMAIL
    msg["To"] = to_email
    msg["Subject"] = f"Payslip - {month}"

    body = f"""Dear {emp_name},

Please find attached your payslip for {month}.

Regards,
HR Team
"""
    msg.attach(MIMEText(body, "plain"))

    with open(filepath, "rb") as f:
        img_data = f.read()
        image = MIMEImage(img_data, name=filename)
        image["Content-Disposition"] = f'attachment; filename="{filename}"'
        msg.attach(image)

    with smtplib.SMTP_SSL("smtp.gmail.com", 465) as server:
        server.login(SENDER_EMAIL, SENDER_PASSWORD)
        server.send_message(msg)
def parse_date_safely(date_string):
    formats = [
        "%Y-%m-%d",
        "%d-%m-%Y",
        "%d/%m/%Y",
        "%m/%d/%Y",
        "%Y-%m-%d %H:%M:%S",
    ]

    for fmt in formats:
        try:
            return datetime.strptime(date_string, fmt).date()
        except:
            continue

    print("Date parse failed for:", date_string)
    return None

def replace_drive_file(file_id, filepath):
    drive = get_drive_service()
    if not drive:
        raise Exception("Drive not authenticated")

    media = MediaFileUpload(
        filepath,
        mimetype="image/png",
        resumable=True
    )

    drive.files().update(
        fileId=file_id,
        media_body=media
    ).execute()
def extract_file_id(drive_link):
    # Example link:
    # https://drive.google.com/file/d/FILE_ID/view?usp=sharing

    if "/d/" in drive_link:
        return drive_link.split("/d/")[1].split("/")[0]
    return None
def get_payslip_record(emp_id, period):
    records = history_sheet.get_all_records()

    for row in records:
        if (
            str(row.get("emp_id", "")).strip() == str(emp_id).strip()
            and str(row.get("month", "")).strip() == str(period).strip()
        ):
            return row

    return None
def get_previous_pay_period(ym_value):
    year, month_num = ym_value.split("-")
    year = int(year)
    month_num = int(month_num)

    if month_num == 1:
        prev_month_num = 12
        prev_year = year - 1
    else:
        prev_month_num = month_num - 1
        prev_year = year

    month_name = calendar.month_name[prev_month_num]
    return f"{month_name} {prev_year}"
def amount_in_words(amount):
    try:
        rupees = int(amount)
        paise = int(round((amount - rupees) * 100))

        words = num2words(rupees, lang="en_IN").title() + " Rupees"
        if paise > 0:
            words += " and " + num2words(paise, lang="en_IN").title() + " Paise"
        return words + " Only"
    except:
        return ""
# ================= LOGIN =================
@app.route("/login")
def login():
    url, _ = flow.authorization_url(prompt="consent")
    return redirect(url)

@app.route("/oauth2callback")
def oauth2callback():
    flow.fetch_token(authorization_response=request.url)
    creds = flow.credentials
    session["drive_creds"] = {
        "token": creds.token,
        "refresh_token": creds.refresh_token,
        "token_uri": creds.token_uri,
        "client_id": creds.client_id,
        "client_secret": creds.client_secret,
        "scopes": creds.scopes
    }
    return redirect("/")

@app.route("/", methods=["GET", "POST"])
def index():
    employees = safe_get_all(employee_sheet)
    available = []

    current_month = date.today().strftime("%Y-%m")
    selected_period = current_month

    if request.method == "POST":
        selected_period = request.form.get("pay_period") or current_month

        if selected_period:
            year, month_num = selected_period.split("-")
            year = int(year)
            month_num = int(month_num)
            period = get_previous_pay_period(selected_period)

            period_start = date(int(year), int(month_num), 1)
            last_day = calendar.monthrange(int(year), int(month_num))[1]
            period_end = date(int(year), int(month_num), last_day)

            for emp in employees:
                emp_id = str(emp.get("emp_id", "")).strip()
                if not emp_id:
                    continue

                doj_str = str(emp.get("date_of_joining", "")).strip()
                if doj_str:
                    doj = parse_date_safely(doj_str)
                    if doj and doj > period_end:
                        continue

                rel_str = str(emp.get("relieving_date", "")).strip()
                if rel_str:
                    rel = parse_date_safely(rel_str)
                    if rel and rel < period_start:
                        continue

                if payslip_sent(emp_id, period):
                    continue

                available.append(emp)

    return render_template(
        "index.html",
        employees=available,
        selected_period=selected_period
    )
# ================= GENERATE =================
@app.route("/payslip", methods=["POST"])
def generate_payslip():
    drive = get_drive_service()
    if not drive:
        return redirect(url_for("login"))
    file_id = request.form.get("file_id")
    single_emp_id = request.form.get("emp_id")

    # Determine employee list and period
    if single_emp_id:
        emp_ids = [single_emp_id]
        period = request.form.get("period")  # Expecting "Month Year"
    else:
        emp_ids = request.form.getlist("emp_ids")
        pay_period = request.form.get("pay_period")  # "YYYY-MM"
        if not emp_ids or not pay_period:
            return "Missing data", 400

        period = get_previous_pay_period(pay_period)

    payslips = []

    for emp_id in emp_ids:
        # Fetch employee data
        if single_emp_id:
            emp = {
                "emp_id": emp_id,
                "name": request.form["name"],
                "mail": request.form["mail"],
                "basic": request.form["basic"],
                "bank":request.form["bank"],
                "phone":request.form["phone"],
                "acc_no":request.form["acc_no"]
            }
        else:
            emp = get_employee(emp_id)

        att = get_attendance(emp_id)
        if not emp or not att:
            continue

        # Salary calculation
        basic = float(emp["basic"])
        total_days = int(att["total_days"])
        present_days = int(att["present_days"])

        cl_days = int(att.get("cl", 0) or 0) if request.form.get("cl_enabled") else 0

        absent_days = max(total_days - present_days - cl_days, 0)
        loss_of_pay = round(absent_days * (basic / total_days), 2)
        

        hra = calc_value(request.form.get("hra_enabled"), request.form.get("hra_value"), basic)
        da = calc_value(request.form.get("da_enabled"), request.form.get("da_value"), basic)
        ita = calc_value(request.form.get("ita_enabled"), request.form.get("ita_value"), basic)
        pf = calc_value(request.form.get("pf_enabled"), request.form.get("pf_value"), basic)
        esi = calc_value(request.form.get("esi_enabled"), request.form.get("esi_value"), basic)

        gross_salary = round(basic + hra + da + ita, 2)
        net_salary = round(gross_salary - (pf + esi + loss_of_pay), 2)
        total_earnings = round(basic + hra + da + ita, 2)
        total_deductions = round(pf + esi + loss_of_pay, 2)
        net_salary_words = amount_in_words(net_salary)
        print("ATT DATA:", att)

        # Create payslip dict
        payslip_data = {
            "emp_id": emp["emp_id"],
            "name": emp["name"],
            "mail": emp["mail"],
            "bank": emp["bank"],
            "phone": emp["phone"],
            "acc_no": emp["acc_no"],
            "basic": basic,
            "hra": hra,
            "da": da,
            "ita": ita,
            "pf": pf,
            "esi": esi,
            "loss_of_pay": loss_of_pay,
            "gross_salary": gross_salary,
            "net_salary": net_salary,
            "month": period,
            "today": date.today().strftime("%d-%m-%Y"),
            "present_days": present_days,
            "absent_days": absent_days,
            "cl": cl_days,
            "total_days": total_days,
            "file_id": file_id,
            "total_earnings": total_earnings,
            "total_deductions": total_deductions,
            "net_salary_words": net_salary_words,
            
        }

        payslips.append(payslip_data)  # ✅ Only append once

        # 🔹 Generate PDF using the same dict
        filename = f"{safe_filename(emp['name'])}_{period}.png"
        filepath = os.path.join(IMAGE_FOLDER, filename)
        logo_base64 = get_logo_base64()

        html = render_template(
            "payslip_pdf.html",
            payslips=[payslip_data],
            logo_base64=logo_base64
        )
        imgkit.from_string(html, filepath, config=img_config, options=options)
        # 🔹 Upload / Replace in Drive
        payslip_data["file_name"] = filename
        payslip_data["file_path"] = filepath
    # Save in session
    session["generated"] = {"period": period, "payslips": payslips,"file_id": file_id}
    logo_base64 = get_logo_base64()
    return render_template("pay1.html", payslips=payslips, period=period, logo_base64=logo_base64)
# ================= SEND =================
@app.route("/send_all_payslips", methods=["POST"])
def send_all_payslips():

    if "drive_creds" not in session:
        return redirect("/login")

    data = session.get("generated")

    if not data or "payslips" not in data:
        return redirect("/")

    period = data["period"]
    payslips = data["payslips"]

    history_data = history_sheet.get_all_records()

    for slip in payslips:

        emp_id = slip["emp_id"]
        emp = get_employee(emp_id)

        if not emp:
            continue

        today = date.today()
        logo_base64 = get_logo_base64()

        html = render_template(
            "payslip_pdf.html",
            payslips=[{
                "emp_id": slip["emp_id"],
                "name": slip["name"],
                "mail": slip.get("mail", ""),
                "bank": slip.get("bank", ""),
                "acc_no": slip.get("acc_no", ""),
                "phone": slip.get("phone", ""),
                "basic": slip["basic"],
                "hra": slip["hra"],
                "da": slip["da"],
                "ita": slip["ita"],
                "pf": slip["pf"],
                "esi": slip["esi"],
                "loss_of_pay": slip["loss_of_pay"],
                "gross_salary": slip["gross_salary"],
                "net_salary": slip["net_salary"],
                "present_days": slip.get("present_days", 0),
                "absent_days": slip.get("absent_days", 0),
                "cl": slip.get("cl", 0),
                "total_days": slip.get("total_days", 0),
                "month": period,
                "today": today.strftime("%d-%m-%Y"),
                "total_earnings": slip.get("total_earnings", 0),
                "total_deductions": slip.get("total_deductions", 0),
                "net_salary_words": slip.get("net_salary_words", "")
            }],
            logo_base64=logo_base64
        )

        filename = f"{safe_filename(emp['name'])}_Payslip_{period.replace(' ', '_')}.png"
        filepath = os.path.join(IMAGE_FOLDER, filename)

        imgkit.from_string(html, filepath, config=img_config, options=options)
        file_id_from_edit = slip.get("file_id")

        # ================= FIND EXISTING RECORD =================

        existing_row_index = None
        existing_link = None
        existing_file_id = None

        for idx, row in enumerate(history_data, start=2):

            if (
                str(row.get("emp_id", "")).strip() == str(emp_id).strip()
                and str(row.get("month", "")).strip().lower() == str(period).strip().lower()
            ):

                existing_row_index = idx
                existing_link = row.get("pdf_link", "")
                existing_file_id = row.get("file_id", "")
                break

        # ================= DRIVE UPLOAD / REPLACE =================

        try:
            drive_file_id = None

            if file_id_from_edit and str(file_id_from_edit).strip():
                drive_file_id = str(file_id_from_edit).strip()

            elif existing_file_id and str(existing_file_id).strip():
                drive_file_id = str(existing_file_id).strip()

            elif existing_link and str(existing_link).strip():
                drive_file_id = extract_file_id(existing_link)

            if drive_file_id:
                print("REPLACING EXISTING FILE")
                replace_drive_file(drive_file_id, filepath)
                file_id = drive_file_id
                link = existing_link if existing_link else f"https://drive.google.com/file/d/{file_id}/view"
            else:
                print("CREATING NEW FILE")
                file_id, link = upload_to_drive(filepath, filename, period)

        except Exception as e:
            print("Drive Upload/Replace Error:", e)
            continue

        # ================= SEND EMAIL =================

        try:

            send_email(
                slip["mail"],
                slip["name"],
                period,
                filepath,
                filename
            )

            print("Mail Sent:", slip["mail"])

        except Exception as e:

            print("Mail Error:", e)
            continue

        # ================= UPDATE GOOGLE SHEET =================

        if existing_row_index:

            history_sheet.update(
                [[
                    slip["name"],
                    period,
                    slip["gross_salary"],
                    link,
                    file_id,
                    slip.get("cl", 0),
                    datetime.now().strftime("%d-%m-%Y %H:%M"),
                    "Sent"
                ]],
                f"B{existing_row_index}:H{existing_row_index}"
            )

        else:

            history_sheet.append_row([
                emp_id,
                slip["name"],
                period,
                slip["gross_salary"],
                link,
                file_id,
                #slip.get("cl", 0),
                datetime.now().strftime("%d-%m-%Y %H:%M"),
                "Sent"
            ])

    return redirect("/history")
# ================= HISTORY =================
@app.route("/history")
def history_page():
    year_filter = request.args.get("year")
    records = history_sheet.get_all_records()
    employees = {}

    for r in records:
        # month column format: "March 2026"
        try:
            month, yr = r["month"].split()
        except ValueError:
            continue

        if year_filter and yr != year_filter:
            continue

        emp_id = r["emp_id"]

        if emp_id not in employees:
            employees[emp_id] = {
                "emp_id": emp_id,
                "emp_name": r["emp_name"],
                "months": {m: None for m in MONTHS}
            }

        employees[emp_id]["months"][month] = {
            "salary": r.get("gross") or r.get("basic") or "",
            "link": r.get("pdf_link", ""),
            "year": yr
        }

    years = sorted({r["month"].split()[1] for r in records if "month" in r and r["month"] and len(r["month"].split()) > 1})

    return render_template(
        "history.html",
        employees=employees.values(),
        months=MONTHS,
        years=years,
        selected_year=year_filter
    )
@app.route("/edit_payslip")
def edit_payslip():

    emp_id = request.args.get("emp_id")
    month = request.args.get("month")
    year = request.args.get("year")
    if not emp_id or not month or not year:
        return "Missing edit parameters", 400

    period = f"{month} {year}"

    emp = get_employee(emp_id)

    if not emp:
        return "Employee not found", 404
    # Fetch payslip row from sheet
    payslip_data = get_payslip_record(emp_id, period)

    if not payslip_data:
        return "Payslip not found", 404

    # 🔥 Extract file_id from stored pdf_link
    file_id = extract_file_id(payslip_data["pdf_link"])

    return render_template(
        "editpayslip.html",
        emp=emp,
        period=period,
        payslip=payslip_data,
        file_id=file_id   
    )   
@app.route("/update_payslip", methods=["POST"])
def update_payslip():

    if "drive_creds" not in session:
        return redirect("/login")

    emp_id = request.form["emp_id"]
    period = request.form["period"]
    file_id = request.form["file_id"]

    # ================= GET FORM VALUES =================
    name = request.form["name"]
    department = request.form["department"]
    phone = request.form["phone"]
    mail = request.form["mail"]
    bank = request.form["bank"]
    acc_no = request.form["acc_no"]
    basic = float(request.form["basic"])

    # ================= GET ATTENDANCE =================
    att = get_attendance(emp_id)
    if not att:
        return "Attendance not found", 404

    total_days = int(att["total_days"])
    present_days = int(att["present_days"])
    cl_days = int(att.get("cl", 0) or 0) if request.form.get("cl_enabled") else 0

    absent_days = max(total_days - present_days - cl_days, 0)
    loss_of_pay = round(absent_days * (basic / total_days), 2)

    # ================= ALLOWANCES / DEDUCTIONS =================
    hra = calc_value(request.form.get("hra_enabled"), request.form.get("hra_value"), basic)
    da = calc_value(request.form.get("da_enabled"), request.form.get("da_value"), basic)
    ita = calc_value(request.form.get("ita_enabled"), request.form.get("ita_value"), basic)
    pf = calc_value(request.form.get("pf_enabled"), request.form.get("pf_value"), basic)
    esi = calc_value(request.form.get("esi_enabled"), request.form.get("esi_value"), basic)

    gross_salary = round(basic + hra + da + ita, 2)
    net_salary = round(gross_salary - (pf + esi + loss_of_pay), 2)
    total_earnings = round(basic + hra + da + ita, 2)
    total_deductions = round(pf + esi + loss_of_pay, 2)
    net_salary_words = amount_in_words(net_salary)

    # ================= UPDATE EMPLOYEE SHEET =================
    employee_data = employee_sheet.get_all_records()

    for idx, row in enumerate(employee_data):
        if str(row["emp_id"]).strip() == str(emp_id).strip():
            employee_sheet.update(
                f"A{idx+2}:H{idx+2}",
                [[emp_id, name, department, phone, mail, bank, acc_no, basic]]
            )
            break

    # ================= GENERATE UPDATED IMAGE =================
    today = date.today()

    logo_base64 = get_logo_base64()

    html = render_template(
        "payslip_pdf.html",
        payslips=[{
            "emp_id": emp_id,
            "name": name,
            "department": department,
            "phone": phone,
            "mail": mail,
            "bank": bank,
            "acc_no": acc_no,
            "basic": basic,
            "month": period,
            "today": today.strftime("%d-%m-%Y"),
            "hra": hra,
            "da": da,
            "ita": ita,
            "pf": pf,
            "esi": esi,
            "loss_of_pay": loss_of_pay,
            "gross_salary": gross_salary,
            "net_salary": net_salary,
            "present_days": present_days,
            "absent_days": absent_days,
            "cl": cl_days,
            "total_days": total_days,
            "total_earnings": total_earnings,
            "total_deductions": total_deductions,
            "net_salary_words": net_salary_words,
        }],
        logo_base64=logo_base64
    )

    filename = f"{safe_filename(name)}_Payslip_{period.replace(' ', '_')}.png"
    filepath = os.path.join(IMAGE_FOLDER, filename)

    imgkit.from_string(html, filepath, config=img_config, options=options)
    # ================= REPLACE FILE IN GOOGLE DRIVE =================
    replace_drive_file(file_id, filepath)
    history_records = history_sheet.get_all_records()

    for idx, row in enumerate(history_records, start=2):
        if (
            str(row.get("emp_id", "")).strip() == str(emp_id).strip()
            and str(row.get("month", "")).strip() == str(period).strip()
        ):
            history_sheet.update(
                [[
                    emp_id,
                    name,
                    period,
                    gross_salary,
                    row.get("pdf_link", ""),
                    row.get("file_id", file_id),
                    datetime.now().strftime("%d-%m-%Y %H:%M"),
                    "Updated"
                ]],
                f"A{idx}:H{idx}"
            )
            break

    return redirect("/history")
def generate_updated_payslip():

    emp_id = request.form["emp_id"]
    period = request.form["period"]

    name = request.form["name"]
    department = request.form["department"]
    phone = request.form["phone"]
    mail = request.form["mail"]
    bank = request.form["bank"]
    acc_no = request.form["acc_no"]
    basic = float(request.form["basic"])

    today = date.today()

    html = render_template(
        "payslip_pdf.html",
        payslips=[{
            "emp_id": emp_id,
            "name": name,
            "basic": basic,
            "department": department,
            "phone": phone,
            "mail": mail,
            "bank": bank,
            "acc_no": acc_no,
            "month": period,
            "today": today.strftime("%d-%m-%Y")
        }]
    )

    filename = f"{safe_filename(name)}_Payslip_{period.replace(' ', '_')}.pdf"
    filepath = os.path.join(PDF_FOLDER, filename)

    pdfkit.from_string(html, filepath, configuration=config)
    # 🔥 VERY IMPORTANT PART
    file_id = request.form["file_id"]   # you must send this from HTML

    replace_drive_file(file_id, filepath)
    return send_file(filepath, as_attachment=True)

# ================= RUN =================
if __name__ == "__main__":
    app.run(debug=True)