import os
import json
import tempfile
import pandas as pd
import qrcode

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# === CONFIGURATION ===
BASE_ROLL = 5017010   # Roll number prefix
FACULTY_DEPARTMENTS = {
    "Faculty of Science and Engineering": [
        "Mathematics",
        "Computer Science & Engineering",
        "Chemistry",
        "Physics",
        "Geology and Mining",
        "Statistics",
    ],
    "Faculty of Bio-Sciences": [
        "Soil and Environmental Sciences",
        "Botany",
        "Coastal Studies and Disaster Management",
        "Biochemistry and Biotechnology",
    ],
    "Faculty of Business Studies": [
        "Management Studies",
        "Accounting and Information Systems",
        "Marketing",
        "Finance and Banking",
    ],
    "Faculty of Social Sciences": [
        "Economics",
        "Political Science",
        "Sociology",
        "Public Administration",
        "Mass Communication & Journalism",
        "Social Work",
    ],
    "Faculty of Arts and Humanities": [
        "Bangla",
        "English",
        "Philosophy",
        "History",
    ],
    "Faculty of Law": [
        "Law",
    ],
}

CARDS_PER_ROW = 1
CARDS_PER_COL = 5   # 4 admit cards per page
CARD_WIDTH = 100 * mm
CARD_HEIGHT = 50 * mm

# === Load Excel ===
df = pd.read_excel("students.xlsx")

# Add Faculty Column
def get_faculty(dept):
    for faculty, dept_list in FACULTY_DEPARTMENTS.items():
        if dept in dept_list:
            return faculty
    return "Others"

df["Faculty"] = df["Department"].apply(get_faculty)
# Now sort by Faculty, Department, and Name
df = df.sort_values(["Faculty", "Department", "Full Name (as per certificate)"])

# Assign Roll Numbers with gap of 20 after each faculty
roll_numbers = {}
serial = 1
for faculty, faculty_df in df.groupby("Faculty", sort=False):
    for idx in faculty_df.index:
        roll_no = f"{BASE_ROLL}{serial:04d}"
        roll_numbers[idx] = roll_no
        serial += 1
    serial += 20  # gap after each faculty

df["Roll No"] = df.index.map(roll_numbers)

# === PDF CONFIG ===
FONT_PATH = os.path.join(os.path.dirname(__file__), "kalpurush.ttf")
if not os.path.exists(FONT_PATH):
    raise FileNotFoundError("kalpurush.ttf not found in project directory")

pdfmetrics.registerFont(TTFont("Kalpurush", FONT_PATH))

event_logo_path = os.path.join(os.path.dirname(__file__), "event_logo.jpg")
org_logo_path = os.path.join(os.path.dirname(__file__), "org_logo.png")

# === Create PDF ===
c = canvas.Canvas("Admit_Cards_All.pdf", pagesize=A4)
page_width, page_height = A4

x_margin, y_margin = 10 * mm, 10 * mm
x, y = x_margin, page_height - y_margin - CARD_HEIGHT
count = 0

for _, row in df.iterrows():
    # Draw rectangle
    c.setStrokeColorRGB(0, 0, 0)
    c.rect(x, y, CARD_WIDTH, CARD_HEIGHT)

    # Logos
    if os.path.exists(event_logo_path):
        c.drawImage(event_logo_path, x + 2*mm, y + CARD_HEIGHT - 17*mm, 30*mm, 15*mm, preserveAspectRatio=True, mask="auto")
    if os.path.exists(org_logo_path):
        c.drawImage(org_logo_path, x + CARD_WIDTH - 17*mm, y + CARD_HEIGHT - 17*mm, 15*mm, 15*mm, preserveAspectRatio=True, mask="auto")

    # Phone Format
    visible_phone = str(row.get('Phone Number (WhatsApp preferred)', '')).strip()
    if visible_phone and len(visible_phone) == 10 and not visible_phone.startswith('0'):
        visible_phone = '0' + visible_phone
    elif visible_phone and len(visible_phone) == 11 and not visible_phone.startswith('0'):
        visible_phone = '0' + visible_phone[1:]

    # Student Info Text
    text_obj = c.beginText()
    text_obj.setTextOrigin(x + 5*mm, y + CARD_HEIGHT - 25*mm)
    text_obj.setFont("Kalpurush", 9)

    text_obj.textLine(f"Name: {row['Full Name (as per certificate)']}")
    text_obj.textLine(f"Dept: {row['Department']}")
    text_obj.textLine(f"Faculty: {row['Faculty']}")
    text_obj.textLine(f"Phone: {visible_phone}")
    text_obj.textLine(f"Email: {row['Email Address (Please double check and submit correct email address)']}")
    text_obj.textLine(f"Roll/Reg. ID: {row['Roll No']}")

    c.drawText(text_obj)

    # === QR Code ===
    phone = str(row.get('Phone Number (WhatsApp preferred)', '')).strip()
    if phone and len(phone) == 10 and not phone.startswith('0'):
        phone = '0' + phone
    elif phone and len(phone) == 11 and not phone.startswith('0'):
        phone = '0' + phone[1:]

    student_info = {
        "name": str(row.get('Full Name (as per certificate)', '')),
        "roll": str(row.get('Roll No', '')),
        "department": str(row.get('Department', '')),
        "phone": phone
    }

    qr_data = json.dumps(student_info, ensure_ascii=False)
    qr = qrcode.QRCode(box_size=2, border=1)
    qr.add_data(qr_data)
    qr.make(fit=True)
    img = qr.make_image(fill_color="black", back_color="white")

    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp_qr:
        img.save(tmp_qr, format="PNG")
        tmp_qr_path = tmp_qr.name

    # Enlarge QR code by 5% (from 20mm to 21mm)
    qr_size = 21 * mm
    c.drawImage(tmp_qr_path, x + CARD_WIDTH - qr_size - 4*mm, y + 6*mm, qr_size, qr_size, preserveAspectRatio=True, mask="auto")
    os.remove(tmp_qr_path)

    # Next card position
    count += 1
    if count % CARDS_PER_ROW == 0:
        x = x_margin
        y -= CARD_HEIGHT + 5*mm
    else:
        x += CARD_WIDTH + 5*mm

    # New page if full
    if count % (CARDS_PER_ROW * CARDS_PER_COL) == 0:
        c.showPage()
        x, y = x_margin, page_height - y_margin - CARD_HEIGHT

c.save()
print("✅ Combined Admit Card PDF generated with QR codes (ReportLab version)!")
