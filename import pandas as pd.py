import pandas as pd
from fpdf import FPDF
import qrcode
from io import BytesIO

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
CARDS_PER_COL = 6   # 6 admit cards per page
CARD_WIDTH = 100
CARD_HEIGHT = 40

# === Load Excel ===
df = pd.read_excel("students.xlsx")

# Add Faculty Column
def get_faculty(dept):
    for faculty, dept_list in FACULTY_DEPARTMENTS.items():
        if dept in dept_list:
            return faculty
    return "Others"

df["Faculty"] = df["Department"].apply(get_faculty)

# Assign Roll Numbers with gap of 20 after each faculty
roll_numbers = []
serial = 1
for faculty, faculty_df in df.groupby("Faculty"):
    for _ in faculty_df.itertuples():
        roll_no = f"{BASE_ROLL}{serial:04d}"
        roll_numbers.append((_.Index, roll_no))
        serial += 1
    serial += 20  # gap after each faculty

df["Roll No"] = df.index.map(dict(roll_numbers))

# === Create PDF with multiple admit cards per page ===
import os
FONT_PATH = os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf")
if not os.path.exists(FONT_PATH):
    raise FileNotFoundError("DejaVuSans.ttf font file not found in project directory. Please add it for Unicode support.")
pdf = FPDF("P", "mm", "A4")
pdf.add_page()
pdf.add_font("DejaVu", "", FONT_PATH, uni=True)
pdf.set_font("DejaVu", '', 9)

x_margin, y_margin = 10, 10
x, y = x_margin, y_margin
count = 0

for _, row in df.iterrows():
    # Draw rectangle for card
    pdf.rect(x, y, CARD_WIDTH, CARD_HEIGHT)

    # Write student details (leave space for QR)
    pdf.set_xy(x + 3, y + 3)
    # Format phone for visible text as well
    visible_phone = str(row.get('Phone Number (WhatsApp preferred)', ''))
    if visible_phone and len(visible_phone) == 10 and not visible_phone.startswith('0'):
        visible_phone = '0' + visible_phone
    elif visible_phone and len(visible_phone) == 11 and not visible_phone.startswith('0'):
        visible_phone = '0' + visible_phone[1:]

    pdf.multi_cell(
        CARD_WIDTH - 28, 5,
        f"Name: {row['Full Name (as per certificate)']}\n"
        f"Dept: {row['Department']}\n"
        f"Faculty: {row['Faculty']}\n"
        f"Phone: {visible_phone}\n"
        f"Email: {row['Email Address (Please double check and submit correct email address)']}\n"
        f"Roll: {row['Roll No']}",
        align='L'
    )

    # === Generate QR code with all student info as JSON ===
    import tempfile, json
    # Ensure phone number starts with '0' and is 11 digits
    phone = str(row.get('Phone Number (WhatsApp preferred)', ''))
    if phone and len(phone) == 10 and not phone.startswith('0'):
        phone = '0' + phone
    elif phone and len(phone) == 11 and not phone.startswith('0'):
        phone = '0' + phone[1:]
    # Optionally, you can add more validation for phone numbers here

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

    # Insert QR image (20x20mm) inside the card
    pdf.image(tmp_qr_path, x + CARD_WIDTH - 25, y + 5, 20, 20)

    # Move to next column / row
    count += 1
    if count % CARDS_PER_ROW == 0:
        x = x_margin
        y += CARD_HEIGHT + 5
    else:
        x += CARD_WIDTH + 5

    # New page if full
    if count % (CARDS_PER_ROW * CARDS_PER_COL) == 0:
        pdf.add_page()
        x, y = x_margin, y_margin

# Save final PDF
pdf.output("Admit_Cards_All.pdf")

print("✅ Combined Admit Card PDF generated with QR codes (6 per page)!")
