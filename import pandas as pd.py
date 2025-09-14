import pandas as pd
from fpdf import FPDF

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
CARDS_PER_COL = 6   # 1 × 6 = 6 admit cards per page
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

# Add Unicode font (DejaVuSans)
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

    # Write details inside card
    pdf.set_xy(x + 3, y + 3)
    pdf.multi_cell(
        CARD_WIDTH - 6, 5,
        f"Name: {row['Full Name (as per certificate)']}\n"
        f"Dept: {row['Department']}\n"
        f"Faculty: {row['Faculty']}\n"
        f"Phone: {row['Phone Number (WhatsApp preferred)']}\n"
        f"Email: {row['Email Address (Please double check and submit correct email address)']}\n"
        f"Roll: {row['Roll No']}",
        align='L'
    )

    # Move to next column
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

print("✅ Combined Admit Card PDF generated with 6 per A4 page!")
