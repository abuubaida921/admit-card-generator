# Admit Card Generator

Generate professional admit cards with QR codes for events, exams, or institutions. Supports Bengali fonts, department/faculty grouping, and custom logos.

## Features
- Bulk admit card generation from Excel (`students.xlsx`)
- QR code on each card for quick attendance (JSON data)
- Bengali font support (e.g., `kalpurush.ttf`)
- Custom event and organization logos
- Roll number auto-generation (10 digits)
- Cards grouped by faculty and department
- PDF output (A4, multiple cards per page)

## Requirements
- Python 3.8+
- pandas
- reportlab
- qrcode
- pillow

## Setup
1. Clone the repository:
   ```bash
   git clone https://github.com/abuubaida921/admit-card-generator.git
   cd admit-card-generator
   ```
2. Install dependencies:
   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install pandas reportlab qrcode pillow
   ```
3. Place the following files in the project folder:
   - `students.xlsx` (student data)
   - `kalpurush.ttf` (Bengali font)
   - `event_logo.jpg` (event logo)
   - `org_logo.png` (organization logo)

## Usage
1. Edit `students.xlsx` with columns:
   - `Full Name (as per certificate)`
   - `Department`
   - `Phone Number (WhatsApp preferred)`
   - `Email Address (Please double check and submit correct email address)`
   - (Other columns as needed)
2. Run the generator:
   ```bash
   python generate_admit_cards.py
   ```
3. Find the output PDF:
   - `Admit_Cards_All.pdf`

## QR Code Attendance
- Scan QR codes using any Android QR scanner app
- Data includes name, roll, department, phone (JSON)
- Export scanned data for attendance tracking

## Customization
- Change logos, fonts, or layout in `generate_admit_cards.py`
- Adjust card size, number per page, or roll number format as needed

## License
This project is open source for educational and event use. Font files may have their own licenses.

## Author
Developed by abuubaida921

---
For issues or feature requests, open an issue on GitHub.
