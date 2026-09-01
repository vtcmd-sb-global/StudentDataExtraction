STUDENT DATA EXTRACTION – USER GUIDE (GUI Version)

FOLDER STRUCTURE
----------------
DataExtraction/
│
├── source/               ← Put ANY named source Excel files here
│   ├── BookingConfirmationRegisterReport.xlsx   (or any name)
│   └── StudentPersonalDetailReport.xlsx         (or any name)
│
├── template/             ← Put the destination TEMPLATE here
│   └── Client  Data Collection File.xlsx        (or any name)
│       (This file is NEVER modified – it is only used as a pattern)
│
├── destination/          ← Generated result files appear here
│   └── Client_Data_Collection_YYYYMMDD_HHMMSS.xlsx
│
├── data_extraction.py   ← Main program (GUI + logic)
├── Run_Data_Extraction.bat         ← Double-click to launch GUI
├── requirements.txt
└── matching_log.txt              (created after each run)


HOW TO USE (GUI)
----------------
1. Place one or more Booking Confirmation report(s) into  source/
2. Place one or more Student Personal Detail report(s) into  source/
3. Place the Client Data Collection template into  template/
4. Double-click  Run_Data_Extraction.bat   (or run: python consolidate_student_data.py)
5. The GUI window opens:
   - Click “Refresh” if you just added files
   - Check the detected Booking / Personal files lists
   - Click “Run Data Extraction”
   - Watch the live log and progress bar
   - When finished, click “Open Output File” or open the destination folder
6. The result appears in destination/ with a timestamp so nothing is overwritten


ALTERNATIVE – COMMAND LINE
--------------------------
You can still run the original console mode:
  python consolidate_student_data.py --cli


IMPORTANT RULES
---------------
• File names can be ANYTHING – the program auto-detects which
  files are Booking reports and which are Personal reports by
  reading their headers.

• You can put MULTIPLE Booking files and MULTIPLE Personal files
  in source/ at the same time. All of them will be processed.

• NO DUPLICATE CHECKING – every row from every Booking file is
  written to the result, even if the same student name appears
  many times. Every student’s data is kept.

• The template file in template/ is never changed. A fresh copy
  is always written into destination/ with a timestamp in the
  filename so older results are never overwritten.


COLUMN MAPPING
-------------------------------------------------
Destination Column                  Source
-------------------------------     --------------------------------
Studentname                         Booking → Student Name (titles removed)
Fathername                          Personal → Guardian Name / Father Name
Dateofbirth                         Personal → Birth date
Familynumber                        Booking → Course Code
Admissiondate                       Booking → Booking Confirmation Date
Genderid                            Personal → Title (1=Male, 2=Female)
Admissionnumber/ Registration No.   Booking → Revised or Original BC Number
Rollnumber                          Booking → Enrollment No.
Mobilenumber                        Personal → Mobile No.
Studentaddress                      Personal → Address 1 + Address 2
Class / Section / Guardiancnic /
Tuition Fee                         Booking → Monthly Fee After Discount
Previous Balance                    left empty


REQUIREMENTS
------------
• Python 3.8 or higher
• openpyxl  (install with:  pip install openpyxl)
  or:  pip install -r requirements.txt

The GUI uses only the built-in tkinter library – no extra GUI packages needed.


TROUBLESHOOTING
---------------
• “No Excel files found in source/”
  → Make sure the .xlsx files are inside the source folder,
    not next to the script. Click Refresh in the GUI.

• “No Booking Confirmation file detected”
  → The file must contain a header with “Student Name” and
    “Enrollment” (or the word “Booking Confirmation” in the
    title area). Rename is not required.

• Matched count is low
  → Names are matched after removing titles (Mr/Ms/…) and
    converting to upper-case. Spelling differences between
    the two reports will cause unmatched rows (booking data
    is still written, just without personal details).

• GUI does not open / black window
  → Ensure Python is installed and on PATH.
    Run “python --version” in a command prompt to verify.
    Install openpyxl if missing:  pip install openpyxl


Created for flexible, name-independent use.
GUI version keeps 100 % of the original matching & mapping logic.
============================================================
