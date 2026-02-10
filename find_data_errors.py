import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
from fuzzywuzzy import fuzz

# Set up OAuth credentials
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]

def get_credentials():
    creds = None
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds = pickle.load(token)
    
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', SCOPES)
            creds = flow.run_local_server(port=0)
        
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds, token)
    
    return creds

print("Authenticating...")
creds = get_credentials()
print("Authentication successful!")

# Connect to Google Sheets
gc = gspread.authorize(creds)

# Open spreadsheet
spreadsheet = gc.open('MASTER SPRING 2026')
master_sheet = spreadsheet.worksheet('MASTER')

print(f"Reading MASTER sheet...")

# Get all data
data = master_sheet.get_all_values()
rows = data[1:]

print(f"Found {len(rows)} rows")

# Column indices
col_school = 47   # Column AV
col_student = 48  # Column AW
col_teacher = 49  # Column AX
col_grade = 50    # Column AY

# Collect student data
schools_students = {}  # {school: {student_name: count}}
all_students = {}  # {student_name: {schools: set(), grades: set(), teachers: set()}}

for row in rows:
    if len(row) > col_grade:
        school = row[col_school].strip()
        student = row[col_student].strip()
        teacher = row[col_teacher].strip() if len(row) > col_teacher else ''
        grade = row[col_grade].strip() if len(row) > col_grade else ''
        
        if not school or not student:
            continue
        
        # Track students by school
        if school not in schools_students:
            schools_students[school] = {}
        
        if student not in schools_students[school]:
            schools_students[school][student] = 0
        
        schools_students[school][student] += 1
        
        # Track all student data globally
        if student not in all_students:
            all_students[student] = {
                'schools': set(),
                'grades': set(),
                'teachers': set()
            }
        
        all_students[student]['schools'].add(school)
        if grade:
            all_students[student]['grades'].add(grade)
        if teacher:
            all_students[student]['teachers'].add(teacher)

print(f"\nFound {len(schools_students)} schools")
print(f"Found {len(all_students)} unique student names")

# Initialize Error Log
error_log_data = []
error_log_data.append(['Error Type', 'School', 'Student Name', 'Details', 'Suggestion'])

total_issues = 0

# ERROR CHECK 1: Missing last name (single word names)
print("\n" + "="*60)
print("Checking for missing last names...")
print("="*60)

for student_name in all_students.keys():
    if ' ' not in student_name.strip():
        # Only one word - missing last name
        schools_list = ', '.join(all_students[student_name]['schools'])
        total_issues += 1
        
        error_log_data.append([
            'Missing Last Name',
            schools_list,
            student_name,
            'Student name has only one word',
            'Add last name or verify if correct'
        ])
        
        print(f"  ⚠️  '{student_name}' in {schools_list}")

# ERROR CHECK 2: Same student in multiple schools
print("\n" + "="*60)
print("Checking for students in multiple schools...")
print("="*60)

for student_name, data in all_students.items():
    if len(data['schools']) > 1:
        schools_list = ', '.join(data['schools'])
        total_issues += 1
        
        error_log_data.append([
            'Multiple Schools',
            schools_list,
            student_name,
            f"Appears in {len(data['schools'])} schools",
            'Verify correct school and remove duplicates'
        ])
        
        print(f"  ⚠️  '{student_name}' appears in: {schools_list}")

# ERROR CHECK 3: Same student with different grades
print("\n" + "="*60)
print("Checking for students with multiple grades...")
print("="*60)

for student_name, data in all_students.items():
    if len(data['grades']) > 1:
        schools_list = ', '.join(data['schools'])
        grades_list = ', '.join(data['grades'])
        total_issues += 1
        
        error_log_data.append([
            'Multiple Grades',
            schools_list,
            student_name,
            f"Listed as: {grades_list}",
            'Verify correct grade'
        ])
        
        print(f"  ⚠️  '{student_name}' has multiple grades: {grades_list}")

# ERROR CHECK 4: Same student with different teachers
print("\n" + "="*60)
print("Checking for students with multiple teachers...")
print("="*60)

for student_name, data in all_students.items():
    if len(data['teachers']) > 1:
        schools_list = ', '.join(data['schools'])
        teachers_list = ', '.join(data['teachers'])
        total_issues += 1
        
        error_log_data.append([
            'Multiple Teachers',
            schools_list,
            student_name,
            f"Listed with: {teachers_list}",
            'Verify correct teacher'
        ])
        
        print(f"  ⚠️  '{student_name}' has multiple teachers: {teachers_list}")

# ERROR CHECK 5: Similar names within same school (typos/misspellings)
print("\n" + "="*60)
print("Checking for similar names (possible typos)...")
print("="*60)

for school_name, students in schools_students.items():
    student_list = list(students.items())
    
    for i in range(len(student_list)):
        for j in range(i + 1, len(student_list)):
            name1, count1 = student_list[i]
            name2, count2 = student_list[j]
            
            # Calculate similarity (0-100)
            similarity = fuzz.ratio(name1.lower(), name2.lower())
            
            # If similarity is high (but not exact match), flag it
            if 70 <= similarity < 100:
                total_issues += 1
                
                # Suggest which name to keep (the one with more orders)
                if count1 >= count2:
                    suggestion = f"Keep '{name1}' ({count1} orders), merge '{name2}' ({count2} orders)"
                else:
                    suggestion = f"Keep '{name2}' ({count2} orders), merge '{name1}' ({count1} orders)"
                
                error_log_data.append([
                    'Similar Names',
                    school_name,
                    f"{name1} / {name2}",
                    f"{similarity}% similar",
                    suggestion
                ])
                
                print(f"  ⚠️  '{name1}' ≈ '{name2}' in {school_name} ({similarity}% similar)")

# Summary
print(f"\n{'='*60}")
print(f"SUMMARY")
print(f"{'='*60}")
print(f"Total issues found: {total_issues}")
print(f"{'='*60}")

# Create or update Error Log sheet
print("\nUpdating Error Log sheet...")

try:
    error_log_sheet = spreadsheet.worksheet('Error Log')
    print("  Found existing 'Error Log' sheet - clearing it...")
    error_log_sheet.clear()
except:
    print("  Creating new 'Error Log' sheet...")
    error_log_sheet = spreadsheet.add_worksheet(title='Error Log', rows=1000, cols=10)

# Write data to Error Log
if error_log_data:
    error_log_sheet.update('A1', error_log_data)
    
    # Format header row
    error_log_sheet.format('A1:E1', {
        'backgroundColor': {'red': 0.8, 'green': 0.2, 'blue': 0.2},
        'textFormat': {
            'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
            'bold': True,
            'fontSize': 12
        },
        'horizontalAlignment': 'CENTER'
    })
    
    # Auto-resize columns
    error_log_sheet.columns_auto_resize(0, 4)
    
    print(f"  ✓ Wrote {len(error_log_data) - 1} issues to Error Log")

print(f"\n✅ COMPLETE!")
print(f"\nCheck the 'Error Log' sheet in your MASTER SPRING 2026 spreadsheet")

if total_issues == 0:
    print("\n🎉 No errors found! All student data looks good.")
else:
    print(f"\n⚠️  Found {total_issues} issues that need review")