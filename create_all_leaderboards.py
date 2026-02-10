import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
from datetime import datetime

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

def create_leaderboard_html(school_name, top_students, timestamp):
    """Generate HTML for a school's leaderboard"""
    
    medals = ['🥇', '🥈', '🥉', '🌟', '⭐']
    
    html = f"""
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>{school_name} Top Sellers</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Arial', sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            padding: 20px;
            min-height: 100vh;
            display: flex;
            align-items: center;
            justify-content: center;
        }}
        
        .leaderboard {{
            background: white;
            border-radius: 20px;
            padding: 40px;
            max-width: 600px;
            width: 100%;
            box-shadow: 0 20px 60px rgba(0,0,0,0.3);
        }}
        
        .header {{
            text-align: center;
            margin-bottom: 30px;
        }}
        
        .header h1 {{
            color: #2d3748;
            font-size: 2.5em;
            margin-bottom: 10px;
        }}
        
        .header .trophy {{
            font-size: 3em;
            margin-bottom: 10px;
        }}
        
        .header .subtitle {{
            color: #718096;
            font-size: 1.1em;
        }}
        
        .student {{
            display: flex;
            align-items: center;
            padding: 20px;
            margin-bottom: 15px;
            border-radius: 15px;
            transition: transform 0.3s ease, box-shadow 0.3s ease;
            position: relative;
        }}
        
        .student:hover {{
            transform: translateY(-5px);
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }}
        
        .rank-1 {{
            background: linear-gradient(135deg, #ffd700 0%, #ffed4e 100%);
            border: 3px solid #d4af37;
        }}
        
        .rank-2 {{
            background: linear-gradient(135deg, #c0c0c0 0%, #e8e8e8 100%);
            border: 3px solid #a8a8a8;
        }}
        
        .rank-3 {{
            background: linear-gradient(135deg, #cd7f32 0%, #e9a76b 100%);
            border: 3px solid #b87333;
        }}
        
        .rank-4, .rank-5 {{
            background: linear-gradient(135deg, #e0e7ff 0%, #f0f4ff 100%);
            border: 3px solid #c7d2fe;
        }}
        
        .rank {{
            font-size: 2em;
            font-weight: bold;
            width: 60px;
            text-align: center;
            color: #2d3748;
        }}
        
        .info {{
            flex: 1;
            padding: 0 20px;
        }}
        
        .name {{
            font-size: 1.4em;
            font-weight: bold;
            color: #2d3748;
            margin-bottom: 5px;
        }}
        
        .grade {{
            color: #718096;
            font-size: 1em;
        }}
        
        .sales {{
            font-size: 1.6em;
            font-weight: bold;
            color: #2d3748;
            white-space: nowrap;
            margin-right: 10px;
        }}
        
        .medal {{
            font-size: 2.5em;
        }}
        
        .updated {{
            text-align: center;
            color: #718096;
            font-size: 0.9em;
            margin-top: 30px;
        }}
        
        @media (max-width: 600px) {{
            .leaderboard {{
                padding: 20px;
            }}
            
            .header h1 {{
                font-size: 1.8em;
            }}
            
            .name {{
                font-size: 1.1em;
            }}
            
            .sales {{
                font-size: 1.3em;
            }}
            
            .rank {{
                font-size: 1.5em;
                width: 40px;
            }}
        }}
    </style>
</head>
<body>
    <div class="leaderboard">
        <div class="header">
            <div class="trophy">🏆</div>
            <h1>Top Sellers</h1>
            <div class="subtitle">{school_name}</div>
        </div>
        
"""
    
    # Add top students
    for idx, (name, data) in enumerate(top_students, 1):
        html += f"""
        <div class="student rank-{idx}">
            <div class="rank">#{idx}</div>
            <div class="info">
                <div class="name">{name}</div>
                <div class="grade">Grade {data['grade']}</div>
            </div>
            <div class="sales">${data['total']:,.2f}</div>
            <div class="medal">{medals[idx-1]}</div>
        </div>
"""
    
    # Add footer
    html += f"""
        
        <div class="updated">
            Last updated: {timestamp}
        </div>
    </div>
</body>
</html>
"""
    
    return html

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
col_order = 0
col_quantity = 16
col_flavor = 17
col_line_total = 18
col_school = 47
col_student = 48
col_grade = 50

# Group sales by school and student
schools_data = {}

for row in rows:
    if len(row) > col_school:
        school = row[col_school].strip()
        student = row[col_student].strip()
        grade = row[col_grade].strip()
        
        if not school or not student:
            continue
        
        # Calculate line total (Quantity × Price)
        try:
            quantity = int(row[col_quantity]) if row[col_quantity].isdigit() else 0
            price = float(row[col_line_total].replace('$', '').replace(',', '').strip())
            amount = quantity * price
        except:
            amount = 0.0
        
        # Initialize school if needed
        if school not in schools_data:
            schools_data[school] = {}
        
        # Initialize student if needed
        if student not in schools_data[school]:
            schools_data[school][student] = {
                'grade': grade,
                'total': 0.0
            }
        
        # Add to student's total
        schools_data[school][student]['total'] += amount

print(f"\nFound {len(schools_data)} schools")

# Generate leaderboards for each school
timestamp = datetime.now().strftime("%B %d, %Y at %I:%M %p")
leaderboards_created = []

for school_name, students in schools_data.items():
    print(f"\nProcessing {school_name}...")
    
    # Sort students by total sales
    top_students = sorted(
        students.items(),
        key=lambda x: x[1]['total'],
        reverse=True
    )[:5]  # Top 5
    
    if not top_students:
        print(f"  No students found for {school_name}")
        continue
    
    print(f"  Top 5 students:")
    for idx, (name, data) in enumerate(top_students, 1):
        print(f"    {idx}. {name} (Grade {data['grade']}): ${data['total']:.2f}")
    
    # Generate HTML
    html_content = create_leaderboard_html(school_name, top_students, timestamp)
    
    # Save to file
    # Clean school name for filename
    safe_school_name = school_name.replace(' ', '_').replace('/', '_')
    filename = f"leaderboard_{safe_school_name}.html"
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    leaderboards_created.append({
        'school': school_name,
        'file': filename,
        'count': len(top_students)
    })
    
    print(f"  ✓ Created {filename}")

print(f"\n✅ COMPLETE! Created {len(leaderboards_created)} leaderboards:")
for lb in leaderboards_created:
    print(f"  • {lb['school']}: {lb['file']} ({lb['count']} students)")

print(f"\nAll leaderboard files are in: C:\\Users\\jfc97\\Documents\\OrderAutomation\\")
print(f"\nYou can open each HTML file in your browser to preview!")