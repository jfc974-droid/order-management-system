import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle

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
all_data = master_sheet.get_all_values()
headers = all_data[0]
rows = all_data[1:]

print(f"Found {len(rows)} rows")

# Column indices (0-based)
col_A = 0   # Order Number
col_O = 14  # Shipping Method
col_Q = 16  # Quantity
col_R = 17  # Flavor
col_S = 18  # Price per item
col_Y = 24  # Billing Name
col_AV = 47 # School
col_AW = 48 # Student Name
col_AY = 50 # Grade

# Define readable pastel colors for highlighting
SCHOOL_COLORS = [
    {'red': 1.0, 'green': 0.9, 'blue': 0.9},     # Light pink
    {'red': 0.9, 'green': 1.0, 'blue': 0.9},     # Light green
    {'red': 0.9, 'green': 0.9, 'blue': 1.0},     # Light blue
    {'red': 1.0, 'green': 1.0, 'blue': 0.9},     # Light yellow
    {'red': 1.0, 'green': 0.9, 'blue': 1.0},     # Light purple
    {'red': 0.9, 'green': 1.0, 'blue': 1.0},     # Light cyan
    {'red': 1.0, 'green': 0.95, 'blue': 0.9},    # Light peach
    {'red': 0.95, 'green': 0.95, 'blue': 1.0},   # Light lavender
    {'red': 0.9, 'green': 1.0, 'blue': 0.95},    # Light mint
    {'red': 1.0, 'green': 0.9, 'blue': 0.95},    # Light coral
]

# Group orders by school
schools = {}
school_color_map = {}
color_index = 0

for idx, row in enumerate(rows):
    if len(row) > col_AV:
        school_name = row[col_AV].strip()
        
        if school_name:
            # Assign color if new school
            if school_name not in school_color_map:
                school_color_map[school_name] = SCHOOL_COLORS[color_index % len(SCHOOL_COLORS)]
                color_index += 1
            
            if school_name not in schools:
                schools[school_name] = []
            
            # Extract and rearrange columns: A, AW, AY, Q, R, S, O, Y, AV
            new_row = [
                row[col_A] if len(row) > col_A else '',      # A - Order Number
                row[col_AW] if len(row) > col_AW else '',    # AW - Student Name
                row[col_AY] if len(row) > col_AY else '',    # AY - Grade
                row[col_Q] if len(row) > col_Q else '',      # Q - Quantity
                row[col_R] if len(row) > col_R else '',      # R - Flavor
                row[col_S] if len(row) > col_S else '',      # S - Price per item
                row[col_O] if len(row) > col_O else '',      # O - Shipping Method
                row[col_Y] if len(row) > col_Y else '',      # Y - Billing Name
                row[col_AV] if len(row) > col_AV else '',    # AV - School
            ]
            
            schools[school_name].append({
                'row_index': idx + 2,  # +2 because: +1 for header, +1 for 1-based index
                'data': new_row
            })

print(f"\nFound {len(schools)} schools:")
for school, orders in schools.items():
    print(f"  {school}: {len(orders)} orders")

# Step 1: Highlight rows in MASTER sheet by school color
print("\nHighlighting rows in MASTER sheet...")

batch_updates = []

for school_name, school_data in schools.items():
    color = school_color_map[school_name]
    
    for order in school_data:
        row_idx = order['row_index']
        
        # Format entire row with school color
        batch_updates.append({
            'repeatCell': {
                'range': {
                    'sheetId': master_sheet.id,
                    'startRowIndex': row_idx - 1,
                    'endRowIndex': row_idx,
                },
                'cell': {
                    'userEnteredFormat': {
                        'backgroundColor': color
                    }
                },
                'fields': 'userEnteredFormat.backgroundColor'
            }
        })

# Apply all highlighting at once
if batch_updates:
    spreadsheet.batch_update({'requests': batch_updates})
    print(f"✓ Highlighted {len(batch_updates)} rows")

# Step 2: Create/update school sheets
print("\nCreating/updating school sheets...")

# Get new header order: A, AW, AY, Q, R, S, O, Y, AV
new_headers = [
    headers[col_A],
    headers[col_AW],
    headers[col_AY],
    headers[col_Q],
    headers[col_R],
    headers[col_S],
    headers[col_O],
    headers[col_Y],
    headers[col_AV]
]

for school_name, school_orders in schools.items():
    sheet_name = f"{school_name} MASTER"
    
    print(f"  Processing {sheet_name}...")
    
    # Check if sheet exists
    try:
        school_sheet = spreadsheet.worksheet(sheet_name)
        existing_sheet = True
        print(f"    ✓ Found existing sheet")
    except:
        # Create new sheet
        school_sheet = spreadsheet.add_worksheet(title=sheet_name, rows=1000, cols=20)
        existing_sheet = False
        print(f"    ✓ Created new sheet")
    
    if not existing_sheet:
        # New sheet - add headers
        school_sheet.update('A1:I1', [new_headers])
        
        # Format header
        school_sheet.format('A1:I1', {
            'backgroundColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2},
            'textFormat': {'foregroundColor': {'red': 1, 'green': 1, 'blue': 1}, 'bold': True}
        })
        
        # Add all data
        data_to_add = [order['data'] for order in school_orders]
        if data_to_add:
            school_sheet.append_rows(data_to_add)
        
        print(f"    ✓ Added {len(data_to_add)} orders")
    
    else:
        # Existing sheet - check which orders are new
        existing_data = school_sheet.get_all_values()
        existing_order_nums = set()
        
        if len(existing_data) > 1:
            # Get existing order numbers (column A)
            for row in existing_data[1:]:
                if row and row[0]:
                    existing_order_nums.add(row[0])
        
        # Find new orders
        new_orders = []
        for order in school_orders:
            order_num = order['data'][0]
            if order_num not in existing_order_nums:
                new_orders.append(order['data'])
        
        if new_orders:
            school_sheet.append_rows(new_orders)
            print(f"    ✓ Added {len(new_orders)} new orders")
        else:
            print(f"    ✓ No new orders to add")

print(f"\n✅ COMPLETE!")
print(f"Processed {len(schools)} schools")
print(f"MASTER sheet rows are now color-coded by school")