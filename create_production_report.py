import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import os
import pickle
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
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
col_quantity = 16     # Column Q
col_flavor = 17       # Column R
col_delivery = 14     # Column O
col_school = 47       # Column AV

# Data structure: {school: {flavor: {pickup: count, shipping: count}}}
schools_data = {}
all_flavors_data = {}  # For combined totals

for row in rows:
    if len(row) > col_school:
        school = row[col_school].strip()
        flavor = row[col_flavor].strip()
        delivery = row[col_delivery].strip()
        
        try:
            quantity = int(row[col_quantity]) if row[col_quantity].isdigit() else 0
        except:
            quantity = 0
        
        if not school or not flavor or quantity == 0:
            continue
        
        # Normalize delivery method
        if 'pick' in delivery.lower():
            delivery_type = 'pickup'
        else:
            delivery_type = 'shipping'
        
        # Track by school
        if school not in schools_data:
            schools_data[school] = {}
        
        if flavor not in schools_data[school]:
            schools_data[school][flavor] = {'pickup': 0, 'shipping': 0}
        
        schools_data[school][flavor][delivery_type] += quantity
        
        # Track combined totals
        if flavor not in all_flavors_data:
            all_flavors_data[flavor] = {'pickup': 0, 'shipping': 0}
        
        all_flavors_data[flavor][delivery_type] += quantity

print(f"\nFound {len(schools_data)} schools")
print(f"Found {len(all_flavors_data)} unique flavors")

# Calculate grand totals
grand_pickup_total = sum(f['pickup'] for f in all_flavors_data.values())
grand_shipping_total = sum(f['shipping'] for f in all_flavors_data.values())

# Create PDF
print("\nCreating PDF report...")

pdf_filename = f"Production_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf"
doc = SimpleDocTemplate(pdf_filename, pagesize=letter)
story = []

# Styles
styles = getSampleStyleSheet()
title_style = ParagraphStyle(
    'CustomTitle',
    parent=styles['Heading1'],
    fontSize=24,
    textColor=colors.HexColor('#2d3748'),
    spaceAfter=30,
    alignment=1  # Center
)

school_header_style = ParagraphStyle(
    'SchoolHeader',
    parent=styles['Heading2'],
    fontSize=16,
    textColor=colors.HexColor('#2d3748'),
    spaceAfter=12,
    spaceBefore=20
)

# Title
story.append(Paragraph("Production Report", title_style))
story.append(Paragraph(f"Generated: {datetime.now().strftime('%B %d, %Y at %I:%M %p')}", styles['Normal']))
story.append(Spacer(1, 0.3 * inch))

# Add each school's table
for school_name in sorted(schools_data.keys()):
    school_flavors = schools_data[school_name]
    
    # School header
    story.append(Paragraph(school_name, school_header_style))
    
    # Build table data
    table_data = [['Flavor', 'Pick-up', 'Shipping']]
    
    school_pickup_total = 0
    school_shipping_total = 0
    
    for flavor in sorted(school_flavors.keys()):
        pickup = school_flavors[flavor]['pickup']
        shipping = school_flavors[flavor]['shipping']
        
        table_data.append([flavor, str(pickup), str(shipping)])
        
        school_pickup_total += pickup
        school_shipping_total += shipping
    
    # School totals
    table_data.append(['TOTAL', str(school_pickup_total), str(school_shipping_total)])
    
    # Create table
    table = Table(table_data, colWidths=[3*inch, 1.5*inch, 1.5*inch])
    table.setStyle(TableStyle([
        ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
        ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
        ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
        ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
        ('FONTSIZE', (0, 0), (-1, 0), 12),
        ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
        ('BACKGROUND', (0, -1), (-1, -1), colors.beige),
        ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
        ('GRID', (0, 0), (-1, -1), 1, colors.black),
        ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
        ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.lightgrey]),
    ]))
    
    story.append(table)
    story.append(Spacer(1, 0.3 * inch))

# Add combined totals
story.append(Paragraph("ALL SCHOOLS - TOTAL PRODUCTION NEEDED", school_header_style))

combined_table_data = [['Flavor', 'Pick-up', 'Shipping', 'TOTAL']]

for flavor in sorted(all_flavors_data.keys()):
    pickup = all_flavors_data[flavor]['pickup']
    shipping = all_flavors_data[flavor]['shipping']
    total = pickup + shipping
    
    combined_table_data.append([flavor, str(pickup), str(shipping), str(total)])

# Grand totals
combined_table_data.append(['GRAND TOTAL', str(grand_pickup_total), str(grand_shipping_total), str(grand_pickup_total + grand_shipping_total)])

# Create combined table
combined_table = Table(combined_table_data, colWidths=[2.5*inch, 1.3*inch, 1.3*inch, 1.3*inch])
combined_table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.HexColor('#2d3748')),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('ALIGN', (0, 0), (-1, -1), 'CENTER'),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, 0), 12),
    ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
    ('BACKGROUND', (0, -1), (-1, -1), colors.HexColor('#4CAF50')),
    ('TEXTCOLOR', (0, -1), (-1, -1), colors.whitesmoke),
    ('FONTNAME', (0, -1), (-1, -1), 'Helvetica-Bold'),
    ('FONTSIZE', (0, -1), (-1, -1), 12),
    ('GRID', (0, 0), (-1, -1), 1, colors.black),
    ('VALIGN', (0, 0), (-1, -1), 'MIDDLE'),
    ('ROWBACKGROUNDS', (0, 1), (-1, -2), [colors.white, colors.lightgrey]),
]))

story.append(combined_table)

# Build PDF
doc.build(story)
print(f"PDF created: {pdf_filename}")

# Create or clear Production sheet in Google Sheets
print("\nUpdating Google Sheet...")

try:
    production_sheet = spreadsheet.worksheet('Production')
    print("  Found existing 'Production' sheet - clearing it...")
    production_sheet.clear()
except:
    print("  Creating new 'Production' sheet...")
    production_sheet = spreadsheet.add_worksheet(title='Production', rows=1000, cols=10)

# Build the sheet data
sheet_data = []

# Add each school's table
for school_name in sorted(schools_data.keys()):
    school_flavors = schools_data[school_name]
    
    # School header
    sheet_data.append([school_name])
    sheet_data.append(['Flavor', 'Pick-up', 'Shipping'])
    
    school_pickup_total = 0
    school_shipping_total = 0
    
    # Flavor rows
    for flavor in sorted(school_flavors.keys()):
        pickup = school_flavors[flavor]['pickup']
        shipping = school_flavors[flavor]['shipping']
        
        sheet_data.append([flavor, pickup, shipping])
        
        school_pickup_total += pickup
        school_shipping_total += shipping
    
    # School totals
    sheet_data.append(['TOTAL', school_pickup_total, school_shipping_total])
    
    # Blank row between schools
    sheet_data.append([])

# Add combined totals table
sheet_data.append(['ALL SCHOOLS - TOTAL PRODUCTION NEEDED'])
sheet_data.append(['Flavor', 'Pick-up', 'Shipping', 'TOTAL'])

for flavor in sorted(all_flavors_data.keys()):
    pickup = all_flavors_data[flavor]['pickup']
    shipping = all_flavors_data[flavor]['shipping']
    total = pickup + shipping
    
    sheet_data.append([flavor, pickup, shipping, total])

# Grand totals
sheet_data.append(['GRAND TOTAL', grand_pickup_total, grand_shipping_total, grand_pickup_total + grand_shipping_total])

# Write all data to sheet
production_sheet.update(values=sheet_data, range_name='A1')

# Format the sheet
print("\nFormatting Production sheet...")

row_index = 1

for school_name in sorted(schools_data.keys()):
    # Format school header
    production_sheet.format(f'A{row_index}', {
        'backgroundColor': {'red': 0.3, 'green': 0.5, 'blue': 0.8},
        'textFormat': {
            'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
            'bold': True,
            'fontSize': 14
        }
    })
    
    row_index += 1
    
    # Format column headers
    production_sheet.format(f'A{row_index}:C{row_index}', {
        'backgroundColor': {'red': 0.9, 'green': 0.9, 'blue': 0.9},
        'textFormat': {'bold': True},
        'horizontalAlignment': 'CENTER'
    })
    
    row_index += 1
    
    # Flavor rows
    num_flavors = len(schools_data[school_name])
    row_index += num_flavors
    
    # Format totals row
    production_sheet.format(f'A{row_index}:C{row_index}', {
        'backgroundColor': {'red': 1, 'green': 1, 'blue': 0.8},
        'textFormat': {'bold': True}
    })
    
    row_index += 2

# Format combined totals section
# Header
production_sheet.format(f'A{row_index}', {
    'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.2},
    'textFormat': {
        'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
        'bold': True,
        'fontSize': 14
    }
})

row_index += 1

# Column headers
production_sheet.format(f'A{row_index}:D{row_index}', {
    'backgroundColor': {'red': 0.2, 'green': 0.2, 'blue': 0.2},
    'textFormat': {
        'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
        'bold': True
    },
    'horizontalAlignment': 'CENTER'
})

row_index += len(all_flavors_data) + 1

# Grand total row
production_sheet.format(f'A{row_index}:D{row_index}', {
    'backgroundColor': {'red': 0.2, 'green': 0.6, 'blue': 0.2},
    'textFormat': {
        'foregroundColor': {'red': 1, 'green': 1, 'blue': 1},
        'bold': True,
        'fontSize': 12
    }
})

# Auto-resize columns
production_sheet.columns_auto_resize(0, 3)

print(f"\nCOMPLETE!")
print(f"\nProduction report created:")
print(f"  - PDF file: {pdf_filename}")
print(f"  - Google Sheet: Production")
print(f"\nSummary:")
print(f"  - {len(schools_data)} schools")
print(f"  - Grand total: {grand_pickup_total + grand_shipping_total} bags")
print(f"    - Pick-up: {grand_pickup_total}")
print(f"    - Shipping: {grand_shipping_total}")