import gspread
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.http import MediaIoBaseDownload, MediaFileUpload
import os
import pickle
import io
from PyPDF2 import PdfMerger
import time

# Set up OAuth credentials
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/documents',
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

# Connect to services
gc = gspread.authorize(creds)
docs_service = build('docs', 'v1', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

# Open spreadsheet
spreadsheet = gc.open('MASTER SPRING 2026')

# Get all sheets and find school sheets
all_sheets = spreadsheet.worksheets()
school_sheets = [sheet.title for sheet in all_sheets if sheet.title.endswith(' MASTER') and sheet.title != 'MASTER']

if not school_sheets:
    print("No school sheets found!")
    print("Please run organize_schools.py first to create school-specific sheets.")
    exit()

# Display schools and let user choose
print("\nAvailable schools:")
for idx, school_name in enumerate(school_sheets, 1):
    # Remove " MASTER" from display
    display_name = school_name.replace(' MASTER', '')
    print(f"  {idx}. {display_name}")

choice = input("\nEnter the number of the school you want to process: ")

try:
    school_index = int(choice) - 1
    selected_sheet_name = school_sheets[school_index]
    school_name = selected_sheet_name.replace(' MASTER', '')
except:
    print("Invalid choice!")
    exit()

print(f"\n✓ Selected: {school_name}")

# Find or create template
print("\nFinding template...")
template_query = "name='Order Template for PDF' and mimeType='application/vnd.google-apps.document'"
template_results = drive_service.files().list(q=template_query).execute()
templates = template_results.get('files', [])

if not templates:
    print("ERROR: Template 'Order Template for PDF' not found!")
    print("Please create or rename your template to 'Order Template for PDF'")
    exit()

TEMPLATE_ID = templates[0]['id']
print(f"✓ Found template")

# Find or create folder structure
print("\nSetting up folders...")

# Main folder: "[School Name] Orders"
main_folder_name = f"{school_name} Orders"
main_folder_query = f"name='{main_folder_name}' and mimeType='application/vnd.google-apps.folder'"
main_folder_results = drive_service.files().list(q=main_folder_query).execute()
main_folders = main_folder_results.get('files', [])

if main_folders:
    main_folder_id = main_folders[0]['id']
    print(f"✓ Found '{main_folder_name}' folder")
else:
    # Create main folder
    main_folder_metadata = {
        'name': main_folder_name,
        'mimeType': 'application/vnd.google-apps.folder'
    }
    main_folder = drive_service.files().create(body=main_folder_metadata, fields='id').execute()
    main_folder_id = main_folder['id']
    print(f"✓ Created '{main_folder_name}' folder")

# Individual Documents subfolder
docs_folder_name = f"{school_name} Individual Documents"
docs_folder_query = f"name='{docs_folder_name}' and '{main_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
docs_folder_results = drive_service.files().list(q=docs_folder_query).execute()
docs_folders = docs_folder_results.get('files', [])

if docs_folders:
    docs_folder_id = docs_folders[0]['id']
    print(f"✓ Found '{docs_folder_name}' folder")
else:
    docs_folder_metadata = {
        'name': docs_folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [main_folder_id]
    }
    docs_folder = drive_service.files().create(body=docs_folder_metadata, fields='id').execute()
    docs_folder_id = docs_folder['id']
    print(f"✓ Created '{docs_folder_name}' folder")

# PDFs subfolder
pdfs_folder_name = f"{school_name} PDFs"
pdfs_folder_query = f"name='{pdfs_folder_name}' and '{main_folder_id}' in parents and mimeType='application/vnd.google-apps.folder'"
pdfs_folder_results = drive_service.files().list(q=pdfs_folder_query).execute()
pdfs_folders = pdfs_folder_results.get('files', [])

if pdfs_folders:
    pdfs_folder_id = pdfs_folders[0]['id']
    print(f"✓ Found '{pdfs_folder_name}' folder")
else:
    pdfs_folder_metadata = {
        'name': pdfs_folder_name,
        'mimeType': 'application/vnd.google-apps.folder',
        'parents': [main_folder_id]
    }
    pdfs_folder = drive_service.files().create(body=pdfs_folder_metadata, fields='id').execute()
    pdfs_folder_id = pdfs_folder['id']
    print(f"✓ Created '{pdfs_folder_name}' folder")

# Read from school-specific sheet
school_sheet = spreadsheet.worksheet(selected_sheet_name)
data = school_sheet.get_all_values()
headers = data[0]
rows = data[1:]

print(f"\nSuccess! Found {len(rows)} rows in {selected_sheet_name}")

# Filter for "Pick-up at school" orders
# Column indices in school sheet: A, AW, AY, Q, R, S, O, Y, AV
col_order_num = 0   # A
col_student = 1     # AW
col_grade = 2       # AY
col_quantity = 3    # Q
col_flavor = 4      # R
col_price = 5       # S
col_delivery = 6    # O
col_billing = 7     # Y
col_school = 8      # AV

pickup_orders = []
for row in rows:
    if len(row) > col_delivery and row[col_delivery] == 'Pick-up at school':
        pickup_orders.append(row)

print(f"Found {len(pickup_orders)} orders with 'Pick-up at school'")

# Group orders
orders = {}

for row in pickup_orders:
    order_num = row[col_order_num]
    quantity = int(row[col_quantity]) if row[col_quantity].isdigit() else 0
    flavor = row[col_flavor]
    
    student_name = row[col_student] if len(row) > col_student else ''
    student_grade = row[col_grade] if len(row) > col_grade else ''
    billing_name = row[col_billing] if len(row) > col_billing else ''
    school = row[col_school] if len(row) > col_school else ''
    
    if order_num not in orders:
        orders[order_num] = {
            'order_number': order_num,
            'billing_name': billing_name,
            'school': school,
            'student_name': student_name,
            'student_grade': student_grade,
            'items': []
        }
    
    orders[order_num]['items'].append({'flavor': flavor, 'quantity': quantity})

print(f"Grouped into {len(orders)} unique orders")

# Sort orders by grade then student name
def grade_sort_key(grade):
    grade_upper = grade.upper().strip()
    
    if grade_upper == 'K' or grade_upper.startswith('KINDER'):
        return (0, '')
    
    if grade_upper.isdigit():
        return (int(grade_upper), '')
    
    if grade_upper and grade_upper[0].isdigit():
        return (int(grade_upper[0]), grade_upper)
    
    return (999, grade_upper)

sorted_orders = sorted(
    orders.items(),
    key=lambda x: (grade_sort_key(x[1]['student_grade']), x[1]['student_name'])
)

print("\nOrders sorted by grade then student name:")
for order_num, order in sorted_orders:
    print(f"  Grade {order['student_grade']}: {order['student_name']} (Order #{order_num})")

# Delete old files
print("\nCleaning up old files...")

old_docs_query = f"'{docs_folder_id}' in parents"
old_docs = drive_service.files().list(q=old_docs_query).execute().get('files', [])
for doc in old_docs:
    drive_service.files().delete(fileId=doc['id']).execute()
print(f"Deleted {len(old_docs)} old documents")

old_pdfs_query = f"'{pdfs_folder_id}' in parents"
old_pdfs = drive_service.files().list(q=old_pdfs_query).execute().get('files', [])
for pdf in old_pdfs:
    drive_service.files().delete(fileId=pdf['id']).execute()
print(f"Deleted {len(old_pdfs)} old PDFs")

# Create individual documents
print("\nCreating individual order documents...")
pdf_files = []

for order_idx, (order_num, order) in enumerate(sorted_orders):
    print(f"  Creating document for order #{order_num} ({order_idx + 1}/{len(sorted_orders)})...")
    
    # Copy template
    copy_title = f"Grade {order['student_grade']} - {order['student_name']} - Order {order_num}"
    order_copy = drive_service.files().copy(
        fileId=TEMPLATE_ID,
        body={'name': copy_title}
    ).execute()
    order_copy_id = order_copy.get('id')
    
    # Build replacements
    all_requests = []
    
    main_replacements = [
        ('{{Order Number}}', order['order_number']),
        ('{{Billing Name}}', order['billing_name']),
        ('{{Student name}}', order['student_name']),
        ('{{student name}}', order['student_name']),
        ('{{Grade}}', order['student_grade']),
        ('{{School}}', order['school'])
    ]
    
    for placeholder, value in main_replacements:
        all_requests.append({
            'replaceAllText': {
                'containsText': {'text': placeholder, 'matchCase': True},
                'replaceText': value
            }
        })
    
    # Item replacements
    for i in range(1, 14):
        item_index = i - 1
        
        if item_index < len(order['items']):
            item = order['items'][item_index]
            qty_value = str(item['quantity'])
            flavor_value = item['flavor']
        else:
            qty_value = ''
            flavor_value = ''
        
        qty_placeholder = '{{quantity' + str(i) + '}}'
        all_requests.append({
            'replaceAllText': {
                'containsText': {'text': qty_placeholder, 'matchCase': True},
                'replaceText': qty_value
            }
        })
        
        flavor_placeholder = '{{flavor name' + str(i) + '}}'
        all_requests.append({
            'replaceAllText': {
                'containsText': {'text': flavor_placeholder, 'matchCase': True},
                'replaceText': flavor_value
            }
        })
    
    # Apply replacements
    docs_service.documents().batchUpdate(
        documentId=order_copy_id,
        body={'requests': all_requests}
    ).execute()
    
    # Calculate popcorn vs coffee counts
    popcorn_count = 0
    coffee_count = 0
    
    for item in order['items']:
        flavor = item['flavor'].lower()
        quantity = item['quantity']
        
        if 'coffee' in flavor:
            coffee_count += quantity
        else:
            popcorn_count += quantity
    
    # Get document to find items table end
    doc = docs_service.documents().get(documentId=order_copy_id).execute()
    content = doc.get('body').get('content')
    
    # Find the items table
    items_table_end = None
    for element in content:
        if 'table' in element:
            table = element.get('table')
            table_text = ""
            
            for row in table.get('tableRows', []):
                for cell in row.get('tableCells', []):
                    for cell_content in cell.get('content', []):
                        if 'paragraph' in cell_content:
                            for elem in cell_content.get('paragraph', {}).get('elements', []):
                                if 'textRun' in elem:
                                    table_text += elem.get('textRun', {}).get('content', '')
            
            if 'Quantity' in table_text and 'Flavor' in table_text:
                items_table_end = element.get('endIndex')
                break
    
    # Insert summary after the table
    if items_table_end:
        popcorn_label = "bag" if popcorn_count == 1 else "bags"
        coffee_label = "bag" if coffee_count == 1 else "bags"
        summary_text = f"\n\nPopcorn: {popcorn_count} {popcorn_label}     Coffee: {coffee_count} {coffee_label}\n"
        
        summary_requests = [
            {
                'insertText': {
                    'location': {'index': items_table_end},
                    'text': summary_text
                }
            },
            {
                'updateTextStyle': {
                    'range': {
                        'startIndex': items_table_end,
                        'endIndex': items_table_end + len(summary_text)
                    },
                    'textStyle': {
                        'bold': True,
                        'fontSize': {'magnitude': 12, 'unit': 'PT'},
                        'weightedFontFamily': {
                            'fontFamily': 'Lexend',
                            'weight': 700
                        }
                    },
                    'fields': 'bold,fontSize,weightedFontFamily'
                }
            }
        ]
        
        docs_service.documents().batchUpdate(
            documentId=order_copy_id,
            body={'requests': summary_requests}
        ).execute()
    
    # Delete empty rows from items table
    num_items = len(order['items'])
    if num_items < 13:
        doc = docs_service.documents().get(documentId=order_copy_id).execute()
        content = doc.get('body').get('content')
        
        items_table = None
        items_table_start = None
        
        for element in content:
            if 'table' in element:
                table = element.get('table')
                table_text = ""
                
                for row in table.get('tableRows', []):
                    for cell in row.get('tableCells', []):
                        for cell_content in cell.get('content', []):
                            if 'paragraph' in cell_content:
                                for elem in cell_content.get('paragraph', {}).get('elements', []):
                                    if 'textRun' in elem:
                                        table_text += elem.get('textRun', {}).get('content', '')
                
                if 'Quantity' in table_text and 'Flavor' in table_text:
                    items_table = table
                    items_table_start = element.get('startIndex')
                    break
        
        if items_table and items_table_start:
            num_rows = len(items_table.get('tableRows', []))
            
            delete_requests = []
            rows_to_delete = num_rows - (num_items + 1)
            
            if rows_to_delete > 0:
                for i in range(rows_to_delete):
                    delete_requests.append({
                        'deleteTableRow': {
                            'tableCellLocation': {
                                'tableStartLocation': {'index': items_table_start},
                                'rowIndex': num_items + 1,
                                'columnIndex': 0
                            }
                        }
                    })
                
                if delete_requests:
                    docs_service.documents().batchUpdate(
                        documentId=order_copy_id,
                        body={'requests': delete_requests}
                    ).execute()
    
    # Move to Individual Documents folder
    drive_service.files().update(
        fileId=order_copy_id,
        addParents=docs_folder_id,
        fields='id, parents'
    ).execute()
    
    # Export as PDF
    request = drive_service.files().export_media(
        fileId=order_copy_id,
        mimeType='application/pdf'
    )
    
    pdf_filename = f"temp_{order_idx}.pdf"
    fh = io.FileIO(pdf_filename, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
    fh.close()
    
    pdf_files.append(pdf_filename)

print(f"\n✓ Created {len(sorted_orders)} individual documents")

# Combine PDFs
print("\nCombining PDFs in sorted order...")
merger = PdfMerger()

for pdf_file in pdf_files:
    merger.append(pdf_file)

combined_pdf_filename = f"{school_name}_Orders_Combined.pdf"
merger.write(combined_pdf_filename)
merger.close()

print(f"✓ Combined into one PDF")

# Upload combined PDF
print("\nUploading combined PDF to Google Drive...")
file_metadata = {
    'name': f'{school_name} Orders - Combined.pdf',
    'parents': [pdfs_folder_id]
}

media = MediaFileUpload(combined_pdf_filename, mimetype='application/pdf')
combined_pdf = drive_service.files().create(
    body=file_metadata,
    media_body=media,
    fields='id, webViewLink'
).execute()

print(f"✓ Uploaded combined PDF")

# Clean up local files
print("\nCleaning up temporary files...")
time.sleep(2)

for pdf_file in pdf_files:
    try:
        os.remove(pdf_file)
    except:
        pass

try:
    os.remove(combined_pdf_filename)
except:
    pass

print(f"\n✅ COMPLETE!")
print(f"\n📁 Individual documents: {len(sorted_orders)} files in '{school_name} Individual Documents' folder")
print(f"📄 Combined PDF: '{school_name} Orders - Combined.pdf' in '{school_name} PDFs' folder")
print(f"\nView combined PDF: {combined_pdf['webViewLink']}")