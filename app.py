from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import pdfplumber
from io import BytesIO

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
RESULT_FOLDER = 'results'

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(RESULT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['RESULT_FOLDER'] = RESULT_FOLDER

def extract_pdf_to_excel(pdf_path, excel_path):
    """
    Extracts student result data from PDF with fixed column handling
    """
    # Initialize an empty list to store all records
    all_records = []
    
    with pdfplumber.open(pdf_path) as pdf:
        current_roll = None
        current_name = None
        
        for page in pdf.pages:
            lines = page.extract_text().split('\n')
            
            for i, line in enumerate(lines):
                # Skip empty lines and headers
                if not line.strip() or 'Course Code:' in line or 'Branch Name' in line:
                    continue
                
                # If line starts with a roll number (registration number)
                if line.strip().startswith('2'):
                    parts = line.split()
                    current_roll = parts[0]
                    
                    # Extract name (everything between roll number and the next number)
                    name_parts = []
                    for part in parts[1:]:
                        if part.isdigit():
                            break
                        name_parts.append(part)
                    current_name = ' '.join(name_parts)
                    
                    # Look for the subject codes line
                    subject_codes = []
                    for j in range(i, min(i + 3, len(lines))):
                        parts = lines[j].split()
                        for part in parts:
                            if part.isdigit() and len(part) == 4:
                                subject_codes.append(part)
                    
                    # Look for the results line
                    for j in range(i, min(i + 3, len(lines))):
                        if any(x in lines[j] for x in ['P ', 'R ', 'A ']):
                            results = [x for x in lines[j].split() if x in ['P', 'R', 'A']]
                            
                            # Map results to full words
                            result_map = {'P': 'Pass', 'R': 'Fail', 'A': 'Absent'}
                            
                            # Create records for each subject code and result
                            for code, result in zip(subject_codes, results):
                                record = {
                                    'Roll_Number': current_roll,
                                    'Name': current_name,
                                    'Subject_Code': code,
                                    'Result': result_map.get(result, 'Unknown')
                                }
                                all_records.append(record)
                            break
    
    # Create DataFrame
    if all_records:
        df = pd.DataFrame(all_records)
        # Save to Excel
        df.to_excel(excel_path, index=False)
        return df
    else:
        # Create empty DataFrame with correct columns
        df = pd.DataFrame(columns=['Roll_Number', 'Name', 'Subject_Code', 'Result'])
        df.to_excel(excel_path, index=False)
        return df

@app.route('/')
def home():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_pdf():
    if 'file' not in request.files:
        return "No file part", 400
    file = request.files['file']
    if file.filename == '':
        return "No selected file", 400
    
    filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
    file.save(filepath)
    
    excel_path = os.path.join(app.config['RESULT_FOLDER'], file.filename.replace('.pdf', '.xlsx'))
    try:
        extract_pdf_to_excel(filepath, excel_path)
    except Exception as e:
        return f"Error processing file: {str(e)}", 500
    
    return render_template('search.html', filename=excel_path)

@app.route('/search', methods=['POST'])
def search_results():
    subject_code = request.form.get('subject_code')
    pass_fail = request.form.get('pass_fail')
    filename = request.form.get('filename')
    
    if not subject_code or not pass_fail or not filename:
        return "Missing data", 400
    
    # Load the Excel file and filter data
    df = pd.read_excel(filename)
    filtered_df = df[df['Subject_Code'] == subject_code]  # Note the updated column name
    
    if pass_fail == 'pass':
        filtered_df = filtered_df[filtered_df['Result'] == 'Pass']
    elif pass_fail == 'fail':
        filtered_df = filtered_df[filtered_df['Result'] == 'Fail']
    
    # Save filtered results to a temporary file
    output = BytesIO()
    filtered_df.to_excel(output, index=False)
    output.seek(0)
    
    return send_file(
        output,
        as_attachment=True,
        download_name=f"filtered_results_{subject_code}.xlsx"
    )

if __name__ == '__main__':
    app.run(debug=True)