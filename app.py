from flask import Flask, render_template, request
import pandas as pd
from docx import Document
import os
import re

app = Flask(__name__)

# Function to normalize WERS codes by removing special characters
def normalize_code(code):
    return re.sub(r'[-_]', ' ', code)

@app.route('/')
def upload_files():
    return render_template('upload.html')

@app.route('/upload', methods=['POST'])
def process_files():
    if 'excel_file' not in request.files or 'word_file' not in request.files or 'voci_excel_file' not in request.files:
        return "Missing file(s)"

    excel_file = request.files['excel_file']
    word_file = request.files['word_file']
    voci_excel_file = request.files['voci_excel_file']
    excel_header = int(request.form['excel_header'])
    voci_header = int(request.form['voci_header'])

    excel_path = os.path.join('uploads', excel_file.filename)
    word_path = os.path.join('uploads', word_file.filename)
    voci_excel_path = os.path.join('uploads', voci_excel_file.filename)

    excel_file.save(excel_path)
    word_file.save(word_path)
    voci_excel_file.save(voci_excel_path)

    try:
        excel_data = pd.read_excel(excel_path, header=excel_header-1)
    except Exception as e:
        return f"Error reading Excel file: {e}"

    column_name = 'Feature WERS Code'
    if column_name not in excel_data.columns:
        return f"Column '{column_name}' not found in the Excel file."

    codes_from_excel = excel_data[column_name].dropna().astype(str).tolist()

    try:
        doc = Document(word_path)
    except Exception as e:
        return f"Error reading Word file: {e}"

    text_content = [paragraph.text for paragraph in doc.paragraphs]
    full_text = ' '.join(text_content)

    # Normalize WERS codes found in Excel
    normalized_codes_from_excel = [normalize_code(code) for code in codes_from_excel]

    # Find codes in Word document (both original and normalized)
    codes_found_in_word = []
    for code in codes_from_excel:
        normalized_code = normalize_code(code)
        if code in full_text or normalized_code in full_text:
            codes_found_in_word.append(code)

    try:
        voci_data = pd.read_excel(voci_excel_path, header=voci_header-1)
    except Exception as e:
        return f"Error reading VOCI Excel file: {e}"

    required_columns = ['WERS Code', 'Sales Code']
    for col in required_columns:
        if col not in voci_data.columns:
            return f"Column '{col}' not found in the VOCI Excel file."

    voci_codes = voci_data[['WERS Code', 'Sales Code']].dropna().astype(str)

    # Create a mapping from Sales Code to WERS Codes
    sales_code_to_wers = {}
    for index, row in voci_codes.iterrows():
        wers_code = row['WERS Code']
        sales_code = row['Sales Code']
        if sales_code not in sales_code_to_wers:
            sales_code_to_wers[sales_code] = []
        sales_code_to_wers[sales_code].append(wers_code)

    results = []
    for code in codes_found_in_word:
        normalized_code = normalize_code(code)
        matching_sales_codes = voci_codes[voci_codes['WERS Code'].apply(normalize_code) == normalized_code]['Sales Code'].tolist()
        if matching_sales_codes:
            chosen_sales_code = matching_sales_codes[0]  # Choose the first Sales code found
            results.append((code, chosen_sales_code))
        else:
            results.append((code, code))  # If no Sales code matches, use the WERS code itself

    return render_template('results.html', results=results)

if __name__ == '__main__':
    app.run(debug=True)
