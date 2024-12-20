import os
import re
import pdfplumber
import pandas as pd
from flask import Flask, request, redirect, render_template, send_from_directory, url_for
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'pdf'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Allowed file extensions for upload


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Extract text from PDF


def extract_pdf_text(pdf_path):
    with pdfplumber.open(pdf_path) as pdf:
        full_text = ""
        for page in pdf.pages:
            full_text += page.extract_text() + "\n"  # Accumulate text from each page
    return full_text

# Extract product data from extracted text


def extract_product_data_from_text(text):
    # Define regex patterns to capture size and quantity information
    size_pattern = re.compile(r"\d{1,2}-\d{1,2}Y")  # Match sizes like 8-9Y
    # Match quantity like 45 Singles
    quantity_pattern = re.compile(r"(\d+)\sSingles")

    # Define the sizes you are looking for
    size_labels = ['3-4', '5-6', '7-8', '8-9',
                   '9-10', '10-11', '12-13', '14-15']
    # Initialize with zero quantities for each size
    size_quantity_dict = {size: 0 for size in size_labels}

    lines = text.splitlines()  # Split the full text into lines

    current_description = ""
    current_size = None

    for line in lines:
        line = line.strip()

        if line:  # If the line is not empty
            current_description += " " + line  # Accumulate the description

        # Check if the line contains a size
        size_match = size_pattern.search(current_description)
        if size_match:
            # Remove 'Y' from size (e.g., 8-9Y -> 8-9)
            current_size = size_match.group(0).replace('Y', '')

        # Check if the line contains a quantity
        quantity_match = quantity_pattern.search(current_description)
        if quantity_match and current_size:
            quantity = quantity_match.group(1)
            if current_size in size_quantity_dict:
                size_quantity_dict[current_size] = quantity

            # Reset for the next product
            current_description = ""
            current_size = None

    return size_quantity_dict

# Route for uploading files


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')

# Route for uploading and processing multiple PDFs


@app.route('/upload', methods=['POST'])
def upload_file():
    if 'pdf_files' not in request.files:
        return redirect(request.url)

    pdf_files = request.files.getlist('pdf_files')

    if not pdf_files or all(file.filename == '' for file in pdf_files):
        return redirect(request.url)

    # Store the uploaded files
    uploaded_files = []
    for pdf_file in pdf_files:
        if pdf_file and allowed_file(pdf_file.filename):
            filename = secure_filename(pdf_file.filename)
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            pdf_file.save(filepath)
            uploaded_files.append(filepath)

    # Process the uploaded PDFs
    all_extracted_data = []
    for pdf_path in uploaded_files:
        full_text = extract_pdf_text(pdf_path)
        extracted_data = extract_product_data_from_text(full_text)

        # Add the style (PDF name) to the extracted data
        extracted_data['STYLE'] = os.path.basename(pdf_path)
        all_extracted_data.append(extracted_data)

    # Create a DataFrame from the accumulated data
    if all_extracted_data:
        df = pd.DataFrame(all_extracted_data)

        # Save the extracted data to an Excel file
        excel_output_path = os.path.join(
            app.config['OUTPUT_FOLDER'], 'extracted_product_data.xlsx')
        save_to_excel(df, excel_output_path)

        return redirect(url_for('download_file', filename='extracted_product_data.xlsx'))

    return redirect(request.url)

# Route to serve the Excel file for download


@app.route('/download/<filename>')
def download_file(filename):
    return send_from_directory(app.config['OUTPUT_FOLDER'], filename)

# Function to save DataFrame to Excel


def save_to_excel(df, output_path):
    """Save the DataFrame to an Excel file."""
    df.to_excel(output_path, index=False)
    print(f"Data has been successfully saved to {output_path}")


if __name__ == '__main__':
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)
    if not os.path.exists(OUTPUT_FOLDER):
        os.makedirs(OUTPUT_FOLDER)
    app.run(debug=True)
