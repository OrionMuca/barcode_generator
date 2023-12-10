from flask import Flask, render_template, request, send_from_directory
import os
import pandas as pd
from barcode import Code128
from barcode.writer import ImageWriter
from reportlab.lib.units import cm, mm
from reportlab.pdfgen import canvas

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure the 'uploads' directory exists
if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)


def generate_barcode(data, output_path, barcode_width=60, barcode_height=10):
    for index, row in data.iterrows():
        code = str(row['Code'])
        description = str(row['Description'])
        barcode_code = f"{code} {description}"
        barcode_filename = f"{output_path}/barcode_{code}"

        # Generate barcode
        barcode = Code128(code, writer=ImageWriter())
        barcode.save(barcode_filename, options={'width': barcode_width, 'height': barcode_height, 'write_text': False})


def create_labels(data, output_path):
    pdf_filename = f"{output_path}/labels.pdf"
    label_width = 118  # Width of the label in mm
    label_height = 20  # Height of the label in mm
    barcode_width = 60  # Width of the barcode in mm
    barcode_height = 10  # Height of the barcode in mm

    c = canvas.Canvas(pdf_filename, pagesize=(label_width*mm, label_height*mm))

    for index, row in data.iterrows():
        code = str(row['Code'])
        description = str(row['Description'])
        price = str(row['Price'])
        barcode_filename = f"{output_path}/barcode_{code}.png"

        # Add label information to the PDF
        c.setStrokeColorRGB(0, 0, 0)  # Set border color to black
        c.rect(1*mm, 1*mm, (label_width-2)*mm, (label_height-2)*mm)  # Adjusted to fit within the printer specifications

        # Position text on the label
        c.setFont("Helvetica", 8)  # Set font size
        text_start_y = label_height - 5  # Adjusted to start a little below the top
        c.drawCentredString(label_width / 2 * mm, text_start_y*mm, f"{code}: {description}")
        c.drawCentredString(label_width / 2 * mm, (text_start_y - 4)*mm, f"Cmimi: {price}")

        # Calculate the starting x-coordinate to center the barcode
        barcode_start_x = (label_width - barcode_width) / 2

        # Draw barcode below the text
        c.drawInlineImage(barcode_filename, barcode_start_x*mm, 1*mm, width=barcode_width*mm, height=barcode_height*mm)

        # Move to the next page for the next label
        c.showPage()

    c.save()

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download', methods=['GET'])
def upload_file():
    if 'file' not in request.files:
        return render_template('index.html', error='No file part')

    file = request.files['file']

    if file.filename == '':
        return render_template('index.html', error='No selected file')

    if file:
        filename = os.path.join(app.config['UPLOAD_FOLDER'], 'input.xlsx')
        file.save(filename)

        # Read Excel data into a pandas DataFrame
        data = pd.read_excel(filename)

        # Generate barcodes
        generate_barcode(data, app.config['UPLOAD_FOLDER'])

        # Create labels and save as a PDF
        create_labels(data, app.config['UPLOAD_FOLDER'])

        return send_from_directory(app.config['UPLOAD_FOLDER'], 'labels.pdf', as_attachment=True)


if __name__ == '__main__':
    app.run()
