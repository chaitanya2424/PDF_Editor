from flask import Flask, redirect, render_template, request, flash, url_for
from werkzeug.utils import secure_filename
import os
from pdf2docx import Converter
import docx2pdf
import pythoncom

UPLOAD_FOLDER = 'upload'
ALLOWED_EXTENSIONS = {'pdf', 'docx'}

app = Flask(__name__)
app.secret_key = "hrjthrehtrhy"
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def convert(filename, operation):
    try:
        # Initialize COM library for docx2pdf conversion (Windows only)
        pythoncom.CoInitialize()

        if operation == "1" and filename.endswith(".pdf"):
            # Ensure 'static/docx' folder exists for output DOCX files
            upload_dir = os.path.join(app.root_path, 'static', 'docx')
            os.makedirs(upload_dir, exist_ok=True)

            # Construct full paths for input PDF and output DOCX files
            pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            docx_filename = filename.replace(".pdf", ".docx")
            docx_path = os.path.join(upload_dir, docx_filename)

            # Check if the input PDF file exists
            if not os.path.exists(pdf_path):
                return f"File not found: {pdf_path}"

            # Perform the conversion from PDF to DOCX
            cv = Converter(pdf_path)
            cv.convert(docx_path, start=0, end=None)
            cv.close()
            return docx_filename
        
        elif operation == "2" and filename.endswith(".docx"):
            # Ensure 'static/pdf' folder exists for output PDF files
            upload_dir = os.path.join(app.root_path, 'static', 'pdf')
            os.makedirs(upload_dir, exist_ok=True)

            # Construct full paths for input DOCX and output PDF files
            docx_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            pdf_filename = filename.replace(".docx", ".pdf")
            pdf_path = os.path.join(upload_dir, pdf_filename)

            # Check if the input DOCX file exists
            if not os.path.exists(docx_path):
                return f"File not found: {docx_path}"

            # Perform the conversion from DOCX to PDF
            docx2pdf.convert(docx_path, pdf_path)
            return pdf_filename
        
        else:
            return "Invalid operation or file type selected"

    except PermissionError as e:
        return f"Permission denied: {e}"
    except Exception as e:
        return f"An error occurred: {e}"
    finally:
        # Uninitialize COM library
        pythoncom.CoUninitialize()

@app.route("/edit", methods=["POST", "GET"])
def edit():
    if request.method == "POST":
        operation = request.form.get("operation")
        if 'fileUpload' not in request.files:
            flash("No file selected", "danger")
            return redirect(request.url)
        file = request.files['fileUpload']
        if file.filename == '':
            flash('No selected file', "danger")
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            result = convert(filename, operation)
            if result.startswith("An error occurred") or result.startswith("Permission denied"):
                flash(result, "danger")
            elif result == "Invalid operation or file type selected":
                flash(result, "danger")
            else:
                if operation == "1" and result.endswith(".docx"):
                    file_url = url_for("static", filename=f'docx/{result}')
                    file_type = "DOCX"
                elif operation == "2" and result.endswith(".pdf"):
                    file_url = url_for("static", filename=f'pdf/{result}')
                    file_type = "PDF"
                else:
                    flash("Invalid operation or file type selected", "danger")
                    return redirect(request.url)
                flash(f"Your {file_type} file is ready and available <a href='{file_url}' target='_blank'>here!</a>", "success")
    return render_template("edit.html")

@app.route("/")
def home():
    return render_template("l.html")

if __name__ == '__main__':
    app.run(debug=True, port=7909)
