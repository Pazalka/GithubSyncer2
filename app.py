import os
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
from excel_processor import process_excel_files
import tempfile

app = Flask(__name__)
app.secret_key = os.urandom(24)

# Configure upload settings
UPLOAD_FOLDER = tempfile.gettempdir()
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    try:
        uploaded_files = []
        
        # Collect all uploaded files
        files_found = False
        for key in request.files:
            file = request.files[key]
            if file and file.filename:
                if not allowed_file(file.filename):
                    return jsonify({'error': f'סוג קובץ לא נתמך עבור {file.filename}'}), 400
                
                filename = secure_filename(file.filename)
                filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
                file.save(filepath)
                uploaded_files.append(filepath)
                files_found = True
        
        if not files_found:
            return jsonify({'error': 'לא נבחרו קבצים לעיבוד'}), 400

        # Process the files
        output_file = process_excel_files(uploaded_files)
        
        # Clean up uploaded files
        for filepath in uploaded_files:
            try:
                os.remove(filepath)
            except:
                pass

        return send_file(
            output_file,
            as_attachment=True,
            download_name='processed_output.xlsx',
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )

    except Exception as e:
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
