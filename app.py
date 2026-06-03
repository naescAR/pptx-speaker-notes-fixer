import os
import io
from flask import Flask, request, render_template, send_file, jsonify
from werkzeug.utils import secure_filename
from pptx import Presentation
from pptx.util import Inches

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50MB max upload

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file part'}), 400
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': 'No selected file'}), 400
        
    if file and file.filename.endswith('.pptx'):
        try:
            # Load the presentation from the uploaded file stream
            file_stream = io.BytesIO(file.read())
            prs = Presentation(file_stream)
            
            # The Notes Master page is typically 8.5 x 11 inches portrait
            # We will center the Slide Image and Notes body to have 6.5" width and 1.0" left margin
            TARGET_WIDTH = Inches(6.5)
            TARGET_LEFT = Inches(1.0)
            
            if prs.notes_master:
                for shape in prs.notes_master.shapes:
                    # Center the Slide Image and Notes text box
                    if "Slide Image" in shape.name or "Notes" in shape.name:
                        shape.width = TARGET_WIDTH
                        shape.left = TARGET_LEFT
                        
            # Save the modified presentation to a new stream
            output_stream = io.BytesIO()
            prs.save(output_stream)
            output_stream.seek(0)
            
            original_name = os.path.splitext(file.filename)[0]
            fixed_filename = f"{original_name}_Fixed.pptx"
            
            return send_file(
                output_stream,
                as_attachment=True,
                download_name=fixed_filename,
                mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
            )
            
        except Exception as e:
            return jsonify({'error': str(e)}), 500
            
    return jsonify({'error': 'Invalid file format. Please upload a .pptx file.'}), 400

if __name__ == '__main__':
    # Use port 5001 to avoid conflicts
    app.run(debug=True, port=5001)
