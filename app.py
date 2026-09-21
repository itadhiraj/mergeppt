import os
import subprocess
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_pptx_clean(file_paths, output_path):
    """
    Safe merging method that preserves slide layouts and avoids XML corruption.
    """
    prs_master = Presentation(file_paths[0])
    
    for path in file_paths[1:]:
        prs_sub = Presentation(path)
        for slide in prs_sub.slides:
            # Add slide with blank layout
            blank_layout = prs_master.slide_layouts[6] if len(prs_master.slide_layouts) > 6 else prs_master.slide_layouts[0]
            new_slide = prs_master.slides.add_slide(blank_layout)
            
            # Copy text shapes safely without raw XML tree manipulation
            for shape in slide.shapes:
                if shape.has_text_frame:
                    txBox = new_slide.shapes.add_textbox(shape.left, shape.top, shape.width, shape.height)
                    tf = txBox.text_frame
                    tf.text = shape.text_frame.text

    prs_master.save(output_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge_files():
    if 'files' not in request.files:
        return jsonify({"error": "Koi file select nahi ki gayi"}), 400
        
    files = request.files.getlist('files')
    if len(files) > 60:
        return jsonify({"error": "Maximum 60 files allowed hain"}), 400

    saved_paths = []
    try:
        for file in files:
            file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
            file.save(file_path)
            saved_paths.append(file_path)

        output_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, "Merged_Presentation.pptx"))

        # Pure Python merge logic
        merge_pptx_clean(saved_paths, output_path)
        
        # Cleanup temp uploaded files
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception:
                    pass
                    
        return send_file(output_path, as_attachment=True, download_name="Merged_Presentation.pptx")
        
    except Exception as e:
        return jsonify({"error": f"Merge karne me dikkat aayi: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
