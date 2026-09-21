import os
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_python_pptx(file_paths, output_path):
    """
    Linux/Render compatible PowerPoint merge function using python-pptx.
    """
    # Pehli PPT ko base presentation ki tarah load karein
    merged_prs = Presentation(file_paths[0])
    
    # Dusri sabhi PPTs ke slides ko base presentation me append karein
    for file_path in file_paths[1:]:
        prs = Presentation(file_path)
        for slide in prs.slides:
            # Blank slide layout select karein
            blank_slide_layout = merged_prs.slide_layouts[6] 
            new_slide = merged_prs.slides.add_slide(blank_slide_layout)
            
            # Shapes aur content ko copy karein
            for shape in slide.shapes:
                new_slide.shapes._spTree.insert_element_before(
                    shape.element, 'p:extLst'
                )
                
    merged_prs.save(output_path)

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

        # Linux compatible merge trigger karein
        merge_powerpoint_python_pptx(saved_paths, output_path)
        
        # Temp uploaded files cleanup
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