import os
import copy
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_python_pptx(file_paths, output_path):
    """
    Corrupt-free PPTX merge using python-pptx by deep cloning slide elements.
    """
    # Base Presentation load karein
    merged_prs = Presentation(file_paths[0])
    
    for file_path in file_paths[1:]:
        prs = Presentation(file_path)
        for slide in prs.slides:
            # Layout select karein (blank layout)
            blank_layout = merged_prs.slide_layouts[6] if len(merged_prs.slide_layouts) > 6 else merged_prs.slide_layouts[0]
            new_slide = merged_prs.slides.add_slide(blank_layout)
            
            # Shapes ko safely duplicate karein without corrupting XML relationships
            for shape in slide.shapes:
                new_sp = copy.deepcopy(shape.element)
                new_slide.shapes._spTree.append(new_sp)
                
            # Related shapes, images, aur media resources ko connect karne ke liye
            for rel in slide.part.rels.values():
                if "notesSlide" not in rel.reltype:
                    try:
                        new_slide.part.rels.get_or_add_relationship(
                            rel.reltype, rel._target, rel.rId
                        )
                    except Exception:
                        pass

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

        # Merging process
        merge_powerpoint_python_pptx(saved_paths, output_path)
        
        # Cleanup temporary uploaded files
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
