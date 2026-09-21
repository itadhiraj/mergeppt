import os
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

# HREO: Yahan __name__ me double underscore lagana mandatory hai
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_native(file_paths, output_path):
    """
    Linux / Render par python-pptx ka use karke slides ko standard rule se merge karna.
    """
    # Pehli presentation ko base presentation ki tarah open karein
    merged_pres = Presentation(file_paths[0])
    
    # Baaki sabhi PPTs ki slides ko ek-ek karke base presentation me add karein
    for file_path in file_paths[1:]:
        current_pres = Presentation(file_path)
        for slide in current_pres.slides:
            # Slide layout select karein
            slide_layout = merged_pres.slide_layouts[0]
            new_slide = merged_pres.slides.add_slide(slide_layout)
            
            # Shapes ko copy karne ke liye XML element append karein
            for shape in slide.shapes:
                new_slide.shapes._spTree.append(shape.element)
                
    # Merged presentation ko save karein
    merged_pres.save(output_path)

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
        # Files ko strict absolute path ke sath uploads folder me save karna
        for file in files:
            file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
            file.save(file_path)
            saved_paths.append(file_path)

        output_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, "Merged_Presentation.pptx"))

        # Main merge function trigger karna
        merge_powerpoint_native(saved_paths, output_path)
        
        # Kaam hone ke baad temporary uploaded single files ko delete karna
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
                    
        return send_file(output_path, as_attachment=True, download_name="Merged_Presentation.pptx")
        
    except Exception as e:
        return jsonify({"error": f"Merge karne me dikkat aayi: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)import os
from flask import Flask, request, send_file, render_template, jsonify
import win32com.client
import pythoncom  # Threads ko initialize karne ke liye zaroori hai

# HREO: Yahan __name__ me double underscore lagana mandatory hai
app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_native(file_paths, output_path):
    """
    Windows PowerPoint API ka use karke slides ko standard rule se merge karna.
    """
    # Flask thread ke andar Windows COM library ko chalu karna
    pythoncom.CoInitialize()
    
    try:
        # Parde ke peeche PowerPoint application initialize karein
        ppt_instance = win32com.client.Dispatch("PowerPoint.Application")
        
        # Pehli PPT ko base presentation ki tarah open karein (WithWindow=False se background me chalega)
        main_pres = ppt_instance.Presentations.Open(file_paths[0], WithWindow=False)
        
        # Baaki sabhi PPTs ko ek-ek karke loop me end me insert karein
        for file_path in file_paths[1:]:
            slide_count = main_pres.Slides.Count
            # InsertFromFile asli software method hai jo fonts aur layouts bilkul kharab nahi karta
            main_pres.Slides.InsertFromFile(file_path, slide_count, 1, -1)
            
        # Merged presentation ko save aur close karein
        main_pres.SaveAs(output_path)
        main_pres.Close()
        ppt_instance.Quit()
        
    finally:
        # Har haal me memory free karna safely
        pythoncom.CoUninitialize()

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
        # Files ko strict absolute path ke sath uploads folder me save karna
        for file in files:
            file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
            file.save(file_path)
            saved_paths.append(file_path)

        output_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, "Merged_Presentation.pptx"))

        # Main merge function trigger karna
        merge_powerpoint_native(saved_paths, output_path)
        
        # Kaam hone ke baad temporary uploaded single files ko delete karna
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
                    
        return send_file(output_path, as_attachment=True, download_name="Merged_Presentation.pptx")
        
    except Exception as e:
        return jsonify({"error": f"Merge karne me dikkat aayi: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
