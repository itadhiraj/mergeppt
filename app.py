import os
import gc
import io
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

# HREO: Yahan __name__ me double underscore lagana mandatory hai
app = Flask(__name__)

# Render par memory/storage handle karne ke liye limits badhana (Max 500MB)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_optimized(file_paths, output_path):
    """
    Linux / Render par sabhi presentations ki images aur contents ko retain karte hue merge karna.
    """
    try:
        # Pehli presentation ko base presentation ki tarah open karein
        merged_pres = Presentation(file_paths[0])
        
        # Baaki sabhi PPTs ki slides ko ek-ek karke base presentation me add karein
        for i, file_path in enumerate(file_paths[1:]):
            try:
                current_pres = Presentation(file_path)
                for slide in current_pres.slides:
                    # Slide layout select karein (Blank layout index 6 ya 0)
                    blank_layout = merged_pres.slide_layouts[6] if len(merged_pres.slide_layouts) > 6 else merged_pres.slide_layouts[0]
                    new_slide = merged_pres.slides.add_slide(blank_layout)
                    
                    for shape in slide.shapes:
                        try:
                            # Agar shape image hai, toh use binary stream se naye slide me add karein
                            if shape.shape_type == 13:  # 13 matlab Picture shape hota hai
                                image_bytes = shape.image.blob
                                image_stream = io.BytesIO(image_bytes)
                                
                                # Position aur dimensions preserve karte hue image add karein
                                left = shape.left
                                top = shape.top
                                width = shape.width
                                height = shape.height
                                
                                new_slide.shapes.add_picture(image_stream, left, top, width, height)
                            else:
                                # Baaki shapes (text boxes, tables, etc.) ke elements clone karein
                                el = shape.element
                                new_slide.shapes._spTree.append(el)
                        except Exception:
                            pass
                            
                del current_pres
                if i % 5 == 0:
                    gc.collect()
                    
            except Exception as e:
                print(f"File skip hui ({file_path}): {str(e)}")
                continue
                
        # Merged presentation ko save karein
        merged_pres.save(output_path)
        gc.collect()
        
    except Exception as e:
        raise Exception(f"Merge process fail ho gayi: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge_files():
    if 'files' not in request.files:
        return jsonify({"error": "Koi file select nahi ki gayi"}), 400
        
    files = request.files.getlist('files')
    if len(files) > 65:
        return jsonify({"error": "Maximum 60-65 files allowed hain ek baar mein"}), 400

    saved_paths = []
    output_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, "Merged_Presentation.pptx"))
    
    try:
        if os.path.exists(output_path):
            os.remove(output_path)

        for file in files:
            if file and file.filename.endswith('.pptx'):
                file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
                file.save(file_path)
                saved_paths.append(file_path)

        if not saved_paths:
            return jsonify({"error": "Valid .pptx files nahi mili"}), 400

        merge_powerpoint_optimized(saved_paths, output_path)
        
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
                    
        return send_file(output_path, as_attachment=True, download_name="Merged_Presentation.pptx")
        
    except Exception as e:
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
        return jsonify({"error": f"Server Error: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=False, port=5000)import os
import gc
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

# HREO: Yahan __name__ me double underscore lagana mandatory hai
app = Flask(__name__)

# Render par memory/storage handle karne ke liye limits badhana (Max 500MB)
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_optimized(file_paths, output_path):
    """
    Linux / Render par images aur graphs ko retain karte hue presentations merge karna.
    """
    try:
        # Pehli presentation ko base presentation ki tarah open karein
        merged_pres = Presentation(file_paths[0])
        
        # Baaki sabhi PPTs ki slides ko ek-ek karke base presentation me add karein
        for i, file_path in enumerate(file_paths[1:]):
            try:
                current_pres = Presentation(file_path)
                for slide in current_pres.slides:
                    # Slide layout select karein (Blank layout index 6 ya 0)
                    blank_layout = merged_pres.slide_layouts[6] if len(merged_pres.slide_layouts) > 6 else merged_pres.slide_layouts[0]
                    new_slide = merged_pres.slides.add_slide(blank_layout)
                    
                    # Shapes, Images aur Graphs/Charts ke elements ko safely clone karke add karna
                    for shape in slide.shapes:
                        try:
                            # XML element ko clone karke append karna
                            el = shape.element
                            new_slide.shapes._spTree.append(el)
                        except Exception:
                            pass
                            
                del current_pres
                if i % 5 == 0:
                    gc.collect()
                    
            except Exception as e:
                print(f"File skip hui ({file_path}): {str(e)}")
                continue
                
        # Merged presentation ko save karein
        merged_pres.save(output_path)
        gc.collect()
        
    except Exception as e:
        raise Exception(f"Merge process fail ho gayi: {str(e)}")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/merge', methods=['POST'])
def merge_files():
    if 'files' not in request.files:
        return jsonify({"error": "Koi file select nahi ki gayi"}), 400
        
    files = request.files.getlist('files')
    if len(files) > 65:
        return jsonify({"error": "Maximum 60-65 files allowed hain ek baar mein"}), 400

    saved_paths = []
    output_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, "Merged_Presentation.pptx"))
    
    try:
        if os.path.exists(output_path):
            os.remove(output_path)

        for file in files:
            if file and file.filename.endswith('.pptx'):
                file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
                file.save(file_path)
                saved_paths.append(file_path)

        if not saved_paths:
            return jsonify({"error": "Valid .pptx files nahi mili"}), 400

        merge_powerpoint_optimized(saved_paths, output_path)
        
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
                    
        return send_file(output_path, as_attachment=True, download_name="Merged_Presentation.pptx")
        
    except Exception as e:
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
        return jsonify({"error": f"Server Error: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=False, port=5000)
