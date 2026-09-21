import os
import gc
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation

app = Flask(__name__)

# Render par memory/storage handle karne ke liye limits badhana
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # Max 500MB upload limit

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def merge_powerpoint_optimized(file_paths, output_path):
    """
    60+ PowerPoint files ko memory leak aur crash se bachate hue merge karna.
    """
    try:
        # Pehli presentation ko base banayein
        merged_pres = Presentation(file_paths[0])
        
        # Ek-ek karke baaki presentations ko load karke slides append karein
        for i, file_path in enumerate(file_paths[1:]):
            try:
                current_pres = Presentation(file_path)
                for slide in current_pres.slides:
                    # Slide layout layout index safe tarike se lena
                    slide_layout = merged_pres.slide_layouts[0]
                    new_slide = merged_pres.slides.add_slide(slide_layout)
                    
                    # Shapes ko safely copy karna
                    for shape in slide.shapes:
                        try:
                            new_slide.shapes._spTree.append(shape.element)
                        except Exception:
                            pass
                
                # Memory optimize karne ke liye har file ke baad garbage collector chalayein
                del current_pres
                if i % 5 == 0:
                    gc.collect()
                    
            except Exception as e:
                print(f"File skip hui ({file_path}): {str(e)}")
                continue
                
        # Final merged presentation save karein
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
        # Purani output file agar ho toh saaf karein
        if os.path.exists(output_path):
            os.remove(output_path)

        # Files ko save karna
        for file in files:
            if file and file.filename.endswith('.pptx'):
                file_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, file.filename))
                file.save(file_path)
                saved_paths.append(file_path)

        if not saved_paths:
            return jsonify({"error": "Valid .pptx files nahi mili"}), 400

        # Merge function call karna
        merge_powerpoint_optimized(saved_paths, output_path)
        
        # Temporary files delete karna
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
                    
        return send_file(output_path, as_attachment=True, download_name="Merged_Presentation.pptx")
        
    except Exception as e:
        # Error aane par bhi cleanup karna
        for path in saved_paths:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except:
                    pass
        return jsonify({"error": f"Server Error: {str(e)}"}), 500

if __name__ == '__main__':
    app.run(debug=False, port=5000)
