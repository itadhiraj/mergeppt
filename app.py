import os
import io
from flask import Flask, request, send_file, render_template, jsonify
from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

app = Flask(__name__)

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')

if not os.path.exists(UPLOAD_FOLDER):
    os.makedirs(UPLOAD_FOLDER)

def copy_shape(source_shape, target_slide):
    """
    Safely copy shapes (textboxes, images, shapes) without corrupting PPTX XML structure.
    """
    # 1. Copy Picture/Images
    if source_shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
        image_bytes = io.BytesIO(source_shape.image.blob)
        target_slide.shapes.add_picture(
            image_bytes, 
            source_shape.left, 
            source_shape.top, 
            source_shape.width, 
            source_shape.height
        )
    # 2. Copy Text / Textboxes
    elif source_shape.has_text_frame:
        new_box = target_slide.shapes.add_textbox(
            source_shape.left, 
            source_shape.top, 
            source_shape.width, 
            source_shape.height
        )
        tf_src = source_shape.text_frame
        tf_dst = new_box.text_frame
        tf_dst.word_wrap = tf_src.word_wrap
        
        # Paragraphs aur runs ka style aur formatting retain karein
        for p_idx, p_src in enumerate(tf_src.paragraphs):
            p_dst = tf_dst.paragraphs[0] if p_idx == 0 else tf_dst.add_paragraph()
            p_dst.alignment = p_src.alignment
            for r_src in p_src.runs:
                r_dst = p_dst.add_run()
                r_dst.text = r_src.text
                r_dst.font.bold = r_src.font.bold
                r_dst.font.italic = r_src.font.italic
                r_dst.font.size = r_src.font.size
                r_dst.font.color.rgb = r_src.font.color.rgb if r_src.font.color and r_src.font.color.type else None
                r_dst.font.name = r_src.font.name

def merge_powerpoint_clean(file_paths, output_path):
    """
    Clean merge that builds native slides without XML tree manipulation.
    """
    merged_prs = Presentation(file_paths[0])
    
    for file_path in file_paths[1:]:
        prs = Presentation(file_path)
        for slide in prs.slides:
            # Blank slide add karein
            blank_layout = merged_prs.slide_layouts[6] if len(merged_prs.slide_layouts) > 6 else merged_prs.slide_layouts[0]
            new_slide = merged_prs.slides.add_slide(blank_layout)
            
            # Shapes copy karein
            for shape in slide.shapes:
                try:
                    copy_shape(shape, new_slide)
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

        # Clean merge trigger karein
        merge_powerpoint_clean(saved_paths, output_path)
        
        # Temp cleanup
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
