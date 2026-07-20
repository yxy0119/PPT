from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Pt
import tempfile
import os
import uuid

app = Flask(__name__)

@app.route('/extract_template', methods=['POST'])
def extract_template():
    """接收用户上传的风格PPT，返回一个空白模板（只含母版/版式）"""
    file = request.files.get('file')
    if not file:
        return jsonify({'error': 'No file'}), 400

    # 保存上传文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        file.save(tmp.name)
        src_path = tmp.name

    prs = Presentation(src_path)
    # 删除所有幻灯片，保留母版
    while len(prs.slides) > 0:
        rId = prs.slides._sldIdLst[0].get('r:id')
        if rId is None:
            rId = prs.slides._sldIdLst[0].get(
                '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
        prs.part.drop_rel(rId)
        del prs.slides._sldIdLst[0]

    # 保存空白模板
    template_id = str(uuid.uuid4())
    template_path = f'/tmp/{template_id}.pptx'
    prs.save(template_path)
    os.unlink(src_path)

    # 返回文件
    return send_file(template_path, as_attachment=True, download_name='template.pptx')

@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    """接收模板文件和结构化内容JSON，生成最终PPT"""
    template_file = request.files.get('template')
    slides_data = request.form.get('slides_data')  # JSON字符串

    if not template_file or not slides_data:
        return jsonify({'error': 'Missing data'}), 400

    # 保存模板文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        template_file.save(tmp.name)
        template_path = tmp.name

    import json
    slides = json.loads(slides_data)

    prs = Presentation(template_path)
    layouts = prs.slide_layouts

    for slide_data in slides:
        layout_name = slide_data.get('layout', '标题和内容')
        # 简单匹配版式名称
        matched_layout = None
        for layout in layouts:
            if layout_name in layout.name:
                matched_layout = layout
                break
        if not matched_layout:
            matched_layout = layouts[1]  # 默认第二个版式

        slide = prs.slides.add_slide(matched_layout)

        # 标题
        if slide.shapes.title:
            slide.shapes.title.text = slide_data.get('title', '')

        # 正文占位符（通常索引为1）
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1:
                tf = shape.text_frame
                tf.clear()
                for point in slide_data.get('bullet_points', []):
                    p = tf.add_paragraph()
                    p.text = point
                    p.level = 0
                # 固定文字
                for fixed in slide_data.get('fixed_texts', []):
                    p = tf.add_paragraph()
                    p.text = fixed
                    p.font.size = Pt(10)
                break

    output_path = f'/tmp/{uuid.uuid4()}.pptx'
    prs.save(output_path)
    os.unlink(template_path)

    return send_file(output_path, as_attachment=True, download_name='generated.pptx')

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
