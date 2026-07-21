import os
import uuid
import tempfile
import requests
import json
from io import BytesIO
from flask import Flask, request, send_file, jsonify
from pptx import Presentation
from pptx.util import Pt

app = Flask(__name__)


def get_file_from_request(param_name):
    """
    从请求中获取文件对象，优先从上传文件获取，其次从JSON中的URL下载
    返回 (file-like object, filename)
    """
    # 1. 尝试从上传的文件中获取（multipart/form-data）
    file = request.files.get(param_name)
    if file and file.filename:
        return file, file.filename

    # 2. 尝试从 JSON 中获取 URL
    if request.is_json:
        data = request.get_json()
        url = data.get(param_name) if isinstance(data, dict) else None
        if url:
            try:
                resp = requests.get(url, timeout=30)
                resp.raise_for_status()
                # 从 URL 中提取文件名（带扩展名）
                filename = url.split('/')[-1].split('?')[0] or 'file.pptx'
                return BytesIO(resp.content), filename
            except Exception as e:
                print(f"下载文件失败: {e}")
                return None, None
    return None, None


@app.route('/')
def index():
    return 'PPT Service is running! 使用 /extract_template 或 /generate_ppt 接口。'


@app.route('/openapi.json', methods=['GET'])
def openapi_spec():
    """返回 OpenAPI 规范（如果使用云侧插件自动解析）"""
    spec = {
        "openapi": "3.0.0",
        "info": {"title": "PPT 工具", "version": "1.0.0"},
        "servers": [{"url": "https://ppt-production-cca7.up.railway.app"}],
        "paths": {
            "/extract_template": {
                "post": {
                    "operationId": "extract_template",
                    "requestBody": {
                        "content": {
                            "multipart/form-data": {
                                "schema": {
                                    "type": "object",
                                    "properties": {"file": {"type": "string", "format": "binary"}},
                                    "required": ["file"]
                                }
                            }
                        }
                    },
                    "responses": {"200": {"description": "OK"}}
                }
            },
            "/generate_ppt": {
                "post": {
                    "operationId": "generate_ppt",
                    "requestBody": {
                        "content": {
                            "multipart/form-data": {
                                "schema": {
                                    "type": "object",
                                    "properties": {
                                        "template": {"type": "string", "format": "binary"},
                                        "slides_data": {"type": "string"}
                                    },
                                    "required": ["template", "slides_data"]
                                }
                            }
                        }
                    },
                    "responses": {"200": {"description": "OK"}}
                }
            }
        }
    }
    return jsonify(spec)


@app.route('/extract_template', methods=['POST'])
def extract_template():
    """接收风格PPT文件（上传或URL），返回去除了所有幻灯片的空白模板"""
    file_obj, filename = get_file_from_request('file')
    if not file_obj:
        return jsonify({'error': 'No file'}), 400

    # 保存上传/下载的文件到临时路径
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        tmp.write(file_obj.read())
        src_path = tmp.name

    try:
        prs = Presentation(src_path)
        # 删除所有幻灯片，保留母版和版式
        while len(prs.slides) > 0:
            rId = prs.slides._sldIdLst[0].get('r:id')
            if rId is None:
                rId = prs.slides._sldIdLst[0].get(
                    '{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
            prs.part.drop_rel(rId)
            del prs.slides._sldIdLst[0]

        template_path = f'/tmp/{uuid.uuid4()}.pptx'
        prs.save(template_path)
    finally:
        os.unlink(src_path)  # 删除源文件

    return send_file(template_path, as_attachment=True, download_name='template.pptx')


@app.route('/generate_ppt', methods=['POST'])
def generate_ppt():
    """接收模板文件和结构化内容JSON，生成最终PPT"""
    # 获取模板文件
    template_obj, _ = get_file_from_request('template')
    if not template_obj:
        return jsonify({'error': 'Missing template file'}), 400

    # 获取 slides_data（可能在 JSON 或 form-data 中）
    if request.is_json:
        data = request.get_json()
        slides_data = data.get('slides_data', '[]')
    else:
        slides_data = request.form.get('slides_data', '[]')

    # 保存模板到临时文件
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pptx') as tmp:
        tmp.write(template_obj.read())
        template_path = tmp.name
    try:
        slides = json.loads(slides_data)
        prs = Presentation(template_path)
        layouts = prs.slide_layouts

        for slide_data in slides:
            layout_name = slide_data.get('layout', '标题和内容')
            matched_layout = None
            for layout in layouts:
                if layout_name in layout.name:
                    matched_layout = layout
                    break
            if not matched_layout:
                matched_layout = layouts[1]  # 默认使用第二个版式

            slide = prs.slides.add_slide(matched_layout)

            # 设置标题
            if slide.shapes.title:
                slide.shapes.title.text = slide_data.get('title', '')

            # 填充正文占位符（索引为1）
            for shape in slide.placeholders:
                if shape.placeholder_format.idx == 1:
                    tf = shape.text_frame
                    tf.clear()
                    for point in slide_data.get('bullet_points', []):
                        p = tf.add_paragraph()
                        p.text = point
                        p.level = 0
                    # 添加固定文本（小字）
                    for fixed_text in slide_data.get('fixed_texts', []):
                        p = tf.add_paragraph()
                        p.text = fixed_text
                        p.font.size = Pt(10)
                    break

        output_path = f'/tmp/{uuid.uuid4()}.pptx'
        prs.save(output_path)
    finally:
        os.unlink(template_path)  # 删除临时模板

    return send_file(output_path, as_attachment=True, download_name='generated.pptx')


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
