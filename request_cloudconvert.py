#!/usr/bin/env python3
# -*- coding: utf-8 -*-

import os
import io
import uuid
import logging
import requests
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import cloudconvert

# ================= 配置 =================
logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# 你提供的 CloudConvert API Key（JWT）——按你的要求“硬编码”
CLOUDCONVERT_API_KEY = (
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9."
    "eyJhdWQiOiIxIiwianRpIjoiZmQzOGY5MTZjOGYwYjcwZWUzMjg0NDk1ODhlODM3MjkxYmRmYjgwMWI0NTg1OTQ0NmU0MjE4YWMyNWU2NWUyYzI4YjRmYTNiZWExZDQzYTUiLCJpYXQiOjE3NTc5MjI4NjYuNTQ0NjI5LCJuYmYiOjE3NTc5MjI4NjYuNTQ0NjMxLCJleHAiOjQ5MTM1OTY0NjYuNTM5NTU4LCJzdWIiOiI3Mjg3NDIyOSIsInNjb3BlcyI6WyJ1c2VyLnJlYWQiLCJ1c2VyLndyaXRlIiwidGFzay5yZWFkIiwidGFzay53cml0ZSIsIndlYmhvb2sucmVhZCIsIndlYmhvb2sud3JpdGUiLCJwcmVzZXQucmVhZCIsInByZXNldC53cml0ZSJdfQ."
    "ndVsCfxNcTN53vVs2G1qm6TO2Eb9zi_bZZpbRPFj9eUC9ieke6GjE25TMFwR_w1ncPj8H0ZZqTNRWlxt3iYMrQuHtwbeCVPRvIjL_5ldUZRagFjztig3ggP2RGuPnkZzC61Wvv8lv5Yhg3IDd8E1QjyI30_1U_O73n8lwE-5tCvR0SVfeo0NODtCZouC0_WkfpDkhZopKNf9xSuAXO7fOqCg8PdV9YdFFQcBBCfn_FLAbhT9Itkmo9v5DjPbRVJMIWRKJqOBsJMvxQumlU18ct_2Ld1yLJaoaiGhFUrprrYy85RR0h2Xwz7ls_vuTRPF0i24ngMbyv0SSsmtMYBnw8CWZmNT1WKcJRAq4OR6J3iQztQ6a4JWOsA5pzqtenYCKgc_Ig8smnnETOJU6Q_MA7jSMYdCY_5cfCqpT71EIhOmS_uPFd3AJsJ1Z0wYRNEUHnkV7pux0Oy1iDo-dwooH1kUHNKIWfewrt324vI_nmCGml-QITwckmXR9OVkV2LvXWCJY1rPNfcOIOG4ReRi-q1eZ-052ZEHn6bXfyYg5EJ-2Wm00BgJv7aTyyiVDNUDx1LvtxLEPz4BQdpFJ_-YJhVXX9CokCdgOJDrYuwCAVYaYP927NURjboiSvXsGGIhw93JhhRfU6-cmnZ8gSZzSTWWKMDKliVf62IDLH2Sk2M"
)

# 非沙盒（生产）环境；如需沙盒可加 sandbox=True
cloudconvert.configure(api_key=CLOUDCONVERT_API_KEY)


def pdf_to_pptx_cloudconvert(input_pdf_path: str, output_pptx_path: str) -> None:
    """
    使用 CloudConvert 将本地 PDF 转成 PPTX 并保存到 output_pptx_path。
    失败抛异常。
    """
    # 1) 创建 Job：import/upload -> convert -> export/url
    job = cloudconvert.Job.create(payload={
        "tasks": {
            "import-my-file": {"operation": "import/upload"},
            "convert-my-file": {
                "operation": "convert",
                "input": "import-my-file",
                "input_format": "pdf",
                "output_format": "pptx",
                # 可选参数（按需打开）：
                # "engine": "libreoffice",
                # "page_range": "1-10",
            },
            "export-my-file": {
                "operation": "export/url",
                "input": "convert-my-file"
            }
        }
    })

    # 2) 上传你刚保存到磁盘的 PDF
    import_task = next(t for t in job["tasks"] if t["name"] == "import-my-file")
    cloudconvert.Task.upload(file_name=os.path.abspath(input_pdf_path), task=import_task)

    # 3) 等待 Job 完成
    job = cloudconvert.Job.wait(id=job["id"])  # 阻塞直到完成/失败

    # 4) 找到导出 URL，下载保存为 output_pptx_path
    export_task = next(
        t for t in job["tasks"]
        if t["operation"] == "export/url" and t["status"] == "finished"
    )
    files = export_task["result"]["files"]
    if not files:
        raise RuntimeError("CloudConvert 导出未返回文件")

    file_url = files[0]["url"]
    os.makedirs(os.path.dirname(os.path.abspath(output_pptx_path)), exist_ok=True)
    with requests.get(file_url, stream=True, timeout=600) as r:
        r.raise_for_status()
        with open(output_pptx_path, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                if chunk:
                    f.write(chunk)


def convert_pdf_to_pptx(input_pdf_path: str, output_pptx_path: str) -> bool:
    """
    封装成返回 True/False，方便在 Flask 路由里调用。
    """
    try:
        pdf_to_pptx_cloudconvert(input_pdf_path, output_pptx_path)
        logging.info("转换成功: %s -> %s", input_pdf_path, output_pptx_path)
        return True
    except Exception as e:
        logging.exception("CloudConvert 转换失败: %s", e)
        return False


@app.route('/convert', methods=['POST'])
def handle_conversion():
    """
    接收一个名为 'file' 的 PDF，调用 CloudConvert 转 PPTX，并把 PPTX 作为附件返回。
    """
    if 'file' not in request.files:
        return jsonify({'error': '请求中未找到文件部分'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': '无效的文件类型，请上传 PDF 文件。'}), 400

    # 临时保存上传 PDF 到本地，再上传到 CloudConvert
    unique_id = uuid.uuid4().hex
    input_filename = f"{unique_id}.pdf"
    output_filename = f"{unique_id}.pptx"
    input_path = os.path.abspath(os.path.join(UPLOAD_FOLDER, input_filename))
    output_path = os.path.abspath(os.path.join(OUTPUT_FOLDER, output_filename))

    try:
        file.save(input_path)

        # 可选：确保写盘完成，避免刚写入就读取导致截断
        try:
            with open(input_path, 'ab', buffering=0) as f:
                os.fsync(f.fileno())
        except Exception:
            pass

        success = convert_pdf_to_pptx(input_path, output_path)
    finally:
        # 清理上传的 PDF
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except Exception:
            pass

    if not success:
        return jsonify({'error': '文件转换失败，请检查服务器日志。'}), 500

    # 成功则把生成的 PPTX 直接返回，并删除本地文件
    try:
        download_filename = f"{os.path.splitext(file.filename)[0]}.pptx"
        return_data = io.BytesIO()
        with open(output_path, 'rb') as fo:
            return_data.write(fo.read())
        return_data.seek(0)

        try:
            os.remove(output_path)
        except Exception:
            pass

        return send_file(
            return_data,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            download_name=download_filename,
            as_attachment=True
        )
    except Exception as e:
        logging.error("发送文件时出错: %s", e)
        return jsonify({'error': '文件转换成功但发送时出错。'}), 500


if __name__ == '__main__':
    # 运行服务：POST /convert  上传 form-data 字段名 file=your.pdf
    app.run(host='0.0.0.0', port=5000, debug=True)