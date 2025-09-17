# -*- coding: utf-8 -*-
"""
一个统一的 PDF 到 PPTX 转换服务。
该服务集成了 Adobe, CloudConvert, ConvertAPI, 和 LibreOffice 四种转换引擎。

API Endpoints:
- POST /convert/adobe
- POST /convert/cloudconvert
- POST /convert/convertapi
- POST /convert/libreoffice

请求方式:
使用 multipart/form-data 发送 POST 请求, 包含一个名为 'file' 的 PDF 文件。

成功响应:
直接返回转换后的 .pptx 文件供浏览器下载。

失败响应:
返回一个包含错误信息的 JSON 对象。
"""
import os
import io
import uuid
import logging
import shutil
import tempfile
import subprocess
from functools import wraps

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# --- 各 API 客户端库导入 ---
# 1. Adobe
from adobe.pdfservices.operation.auth.service_principal_credentials import ServicePrincipalCredentials
from adobe.pdfservices.operation.exception.exceptions import ServiceApiException, ServiceUsageException, SdkException
from adobe.pdfservices.operation.io.cloud_asset import CloudAsset
from adobe.pdfservices.operation.io.stream_asset import StreamAsset
from adobe.pdfservices.operation.pdf_services import PDFServices
from adobe.pdfservices.operation.pdf_services_media_type import PDFServicesMediaType
from adobe.pdfservices.operation.pdfjobs.jobs.export_pdf_job import ExportPDFJob
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_ocr_locale import ExportOCRLocale
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_params import ExportPDFParams
from adobe.pdfservices.operation.pdfjobs.params.export_pdf.export_pdf_target_format import ExportPDFTargetFormat
from adobe.pdfservices.operation.pdfjobs.result.export_pdf_result import ExportPDFResult

# 2. CloudConvert
import cloudconvert
import requests

# 3. ConvertAPI
import convertapi

# ==============================================================================
# 1. 全局配置 (Configuration)
# ==============================================================================

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
logger = logging.getLogger("PDF2PPTX_Service")

app = Flask(__name__)
CORS(app)  # 允许跨域请求

# --- 文件夹配置 ---
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# --- API 密钥配置 ---
# 建议: 为了安全，最好将这些密钥存储在环境变量中，而不是硬编码在代码里。
# Adobe Credentials
ADOBE_CLIENT_ID = '9e6937018e1542439925726b28509327'
ADOBE_CLIENT_SECRET = 'p8e-JgpUaqdPWuHMkWDTc5rYsJ9kBex2hBjr'

# CloudConvert API Key
CLOUDCONVERT_API_KEY = (
    "eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9."
    "eyJhdWQiOiIxIiwianRpIjoiZmQzOGY5MTZjOGYwYjcwZWUzMjg0NDk1ODhlODM3MjkxYmRmYjgwMWI0NTg1OTQ0NmU0MjE4YWMyNWU2NWUyYzI4YjRmYTNiZWExZDQzYTUiLCJpYXQiOjE3NTc5MjI4NjYuNTQ0NjI5LCJuYmYiOjE3NTc5MjI4NjYuNTQ0NjMxLCJleHAiOjQ5MTM1OTY0NjYuNTM5NTU4LCJzdWIiOiI3Mjg3NDIyOSIsInNjb3BlcyI6WyJ1c2VyLnJlYWQiLCJ1c2VyLndyaXRlIiwidGFzay5yZWFkIiwidGFzay53cml0ZSIsIndlYmhvb2sucmVhZCIsIndlYmhvb2sud3JpdGUiLCJwcmVzZXQucmVhZCIsInByZXNldC53cml0ZSJdfQ."
    "ndVsCfxNcTN53vVs2G1qm6TO2Eb9zi_bZZpbRPFj9eUC9ieke6GjE25TMFwR_w1ncPj8H0ZZqTNRWlxt3iYMrQuHtwbeCVPRvIjL_5ldUZRagFjztig3ggP2RGuPnkZzC61Wvv8lv5Yhg3IDd8E1QjyI30_1U_O73n8lwE-5tCvR0SVfeo0NODtCZouC0_WkfpDkhZopKNf9xSuAXO7fOqCg8PdV9YdFFQcBBCfn_FLAbhT9Itkmo9v5DjPbRVJMIWRKJqOBsJMvxQumlU18ct_2Ld1yLJaoaiGhFUrprrYy85RR0h2Xwz7ls_vuTRPF0i24ngMbyv0SSsmtMYBnw8CWZmNT1WKcJRAq4OR6J3iQztQ6a4JWOsA5pzqtenYCKgc_Ig8smnnETOJU6Q_MA7jSMYdCY_5cfCqpT71EIhOmS_uPFd3AJsJ1Z0wYRNEUHnkV7pux0Oy1iDo-dwooH1kUHNKIWfewrt324vI_nmCGml-QITwckmXR9OVkV2LvXWCJY1rPNfcOIOG4ReRi-q1eZ-052ZEHn6bXfyYg5EJ-2Wm00BgJv7aTyyiVDNUDx1LvtxLEPz4BQdpFJ_-YJhVXX9CokCdgOJDrYuwCAVYaYP927NURjboiSvXsGGIhw93JhhRfU6-cmnZ8gSZzSTWWKMDKliVf62IDLH2Sk2M"
)

# ConvertAPI Secret
CONVERTAPI_SECRET = 'jBwaraViHImQRMxeHWFMxeFxs36B38zY'

# --- API 客户端初始化 ---
try:
    cloudconvert.configure(api_key=CLOUDCONVERT_API_KEY)
    convertapi.api_credentials = CONVERTAPI_SECRET
except Exception as e:
    logger.error(f"初始化 API 客户端时发生错误: {e}")


# ==============================================================================
# 2. 各 API 的转换逻辑函数 (Conversion Logic)
# ==============================================================================

# --- Adobe ---
def convert_with_adobe(input_path: str, output_path: str) -> bool:
    """使用 Adobe PDF Services SDK 将 PDF 转换为 PPTX。"""
    try:
        if not ADOBE_CLIENT_ID or not ADOBE_CLIENT_SECRET:
            logger.error("Adobe API 凭证未配置。")
            return False

        credentials = ServicePrincipalCredentials(client_id=ADOBE_CLIENT_ID, client_secret=ADOBE_CLIENT_SECRET)
        pdf_services = PDFServices(credentials=credentials)
        
        with open(input_path, 'rb') as file:
            input_stream = file.read()
        input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.PDF)

        export_params = ExportPDFParams(target_format=ExportPDFTargetFormat.PPTX, ocr_lang=ExportOCRLocale.EN_US)
        export_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_params)
        
        location = pdf_services.submit(export_job)
        response = pdf_services.get_job_result(location, ExportPDFResult)
        
        result_asset: CloudAsset = response.get_result().get_asset()
        stream_asset: StreamAsset = pdf_services.get_content(result_asset)

        with open(output_path, "wb") as file:
            file.write(stream_asset.get_input_stream())
        
        logger.info(f"Adobe 转换成功: {input_path} -> {output_path}")
        return True
    except (ServiceApiException, ServiceUsageException, SdkException) as e:
        logger.exception(f'Adobe SDK 执行时遇到异常: {e}')
        return False

# --- CloudConvert ---
def convert_with_cloudconvert(input_path: str, output_path: str) -> bool:
    """使用 CloudConvert API 将 PDF 转换为 PPTX。"""
    try:
        job = cloudconvert.Job.create(payload={
            "tasks": {
                "import-file": {"operation": "import/upload"},
                "convert-file": {
                    "operation": "convert",
                    "input": "import-file",
                    "input_format": "pdf",
                    "output_format": "pptx",
                },
                "export-file": {"operation": "export/url", "input": "convert-file"}
            }
        })

        upload_task = job['tasks'][0]
        cloudconvert.Task.upload(file_name=input_path, task=upload_task)
        
        job = cloudconvert.Job.wait(id=job['id'])
        
        export_task = next(
            t for t in job["tasks"]
            if t["operation"] == "export/url" and t["status"] == "finished"
        )
        
        if not export_task or not export_task.get("result", {}).get("files"):
            raise RuntimeError("CloudConvert 导出未返回文件")

        file_url = export_task["result"]["files"][0]["url"]
        with requests.get(file_url, stream=True, timeout=600) as r:
            r.raise_for_status()
            with open(output_path, "wb") as f:
                for chunk in r.iter_content(chunk_size=8192):
                    if chunk:
                        f.write(chunk)
        
        logger.info(f"CloudConvert 转换成功: {input_path} -> {output_path}")
        return True
    except Exception as e:
        logger.exception(f"CloudConvert 转换失败: {e}")
        return False

# --- ConvertAPI ---
def convert_with_convertapi(input_path: str, output_path: str) -> bool:
    """使用 ConvertAPI 将 PDF 转换为 PPTX。"""
    try:
        if not convertapi.api_credentials:
            logger.error("ConvertAPI Secret 未配置。")
            return False

        result = convertapi.convert('pptx', {'File': input_path}, from_format='pdf')
        saved_files = result.save_files(os.path.dirname(output_path))
        
        if not saved_files:
            logger.error("ConvertAPI 未返回任何输出文件。")
            return False
        
        os.replace(saved_files[0], output_path)
        
        logger.info(f"ConvertAPI 转换成功: {input_path} -> {output_path}")
        return True
    except Exception as e:
        logger.exception(f'调用 ConvertAPI 时出现异常: {e}')
        return False

# --- LibreOffice ---
def _get_soffice_path() -> str:
    """获取 soffice 可执行文件路径。"""
    return os.environ.get("SOFFICE_PATH", "soffice")

def convert_with_libreoffice(input_path: str, output_path: str) -> bool:
    """使用 LibreOffice 无头模式将 PDF 转为 PPTX。"""
    soffice = _get_soffice_path()
    temp_outdir = tempfile.mkdtemp(prefix="lo_out_")
    
    cmd = [
        soffice, "--headless", "--nologo", "--nofirststartwizard",
        "--infilter=impress_pdf_import",
        "--convert-to", "pptx",
        "--outdir", temp_outdir,
        input_path
    ]
    
    try:
        completed = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=300
        )
        if completed.returncode != 0:
            logger.error(f"[LibreOffice] 转换失败，返回码: {completed.returncode}")
            logger.error(f"[LibreOffice] STDOUT: {completed.stdout.strip()}")
            logger.error(f"[LibreOffice] STDERR: {completed.stderr.strip()}")
            return False

        base_name = os.path.splitext(os.path.basename(input_path))[0]
        produced_path = os.path.join(temp_outdir, base_name + ".pptx")

        if not os.path.exists(produced_path) or os.path.getsize(produced_path) == 0:
            logger.error("[LibreOffice] 转换未生成文件或生成了空文件。")
            return False

        shutil.move(produced_path, output_path)
        logger.info(f"LibreOffice 转换成功: {input_path} -> {output_path}")
        return True
    except subprocess.TimeoutExpired:
        logger.exception("[LibreOffice] 转换超时。")
        return False
    except FileNotFoundError:
        logger.exception(f"[LibreOffice] 错误: '{soffice}' 命令未找到。请确保 LibreOffice 已安装并已将其路径添加到系统 PATH, 或设置 SOFFICE_PATH 环境变量。")
        return False
    except Exception as e:
        logger.exception(f"[LibreOffice] 发生未知错误: {e}")
        return False
    finally:
        shutil.rmtree(temp_outdir, ignore_errors=True)

# ==============================================================================
# 3. 通用请求处理逻辑 (Generic Request Handler)
# ==============================================================================

def process_conversion_request(conversion_function):
    """
    一个通用的函数，用于处理所有转换请求。
    它负责文件验证、保存、调用指定的转换函数、清理文件并返回结果。
    """
    if 'file' not in request.files:
        return jsonify({'error': '请求中未找到文件部分'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400
    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': '无效的文件类型，请上传 PDF 文件。'}), 400

    unique_id = uuid.uuid4().hex
    input_filename = f"{unique_id}.pdf"
    output_filename = f"{unique_id}.pptx"
    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    try:
        file.save(input_path)
        
        # 调用具体的转换函数 (例如 convert_with_adobe)
        success = conversion_function(input_path, output_path)
        
        if not success:
            return jsonify({'error': f'使用 {conversion_function.__name__} 转换失败，请检查服务器日志。'}), 500

        # 准备文件下载
        download_filename = f"{os.path.splitext(file.filename)[0]}.pptx"
        return_data = io.BytesIO()
        with open(output_path, 'rb') as fo:
            return_data.write(fo.read())
        return_data.seek(0)
        
        return send_file(
            return_data,
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
            download_name=download_filename,
            as_attachment=True
        )

    except Exception as e:
        logger.exception(f"处理请求时发生未知错误: {e}")
        return jsonify({'error': '服务器内部错误，请稍后重试。'}), 500
    finally:
        # 无论成功失败，都清理临时文件
        for path in [input_path, output_path]:
            if os.path.exists(path):
                try:
                    os.remove(path)
                except Exception as e:
                    logger.warning(f"清理文件 {path} 失败: {e}")


# ==============================================================================
# 4. Flask 路由 (Routes)
# ==============================================================================

@app.route('/')
def index():
    return "PDF to PPTX Conversion Service is running. Use POST /convert/<api_name> to convert a file."

@app.route('/convert/adobe', methods=['POST'])
def handle_adobe_conversion():
    return process_conversion_request(convert_with_adobe)

@app.route('/convert/cloudconvert', methods=['POST'])
def handle_cloudconvert_conversion():
    return process_conversion_request(convert_with_cloudconvert)

@app.route('/convert/convertapi', methods=['POST'])
def handle_convertapi_conversion():
    return process_conversion_request(convert_with_convertapi)

@app.route('/convert/libreoffice', methods=['POST'])
def handle_libreoffice_conversion():
    return process_conversion_request(convert_with_libreoffice)

# ==============================================================================
# 5. 启动服务器 (Server Start)
# ==============================================================================

if __name__ == '__main__':
    # 在生产环境中，建议使用 Gunicorn 或 uWSGI 等 WSGI 服务器来运行。
    # 例如: gunicorn --workers 4 --bind 0.0.0.0:5000 app:app
    app.run(host='0.0.0.0', port=5000, debug=True)