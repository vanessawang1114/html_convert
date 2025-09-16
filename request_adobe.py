import logging
import os
import uuid
import io
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# 从 Adobe SDK 导入必要的模块
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

# 配置日志
logging.basicConfig(level=logging.INFO)

# 初始化 Flask 应用
app = Flask(__name__)
# 启用 CORS，允许来自任何源的请求，这对于本地文件测试是必需的
CORS(app)

# 配置上传和输出文件夹
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
# 确保文件夹存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

def convert_pdf_to_pptx(input_pdf_path, output_pptx_path):
    """
    使用 Adobe PDF Services SDK 将 PDF 转换为 PPTX。
    :param input_pdf_path: 输入的 PDF 文件路径。
    :param output_pptx_path: 输出的 PPTX 文件路径。
    :return: 成功时返回 True，失败时返回 False。
    """
    
    try:
        # 从环境变量中获取 Adobe API 凭证
        # 请确保在运行前设置好 ADOBE_CLIENT_ID 和 ADOBE_CLIENT_SECRET
        client_id = '9e6937018e1542439925726b28509327'
        client_secret = 'p8e-JgpUaqdPWuHMkWDTc5rYsJ9kBex2hBjr'

        if not client_id or not client_secret:
            logging.error("错误: ADOBE_CLIENT_ID 和 ADOBE_CLIENT_SECRET 环境变量未设置。")
            return False

        # 1. 初始化凭证
        credentials = ServicePrincipalCredentials(
                client_id='9e6937018e1542439925726b28509327',
                client_secret='p8e-JgpUaqdPWuHMkWDTc5rYsJ9kBex2hBjr'
            )

        # 2. 创建 PDF Services 实例
        pdf_services = PDFServices(credentials=credentials)

        # 3. 读取输入文件并上传
        with open(input_pdf_path, 'rb') as file:
            input_stream = file.read()
        input_asset = pdf_services.upload(input_stream=input_stream, mime_type=PDFServicesMediaType.PDF)

        # 4. 设置导出参数
        export_pdf_params = ExportPDFParams(target_format=ExportPDFTargetFormat.PPTX, ocr_lang=ExportOCRLocale.EN_US)

        # 5. 创建导出任务
        export_pdf_job = ExportPDFJob(input_asset=input_asset, export_pdf_params=export_pdf_params)

        # 6. 提交任务并获取结果
        location = pdf_services.submit(export_pdf_job)
        pdf_services_response = pdf_services.get_job_result(location, ExportPDFResult)

        # 7. 下载结果
        result_asset: CloudAsset = pdf_services_response.get_result().get_asset()
        stream_asset: StreamAsset = pdf_services.get_content(result_asset)

        # 8. 将结果写入输出文件
        with open(output_pptx_path, "wb") as file:
            file.write(stream_asset.get_input_stream())
        
        logging.info(f"成功将 {input_pdf_path} 转换为 {output_pptx_path}")
        return True

    except (ServiceApiException, ServiceUsageException, SdkException) as e:
        logging.exception(f'执行操作时遇到异常: {e}')
        return False

@app.route('/convert', methods=['POST'])
def handle_conversion():
    """
    处理 PDF 转换请求的 API 端点。
    接收一个名为 'file' 的 PDF 文件，并直接返回转换后的 PPTX 文件。
    """
    if 'file' not in request.files:
        return jsonify({'error': '请求中未找到文件部分'}), 400

    file = request.files['file']

    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400

    if file and file.filename.lower().endswith('.pdf'):
        unique_id = uuid.uuid4().hex
        input_filename = f"{unique_id}.pdf"
        output_filename = f"{unique_id}.pptx"

        input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
        output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

        file.save(input_path)
        success = convert_pdf_to_pptx(input_path, output_path)
        os.remove(input_path) # 清理上传的 PDF

        if success:
            try:
                # 准备一个友好的下载文件名，例如 "original_name.pptx"
                download_filename = f"{os.path.splitext(file.filename)[0]}.pptx"
                
                # 将文件作为附件直接发送给客户端
                return_data = io.BytesIO()
                with open(output_path, 'rb') as fo:
                    return_data.write(fo.read())
                # 指针移到开头
                return_data.seek(0)

                # 发送后清理服务器上的文件
                os.remove(output_path)

                return send_file(
                    return_data,
                    mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation',
                    download_name=download_filename,
                    as_attachment=True
                )
            except Exception as e:
                logging.error(f"发送文件时出错: {e}")
                return jsonify({'error': '文件转换成功但发送时出错。'}), 500
        else:
            return jsonify({'error': '文件转换失败，请检查服务器日志。'}), 500
    else:
        return jsonify({'error': '无效的文件类型，请上传 PDF 文件。'}), 400

if __name__ == '__main__':
    # 启动 Flask 服务，监听所有网络接口的 5000 端口
    app.run(host='0.0.0.0', port=5000, debug=True)