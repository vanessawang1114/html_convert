import logging
import os
import uuid
import io

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# New: ConvertAPI client
import convertapi

# -------------------------
# Config & initialization
# -------------------------

logging.basicConfig(level=logging.INFO)

app = Flask(__name__)
CORS(app)

UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# Prefer env var, fallback to the key you provided
CONVERTAPI_SECRET = os.getenv('CONVERTAPI_SECRET', 'jBwaraViHImQRMxeHWFMxeFxs36B38zY')


def convert_pdf_to_pptx(input_pdf_path: str, output_pptx_path: str) -> bool:
    """
    使用 ConvertAPI 将 PDF 转换为 PPTX。
    :param input_pdf_path: 输入 PDF 文件路径
    :param output_pptx_path: 目标 PPTX 文件路径（将被创建/覆盖）
    :return: 成功 True，失败 False
    """
    try:
        if not CONVERTAPI_SECRET:
            logging.error("错误: CONVERTAPI_SECRET 未配置。")
            return False

        # 设置 ConvertAPI 凭证
        convertapi.api_credentials = CONVERTAPI_SECRET

        # 执行转换：从 pdf -> pptx
        # 官方示例：
        # convertapi.convert('pptx', {'File': '/path/to/file.pdf'}, from_format='pdf').save_files('/path/to/dir')
        result = convertapi.convert(
            'pptx',
            {'File': input_pdf_path},
            from_format='pdf'
        )

        # 将转出的文件先保存到输出目录
        saved_files = result.save_files(os.path.dirname(output_pptx_path))
        if not saved_files:
            logging.error("ConvertAPI 未返回任何输出文件。")
            return False

        # ConvertAPI 可能返回临时随机文件名，这里统一改名为我们指定的 output_pptx_path
        tmp_path = saved_files[0]
        # os.replace 支持跨平台原子替换
        os.replace(tmp_path, output_pptx_path)

        logging.info(f"成功将 {input_pdf_path} 转换为 {output_pptx_path}")
        return True

    except Exception as e:
        logging.exception(f'调用 ConvertAPI 转换时出现异常: {e}')
        return False


@app.route('/convert', methods=['POST'])
def handle_conversion():
    """
    接收一个名为 'file' 的 PDF，返回转换后的 PPTX 文件（二进制下载）。
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

        try:
            file.save(input_path)
            success = convert_pdf_to_pptx(input_path, output_path)
        finally:
            # 清理上传的 PDF
            if os.path.exists(input_path):
                try:
                    os.remove(input_path)
                except Exception as e:
                    logging.warning(f"清理上传文件失败: {e}")

        if success:
            try:
                download_filename = f"{os.path.splitext(file.filename)[0]}.pptx"

                return_data = io.BytesIO()
                with open(output_path, 'rb') as fo:
                    return_data.write(fo.read())
                return_data.seek(0)

                # 发送后清理服务器上的文件
                try:
                    os.remove(output_path)
                except Exception as e:
                    logging.warning(f"清理输出文件失败: {e}")

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
    # 启动 Flask 服务
    app.run(host='0.0.0.0', port=5000, debug=True)