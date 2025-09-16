import logging
import os
import uuid
import io
import shutil
import tempfile
import subprocess
from flask import Flask, request, jsonify, send_file
from flask_cors import CORS

# ================= 配置日志 =================
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger("pdf2pptx")

# ================= 初始化 Flask =================
app = Flask(__name__)
CORS(app)  # 如需可收紧白名单

# ================= 目录配置 =================
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'outputs'
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER

# ================= 工具函数 =================
def _get_soffice_path() -> str:
    """
    获取 soffice 可执行文件路径。
    优先使用环境变量 SOFFICE_PATH，否则依赖 PATH 中的 'soffice'。
    """
    return os.environ.get("SOFFICE_PATH", "soffice")

def convert_pdf_to_pptx(input_pdf_path: str, output_pptx_path: str, timeout_sec: int = 300) -> bool:
    """
    使用 LibreOffice 无头模式将 PDF 转为 PPTX。
    - input_pdf_path: 输入 PDF 完整路径
    - output_pptx_path: 目标 PPTX 完整路径
    - timeout_sec: 超时时间（秒）
    返回 True/False 表示成功与否。
    """
    soffice = _get_soffice_path()

    # LibreOffice 会把输出文件写到 --outdir 下，文件名与输入同名改后缀
    # 为避免并发/权限问题，这里用独立临时输出目录，成功后再移动到 output_pptx_path
    temp_outdir = tempfile.mkdtemp(prefix="lo_out_")

    cmd = [
        soffice,
        "--headless",
        "--nologo",
        "--nofirststartwizard",
        "--infilter=impress_pdf_import",  # 关键：以 Impress 方式导入 PDF
        "--convert-to", "pptx",
        "--outdir", temp_outdir,
        input_pdf_path
    ]

    logger.info("Running soffice: %s", " ".join(f'"{c}"' if " " in c else c for c in cmd))
    try:
        completed = subprocess.run(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            timeout=timeout_sec
        )
        if completed.returncode != 0:
            logger.error("[LibreOffice] returncode=%s", completed.returncode)
            logger.error("[LibreOffice] STDOUT: %s", completed.stdout.strip())
            logger.error("[LibreOffice] STDERR: %s", completed.stderr.strip())
            return False

        # 生成的文件名与源文件同名但后缀为 .pptx
        base_name = os.path.splitext(os.path.basename(input_pdf_path))[0]
        produced_path = os.path.join(temp_outdir, base_name + ".pptx")

        if not os.path.exists(produced_path) or os.path.getsize(produced_path) == 0:
            logger.error("[LibreOffice] Conversion produced no output or empty file")
            logger.error("[LibreOffice] STDOUT: %s", completed.stdout.strip())
            logger.error("[LibreOffice] STDERR: %s", completed.stderr.strip())
            return False

        # 使用 shutil.move() 处理跨设备文件系统
        shutil.move(produced_path, output_pptx_path)
        logger.info("Converted OK: %s -> %s", input_pdf_path, output_pptx_path)
        return True

    except subprocess.TimeoutExpired:
        logger.exception("[LibreOffice] Conversion timed out")
        return False
    except Exception as e:
        logger.exception("[LibreOffice] Unexpected error: %s", e)
        return False
    finally:
        # 清理临时输出目录
        try:
            shutil.rmtree(temp_outdir, ignore_errors=True)
        except Exception:
            pass

# ================= 路由 =================
@app.route('/convert', methods=['POST'])
def handle_conversion():
    """
    接收一个名为 'file' 的 PDF 文件，并直接返回转换后的 PPTX 文件。
    """
    if 'file' not in request.files:
        return jsonify({'error': '请求中未找到文件部分'}), 400

    file = request.files['file']
    if file.filename == '':
        return jsonify({'error': '未选择文件'}), 400

    if not file.filename.lower().endswith('.pdf'):
        return jsonify({'error': '无效的文件类型，请上传 PDF 文件。'}), 400

    # 生成随机文件名，避免并发冲突
    unique_id = uuid.uuid4().hex
    input_filename = f"{unique_id}.pdf"
    output_filename = f"{unique_id}.pptx"

    input_path = os.path.join(app.config['UPLOAD_FOLDER'], input_filename)
    output_path = os.path.join(app.config['OUTPUT_FOLDER'], output_filename)

    try:
        # 保存上传的 PDF
        file.save(input_path)

        # 调用 LibreOffice 转换
        ok = convert_pdf_to_pptx(input_path, output_path)

        # 清理上传的源 PDF
        try:
            os.remove(input_path)
        except Exception:
            pass

        if not ok:
            # 失败时清理可能残留的输出
            try:
                if os.path.exists(output_path):
                    os.remove(output_path)
            except Exception:
                pass
            return jsonify({'error': '文件转换失败，请检查服务器日志。'}), 500

        # 成功：将 PPTX 加载进内存并作为附件返回
        download_filename = f"{os.path.splitext(file.filename)[0]}.pptx"

        return_data = io.BytesIO()
        with open(output_path, 'rb') as fo:
            return_data.write(fo.read())
        return_data.seek(0)

        # 发送后清理服务器上的 PPTX
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
        logger.exception("处理转换请求时发生异常: %s", e)
        # 出错时尽力清理
        try:
            if os.path.exists(input_path):
                os.remove(input_path)
        except Exception:
            pass
        try:
            if os.path.exists(output_path):
                os.remove(output_path)
        except Exception:
            pass
        return jsonify({'error': '服务器内部错误，请稍后重试。'}), 500

# ================= 启动 =================
if __name__ == '__main__':
    # 生产建议用 gunicorn/uwsgi + nginx；这里用于本地/容器内测试
    app.run(host='0.0.0.0', port=5000, debug=True)