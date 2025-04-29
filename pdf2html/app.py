from flask import Flask, render_template, request, send_from_directory, redirect, url_for, make_response
import os
import subprocess
import shutil
import requests
from werkzeug.utils import secure_filename
import json
import uuid
import random



app = Flask(__name__)

# PDF转换配置
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'static/outputs'
MARKDOWN_FOLDER = 'static/markdown'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['OUTPUT_FOLDER'] = OUTPUT_FOLDER
app.config['MARKDOWN_FOLDER'] = MARKDOWN_FOLDER

# 爬虫配置
FIRE_CRAWL_BASE = "http://localhost:3002/v1"
API_KEY = "test"  # 替换为真实 API Key
CRAWL_API_KEY = f"fc-{API_KEY}"

# 确保目录存在
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(MARKDOWN_FOLDER, exist_ok=True)



# ======== 工具函数 ========

def convert_pdf_to_html(input_pdf, output_dir, output_html):
    try:
        # 获取宿主机当前工作目录
        host_cwd = os.getcwd()
        # 构造容器内路径
        container_input = os.path.join('/pdf2html', input_pdf).replace('\\', '/')
        container_output_dir = os.path.join('/pdf2html', output_dir).replace('\\', '/')

        # 构建 Docker 命令
        cmd = [
            'docker', 'run', '--rm',
            '-v', f'{host_cwd}:/pdf2html',  # 挂载宿主机目录到容器
            'vinsonwang/pdf2htmlex:latst',
            'pdf2htmlEX',
            '--embed', 'cfijo',
            '--dest-dir', container_output_dir,
            container_input,
            output_html  # 指定输出文件名
        ]

        # 执行命令
        subprocess.run(
            cmd,
            check=True,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE
        )
        return True, None
    except subprocess.CalledProcessError as e:
        return False, e.stderr.decode()

# ======== 路由处理 ========
@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # 处理PDF上传
        if 'pdf_file' in request.files:
            return handle_pdf_upload()

        # 处理爬虫请求
        return handle_crawl_request()

    return render_template('index.html')


def handle_pdf_upload():
    file = request.files['pdf_file']
    if file.filename == '':
        return "未选择文件！", 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    file.save(pdf_path)

    # 生成唯一文件夹名称（UUID）
    unique_folder = str(uuid.uuid4())
    output_dir = os.path.join(app.config['OUTPUT_FOLDER'], unique_folder)
    os.makedirs(output_dir, exist_ok=True)

    # 转换 PDF 并存入唯一文件夹
    output_html = "converted.html"
    success, error = convert_pdf_to_html(pdf_path, output_dir, output_html)
    if not success:
        return f"转换失败: {error}", 500

    # 生成访问 URL
    html_url = url_for('static', filename=f'outputs/{unique_folder}/{output_html}', _external=True)
    return render_template('index.html',
                           pdf_conversion_success=True,
                           html_url=html_url)


def handle_crawl_request():
    url = request.form.get('url')
    action = request.form.get('action')
    headers = create_headers(action)

    # 初始化变量
    markdown_filename = None
    html_filename = None

    try:
        data = build_request_data(request, action)
        response = requests.post(
            f"{FIRE_CRAWL_BASE}/{get_endpoint(action)}",
            headers=headers,
            json=data
        )
        results = response.json()

        # 保存Markdown数据
        if results.get('data') and results['data'].get('markdown'):
            markdown_filename = save_markdown(results['data']['markdown'])

        # 保存rawHtml数据
        if 'rawHtml' in request.form.getlist('formats') and results.get('data') and results['data'].get('rawHtml'):
            html_filename = save_html(results['data']['rawHtml'])

        return render_template('index.html',
                             crawl_results=results,
                             action=action,
                             html_url=url,
                             markdown_filename=markdown_filename,
                             html_filename=html_filename)
    except Exception as e:
        return render_template('index.html', error=str(e))



# ======== 辅助函数 ========
def create_headers(action):
    headers = {
        'Content-Type': 'application/json',
        'Authorization': f'Bearer {CRAWL_API_KEY if action == "crawl" else API_KEY}'
    }
    return headers


def build_request_data(request, action):
    # data = {'url': request.form['url']}
    data = {
        'url': request.form['url'],
        'onlyMainContent': True  # 写死 onlyMainContent=True，排除导航、页眉和页脚
    }

    if action == 'scrape':
        data['formats'] = request.form.getlist('formats')
    elif action == 'map':
        data['search'] = request.form.get('search', '')
    elif action == 'crawl':
        data.update({
            'limit': int(request.form.get('limit', 100)),
            'scrapeOptions': {
                'formats': request.form.getlist('crawl_formats')
            }
        })
    return data


def get_endpoint(action):
    return 'crawl' if action == 'crawl' else action


def save_markdown(markdown_content):
    """保存 Markdown 文件，并添加 4 位随机流水号"""
    serial = random.randint(1000, 9999)
    markdown_filename = f"extracted_content_{serial}.md"
    markdown_path = os.path.join(app.config['MARKDOWN_FOLDER'], markdown_filename)

    with open(markdown_path, 'w', encoding='utf-8') as f:
        f.write(markdown_content)

    return markdown_filename


def save_html(raw_html_content):
    """保存 HTML 文件，并添加 4 位随机流水号"""
    serial = random.randint(1000, 9999)
    html_filename = f"extracted_content_{serial}.html"
    html_path = os.path.join(app.config['UPLOAD_FOLDER'], html_filename)

    with open(html_path, 'w', encoding='utf-8') as f:
        f.write(raw_html_content)

    return html_filename


# ======== 文件下载 ========
@app.route('/download/<folder>/<filename>')
def download(folder, filename):
    """提供下载 PDF 转换后的 HTML 文件"""
    return send_from_directory(os.path.join(app.config['OUTPUT_FOLDER'], folder), filename, as_attachment=True)


@app.route('/download_markdown/<filename>')
def download_markdown(filename):
    """提供 Markdown 文件下载"""
    return send_from_directory(app.config['MARKDOWN_FOLDER'], filename, as_attachment=True)


@app.route('/download_html/<filename>')
def download_html(filename):
    """提供 HTML 文件下载"""
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename, as_attachment=True)


if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8000, debug=True)
