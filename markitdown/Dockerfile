FROM python:3.13-slim-bullseye

# 设置环境变量：关闭字节码写入、开启实时输出
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# 安装系统依赖（支持 Pillow + 安装 Docker CLI）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg62-turbo-dev \
    docker.io \
 && apt-get clean && rm -rf /var/lib/apt/lists/*

# 设置工作目录
WORKDIR /app

# 拷贝 requirements 并安装依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 拷贝项目文件
COPY . .

# 创建上传和输出目录
RUN mkdir -p uploads outputs

# 显式声明端口
EXPOSE 5009

# 默认运行 Flask 应用
CMD ["python", "app.py"]
