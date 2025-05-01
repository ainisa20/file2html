FROM python:3.13-slim

# 设置环境变量：关闭字节码、开启实时输出
ENV PYTHONDONTWRITEBYTECODE=1 \
    PYTHONUNBUFFERED=1

# 安装系统依赖（仅 Pillow 需要）
RUN apt-get update && apt-get install -y --no-install-recommends \
    libjpeg62-turbo-dev \
    && apt-get clean \
    && rm -rf /var/lib/apt/lists/*

# 设置工作目录
WORKDIR /app

# 拷贝项目文件
COPY . .

# 安装 Python 依赖
RUN pip install --no-cache-dir flask requests python-docx pillow -i https://pypi.tuna.tsinghua.edu.cn/simple/

# 创建必要的目录
RUN mkdir -p uploads outputs

# 显式暴露端口
EXPOSE 5009

# 启动应用（请将 app.py 改成你的实际文件名）
CMD ["python", "app.py"]
