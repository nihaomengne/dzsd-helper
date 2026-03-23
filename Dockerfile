FROM python:3.11-slim

# 设置环境变量
ENV PYTHONDONTWRITEBYTECODE=1
ENV PYTHONUNBUFFERED=1
ENV FLASK_ENV=production

# 设置工作目录
WORKDIR /app

# 安装系统依赖
RUN apt-get update && apt-get install -y \
    gcc \
    && rm -rf /var/lib/apt/lists/*

# 拷贝requirements.txt并安装Python依赖
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
RUN pip install --no-cache-dir gunicorn

# 拷贝应用文件
COPY . /app

# 创建数据目录
RUN mkdir -p /app/data

# 开放端口
EXPOSE 1012

# 使用gunicorn启动应用
CMD ["gunicorn", "--bind", "0.0.0.0:1012", "--workers", "2", "--timeout", "120", "app:app"]
