# 用 Python 3.11 作为基础系统
FROM python:3.11-slim

# 安装 Chromium 浏览器（让 kaleido 能导出图片）+ 中文字体（让图表中文不乱码）
RUN apt-get update && apt-get install -y \
    chromium \
    fonts-noto-cjk \
    && rm -rf /var/lib/apt/lists/*

# 设置工作目录
WORKDIR /app

# 先复制依赖文件并安装
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt

# 复制所有代码文件
COPY . .

# 暴露端口
EXPOSE 8501

# 启动 Streamlit
CMD ["streamlit", "run", "app.py", "--server.port=8501", "--server.address=0.0.0.0"]
