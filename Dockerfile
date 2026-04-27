# 1. Python 이미지 사용
FROM python:3.9-slim

# 2. 필수 패키지 및 크롬 브라우저 설치
RUN apt-get update && apt-get install -y \
    wget \
    gnupg \
    unzip \
    curl \
    ca-certificates \
    fonts-liberation \
    libnss3 \
    libatk-bridge2.0-0 \
    libgtk-3-0 \
    libasound2 \
    xdg-utils \
    --no-install-recommends \
    && curl -fSsL https://dl.google.com/linux/linux_signing_key.pub | gpg --dearmor | tee /usr/share/keyrings/google-chrome.gpg > /dev/null \
    && echo "deb [arch=amd64 signed-by=/usr/share/keyrings/google-chrome.gpg] http://dl.google.com/linux/chrome/deb/ stable main" > /etc/apt/sources.list.d/google-chrome.list \
    && apt-get update && apt-get install -y google-chrome-stable \
    && rm -rf /var/lib/apt/lists/*

# 3. 작업 디렉토리 설정
WORKDIR /app

# 4. 소스 코드 복사 및 라이브러리 설치
COPY . .
RUN pip install --no-cache-dir -r requirements.txt

# 5. 서버 실행 (Gunicorn 사용)
CMD gunicorn --bind 0.0.0.0:$PORT app:app
