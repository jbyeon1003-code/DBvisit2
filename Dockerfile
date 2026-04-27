# 1. Python 이미지 사용
FROM python:3.9-slim

# 2. 시스템 의존성 설치
RUN apt-get update && apt-get install -y \
    wget \
    curl \
    && rm -rf /var/lib/apt/lists/*

# 3. 작업 디렉토리 설정
WORKDIR /app

# 4. 소스 코드 복사 및 라이브러리 설치
COPY . .
RUN pip install --no-cache-dir -r requirements.txt

# 5. Playwright 전용 브라우저 설치
RUN playwright install chromium
RUN playwright install-deps chromium

# 6. 서버 실행
CMD gunicorn --bind 0.0.0.0:$PORT app:app
