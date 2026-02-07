# Python 3.9 슬림 버전을 베이스 이미지로 사용
FROM python:3.9-slim

# 작업 디렉토리 설정
WORKDIR /app

# 시스템 의존성 설치 (필요한 경우)
# RUN apt-get update && apt-get install -y ...

# 필수 패키지 설치를 위해 requirements.txt 복사
COPY requirements.txt .

# pip 업그레이드 및 패키지 설치
RUN pip install --no-cache-dir --upgrade pip && \
    pip install --no-cache-dir -r requirements.txt

# 소스 코드 전체 복사
COPY . .

# 포트 8080 노출 (Google Cloud Run 기본 포트)
EXPOSE 8080

# Streamlit 실행 (포트 8080으로 설정)
CMD streamlit run app.py --server.port=8080 --server.address=0.0.0.0
