FROM python:3.10

# 컨테이너 안 작업폴더
WORKDIR /app

# requirements 먼저 설치 → 캐시 활용
COPY requirements.txt .

RUN pip install --no-cache-dir -r requirements.txt

# 소스 코드 복사
COPY . .

# 기본 명령을 bash 로 → 가상환경 켜진 feel
CMD ["/bin/bash"]