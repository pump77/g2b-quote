# 나라장터 견적서 사이트 (Render·Oracle Cloud 등 어디서나 같은 방법으로 실행)
FROM python:3.12-slim

# PDF를 만들 크롬(크로미움)과 한글 글꼴(나눔고딕)
RUN apt-get update \
 && apt-get install -y --no-install-recommends chromium fonts-nanum fontconfig \
 && fc-cache -f \
 && rm -rf /var/lib/apt/lists/*

ENV CHROME_PATH=/usr/bin/chromium \
    PYTHONUNBUFFERED=1 \
    PORT=10000

WORKDIR /app
COPY requirements.txt .
RUN pip install --no-cache-dir -r requirements.txt
COPY g2b_lookup.py quote_excel.py quote_page.html server.py ./

EXPOSE 10000
CMD ["python", "server.py"]
