# -*- coding: utf-8 -*-
"""Oracle Cloud(서울) 서버를 만들 때 붙여 넣을 '초기화 스크립트'를 만든다.
프로그램 파일을 스크립트 안에 넣어 두므로 GitHub 없이도 서버가 스스로 설치·실행된다.
사용: python make_oracle_script.py  ->  oracle_setup.sh 생성"""
import base64
import io
import os
import tarfile

HERE = os.path.dirname(os.path.abspath(__file__))
FILES = ["Dockerfile", "requirements.txt", "g2b_lookup.py", "quote_excel.py", "quote_page.html", "server.py"]

buf = io.BytesIO()
with tarfile.open(fileobj=buf, mode="w:gz") as tar:
    for name in FILES:
        tar.add(os.path.join(HERE, name), arcname=name)
payload = base64.b64encode(buf.getvalue()).decode()

script = """#!/bin/bash
# 나라장터 견적서 사이트 자동 설치 (Oracle Cloud Ubuntu 서버용)
# 서버가 처음 켜질 때 한 번 실행되어, 약 5~10분 뒤 http://<서버 공인 IP>/ 로 접속할 수 있다.
set -e
exec > /var/log/g2b-quote-setup.log 2>&1

# 1) 도커 설치
export DEBIAN_FRONTEND=noninteractive
apt-get update
apt-get install -y docker.io
systemctl enable --now docker

# 2) 프로그램 파일 풀기
mkdir -p /opt/g2b-quote
cd /opt/g2b-quote
echo '__PAYLOAD__' | base64 -d | tar xz

# 3) 만들고 실행 (서버를 다시 켜도 자동으로 실행됨)
docker build -t g2b-quote .
docker rm -f g2b-quote 2>/dev/null || true
docker run -d --name g2b-quote --restart always -p 80:10000 g2b-quote

# 4) 오라클 우분투는 기본으로 80번 문을 막아 두므로 열어 준다
iptables -I INPUT 6 -m state --state NEW -p tcp --dport 80 -j ACCEPT || iptables -I INPUT -p tcp --dport 80 -j ACCEPT
netfilter-persistent save || true
echo "설치 완료"
""".replace("__PAYLOAD__", payload)

out = os.path.join(HERE, "oracle_setup.sh")
with open(out, "w", encoding="utf-8", newline="\n") as f:
    f.write(script)
print(out, len(script.encode()), "bytes")
