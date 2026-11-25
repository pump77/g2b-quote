# 나라장터 견적서 작성

나라장터 물품식별번호만 넣으면 품목·사진·단가가 자동으로 채워지고, 견적서를 엑셀·PDF로 받을 수 있는 웹사이트입니다.

- 실행: `python server.py` (또는 `docker build -t g2b-quote . && docker run -p 10000:10000 g2b-quote`)
- Render: `render.yaml` (무료, 싱가포르)
- Oracle Cloud(서울): 서버를 만들 때 `oracle_setup.sh` 를 초기화 스크립트로 넣기. 코드를 고쳤다면 `python make_oracle_script.py` 로 다시 만들기.
