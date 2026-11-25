# -*- coding: utf-8 -*-
"""나라장터 견적서 작성 (웹사이트용).
누구나 인터넷 주소로 들어와 식별번호만 넣으면 견적서(엑셀·PDF)를 받을 수 있다.
공급자 정보는 서버에 저장하지 않고, 각자의 브라우저에만 기억된다.
실행:  python server.py   (PORT 환경변수가 있으면 그 번호, 없으면 8000)"""
import base64
import json
import mimetypes
import os
import secrets
import shutil
import subprocess
import sys
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, ThreadingHTTPServer
from urllib.parse import parse_qs, quote, urlparse

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)
import g2b_lookup  # noqa: E402
import quote_excel  # noqa: E402

PORT = int(os.environ.get("PORT") or "8000")
IMG_DIR = os.environ.get("IMG_DIR") or os.path.join(tempfile.gettempdir(), "g2b_images")
PAGE = os.path.join(HERE, "quote_page.html")
MAX_UPLOAD = 10 * 1024 * 1024   # 불러오기 파일은 10MB까지
MAX_BODY = 2 * 1024 * 1024      # 견적 내용(JSON)은 2MB까지
MAX_ITEMS = 300

PRINT_JOBS = {}
PDF_LOCK = threading.Semaphore(1)  # 무료 서버는 메모리가 작으므로 PDF는 한 번에 하나씩 만든다
PDF_MARK = b"\n%G2BQUOTE:"         # PDF 끝에 견적 내용을 붙여 두어, 나중에 다시 불러와 고칠 수 있게 한다


def find_browser():
    for exe in (os.environ.get("CHROME_PATH"), "chromium", "chromium-browser", "google-chrome", "google-chrome-stable"):
        if exe and (os.path.exists(exe) or shutil.which(exe)):
            return shutil.which(exe) or exe
    return None


def make_pdf(data):
    """화면의 인쇄 모양 그대로 PDF를 만들고, 끝에 견적 내용을 붙여 돌려준다."""
    exe = find_browser()
    if not exe:
        raise RuntimeError("지금은 PDF를 만들 수 없습니다. '인쇄' 버튼을 눌러 'PDF로 저장'을 골라 주세요.")
    token = secrets.token_hex(12)
    tmp = tempfile.mkdtemp(prefix="g2bquote_")
    out = os.path.join(tmp, "quote.pdf")
    with PDF_LOCK:
        PRINT_JOBS[token] = data
        try:
            subprocess.run(
                [exe, "--headless=new", "--no-sandbox", "--disable-gpu", "--disable-dev-shm-usage",
                 "--no-first-run", "--no-pdf-header-footer", "--hide-scrollbars", "--mute-audio",
                 "--disable-extensions", "--disable-background-networking",
                 "--user-data-dir=" + os.path.join(tmp, "profile"), "--virtual-time-budget=10000",
                 "--print-to-pdf=" + out, "http://127.0.0.1:%d/?print=%s" % (PORT, token)],
                timeout=90, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
            with open(out, "rb") as f:
                pdf = f.read()
        except Exception:
            raise RuntimeError("PDF를 만들지 못했습니다. 잠시 뒤 다시 눌러 주세요.")
        finally:
            PRINT_JOBS.pop(token, None)
            shutil.rmtree(tmp, ignore_errors=True)
    payload = json.dumps({"info": data["info"], "items": data["items"]}, ensure_ascii=False).encode("utf-8")
    return pdf + PDF_MARK + base64.b64encode(payload) + b"\n"


def read_pdf(raw):
    i = raw.rfind(PDF_MARK)
    if i < 0:
        raise ValueError("이 사이트에서 'PDF로 저장'한 파일만 불러올 수 있습니다.")
    data = json.loads(base64.b64decode(raw[i + len(PDF_MARK):].strip()).decode("utf-8"))
    return {"info": data.get("info") or {}, "items": data.get("items") or []}


def _int(v):
    try:
        return int(float(str(v or 0).replace(",", "")))
    except ValueError:
        return 0


def clean_items(raw_items):
    items = []
    for it in (raw_items or [])[:MAX_ITEMS]:
        number = "".join(ch for ch in str(it.get("number") or "") if ch.isdigit())[:20]
        items.append({
            "name": str(it.get("name") or "")[:200], "model": str(it.get("model") or "")[:200],
            "spec": str(it.get("spec") or "")[:200], "qty": _int(it.get("qty")),
            "unit": str(it.get("unit") or "EA")[:20], "price": _int(it.get("price")), "number": number,
            "note": str(it.get("note") or "")[:200],
        })
    return items


def clean_info(raw):
    return {str(k)[:30]: str(v or "")[:300] for k, v in (raw or {}).items()}


def find_image(number):
    for ext in (".jpg", ".jpeg", ".png", ".gif"):
        p = os.path.join(IMG_DIR, number + ext)
        if os.path.exists(p) and os.path.getsize(p) > 0:
            return p
    return None


def ensure_image(number):
    """사진이 서버에 없으면 나라장터에서 다시 받아 둔다 (무료 서버는 잠들면 사진이 지워지므로)."""
    if not number or find_image(number):
        return find_image(number) if number else None
    try:
        found = g2b_lookup.lookup(number)
        if found:
            g2b_lookup.download_image(found, IMG_DIR)
    except Exception:
        pass
    return find_image(number)


class Handler(BaseHTTPRequestHandler):
    server_version = "g2b-quote"

    def log_message(self, fmt, *args):
        pass

    def send_bytes(self, data, ctype, status=200, extra=None, cache="no-store"):
        self.send_response(status)
        self.send_header("Content-Type", ctype)
        self.send_header("Content-Length", str(len(data)))
        self.send_header("Cache-Control", cache)
        self.send_header("X-Content-Type-Options", "nosniff")
        for k, v in (extra or {}).items():
            self.send_header(k, v)
        self.end_headers()
        if self.command != "HEAD":
            self.wfile.write(data)

    def send_json(self, obj, status=200):
        self.send_bytes(json.dumps(obj, ensure_ascii=False).encode("utf-8"), "application/json; charset=utf-8", status)

    def read_body(self, limit):
        n = int(self.headers.get("Content-Length") or 0)
        if n > limit:
            raise ValueError("파일이 너무 큽니다.")
        return self.rfile.read(n)

    def do_HEAD(self):
        self.do_GET()

    def do_GET(self):
        url = urlparse(self.path)
        if url.path == "/healthz":
            return self.send_json({"ok": True})
        if url.path == "/api/printdata":
            # PDF를 만드는 서버 안의 브라우저만 쓰는 주소
            if self.client_address[0] not in ("127.0.0.1", "::1"):
                return self.send_bytes(b"not found", "text/plain", 404)
            job = PRINT_JOBS.get((parse_qs(url.query).get("id") or [""])[0])
            if not job:
                return self.send_json({"info": {}, "items": []})
            items = [dict(it, has_image=bool(it["number"] and find_image(it["number"]))) for it in job["items"]]
            return self.send_json({"info": job["info"], "items": items})
        if url.path in ("/", "/index.html"):
            with open(PAGE, "rb") as f:
                return self.send_bytes(f.read(), "text/html; charset=utf-8")
        if url.path == "/api/lookup":
            number = "".join(ch for ch in (parse_qs(url.query).get("no") or [""])[0] if ch.isdigit())[:20]
            if not number:
                return self.send_json({"ok": False, "error": "식별번호를 입력하세요."})
            try:
                item = g2b_lookup.lookup(number)
            except Exception:
                return self.send_json({"ok": False, "error": "나라장터에 연결하지 못했습니다. 잠시 뒤 다시 넣어 주세요."})
            if not item:
                return self.send_json({"ok": False, "error": "나라장터 쇼핑몰에 없는 식별번호입니다."})
            item["has_image"] = bool(g2b_lookup.download_image(item, IMG_DIR))
            return self.send_json({"ok": True, "item": item})
        if url.path.startswith("/img/"):
            number = "".join(ch for ch in url.path[5:] if ch.isdigit())[:20]
            p = ensure_image(number) if number else None
            if not p:
                return self.send_bytes(b"", "text/plain", 404)
            with open(p, "rb") as f:
                return self.send_bytes(f.read(), mimetypes.guess_type(p)[0] or "image/jpeg", cache="public, max-age=86400")
        self.send_bytes(b"not found", "text/plain", 404)

    def do_POST(self):
        path = urlparse(self.path).path
        if path == "/api/import":
            try:
                raw = self.read_body(MAX_UPLOAD)
                result = read_pdf(raw) if raw[:4] == b"%PDF" else quote_excel.read_quote(raw)
                result["items"] = result["items"][:MAX_ITEMS]
            except ValueError as e:
                return self.send_json({"ok": False, "error": str(e)})
            except Exception:
                return self.send_json({"ok": False, "error": "파일을 읽지 못했습니다. 이 사이트에서 저장한 엑셀(.xlsx)이나 PDF만 됩니다."})
            for it in result["items"]:
                it["has_image"] = bool(ensure_image(it.get("number")))
            result["ok"] = True
            return self.send_json(result)
        if path not in ("/api/pdf", "/api/excel"):
            return self.send_bytes(b"not found", "text/plain", 404)
        try:
            body = json.loads(self.read_body(MAX_BODY).decode("utf-8"))
            info, items = clean_info(body.get("info")), clean_items(body.get("items"))
            for it in items:
                ensure_image(it["number"])
        except Exception:
            return self.send_json({"ok": False, "error": "견적 내용을 읽지 못했습니다."}, 400)
        if path == "/api/pdf":
            try:
                data = make_pdf({"info": info, "items": items})
            except RuntimeError as e:
                return self.send_json({"ok": False, "error": str(e)}, 500)
            return self.send_bytes(data, "application/pdf")
        try:
            for it in items:
                it["image"] = find_image(it["number"]) if it["number"] else None
            data = quote_excel.build_quote(items, info)
        except Exception:
            return self.send_json({"ok": False, "error": "엑셀 파일을 만들지 못했습니다."}, 500)
        to = "".join(ch for ch in info.get("to", "") if ch not in '\\/:*?"<>|').strip()
        name = "견적서_%s.xlsx" % to if to else "견적서.xlsx"
        self.send_bytes(data, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        extra={"Content-Disposition": "attachment; filename*=UTF-8''" + quote(name)})


class Server(ThreadingHTTPServer):
    daemon_threads = True
    allow_reuse_address = True


def warm_up():
    """서버가 켜지자마자 PDF용 브라우저를 한 번 미리 띄워 둔다. 첫 'PDF로 저장'이 훨씬 빨라진다."""
    try:
        make_pdf({"info": {}, "items": []})
    except Exception:
        pass


def main():
    os.makedirs(IMG_DIR, exist_ok=True)
    server = Server(("0.0.0.0", PORT), Handler)
    threading.Thread(target=warm_up, daemon=True).start()
    print("나라장터 견적서 사이트 실행 중: http://127.0.0.1:%d/" % PORT, flush=True)
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        pass


if __name__ == "__main__":
    main()
