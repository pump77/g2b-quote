# -*- coding: utf-8 -*-
"""나라장터 종합쇼핑몰에서 물품식별번호로 품목 정보를 조회한다."""
import html
import json
import os
import re
import urllib.request

BASE = "https://shop.g2b.go.kr"
SEARCH_URL = BASE + "/gm/gms/gmsd/newShopUntySrchApi.do"
UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/130.0 Safari/537.36"

_EMPTY_KEYS = (
    "tabDiv apmlNo itemCfnm reSelectValue researchKeyword andKeyword orKeyword notKeyword lCate mCate "
    "etpsNm ctrtClass selectCert hdofcLctn hdofcLctnOrg hdofcSgnguCd hdofcLctnNm allItemCert itemCert "
    "dtlsPrnm shopCtrtTyCd dtlsPrnmNo ctent asctEtpsUntyGrp entFormSeCd spplRgn plor asnt "
    "ftalPrchsTrgtCert devyCndt startPrce endPrce asctCommItmltYn dtlsMrkeDtrctCd ndfmRgstTrgtYn "
    "pdctAtrbCd pdctAtrbCdDtl etpsFctyInfo itmltSeCd spplEtpsHdofcLctn dscplnEvntDmndNo exbtNo usrTyCd "
    "shopGdsCnt srchAtrbBool srchAtrbCode dtlsPrnmNoOff dtlsPrnmNoSrchYn nmmbrNtslYn"
).split()


def _body(number):
    vo = {k: "" for k in _EMPTY_KEYS}
    vo.update({
        "target": "계300001,계300002,계309999",
        "selectValue": "itemIdnfNo",
        "searchKeyword": number,
        "prchsMthdSeCd": "본010001",
        "sortCndt": "SRCH_UPRC_ASC",
        "recordCountPerPage": "10",
        "currentPage": 1,
        "srchSeCd": "검030006",
        "rdoIndex": 1,
        "dgtlSrvcMallYn": "N",
        "dscplnEvntYn": "N",
    })
    return json.dumps({"searchVO": vo}, ensure_ascii=False).encode("utf-8")


def _post(number, timeout=20, tries=3):
    """나라장터가 가끔 연결을 끊으므로 몇 번 다시 시도한다."""
    import time
    for i in range(tries):
        try:
            return _post_once(number, timeout)
        except Exception:
            if i == tries - 1:
                raise
            time.sleep(1 + i)


def _post_once(number, timeout=20):
    req = urllib.request.Request(SEARCH_URL, data=_body(number), method="POST", headers={
        "Content-Type": "application/json; charset=UTF-8",
        "Accept": "application/json",
        "User-Agent": UA,
        "Origin": BASE,
        "Referer": BASE + "/",
    })
    with urllib.request.urlopen(req, timeout=timeout) as res:
        return json.loads(res.read().decode("utf-8"))


def lookup(number):
    """식별번호 하나를 조회해 dict를 돌려준다. 없으면 None."""
    number = re.sub(r"\D", "", str(number))
    if not number:
        return None
    data = _post(number)
    rows = [r for r in (data.get("rsltList") or []) if str(r.get("itemIdnfNo")) == number]
    if not rows:
        return None
    r = rows[0]
    full = html.unescape(r.get("itemIdnfNm") or "")
    parts = [p.strip() for p in full.split(",")]
    spec = ""
    m = re.search(r"\d+(?:\.\d+)?(?:\s*[×xX*]\s*\d+(?:\.\d+)?){1,2}", full)
    if m:
        spec = re.sub(r"\s*[xX*×]\s*", "×", m.group(0))
    unit = html.unescape(r.get("ctrtUntVal") or "")
    company = html.unescape(r.get("rprsCtentUntyGrpNm") or r.get("mnftrEtpsNm") or "")
    company = re.sub(r"주식회사|유한회사|\(주\)|㈜|\([^)]*\)", "", company).strip()
    return {
        "company": company or (parts[1] if len(parts) > 1 else ""),
        "number": number,
        "name": (parts[0] if parts and parts[0] else html.unescape(r.get("itemCfnm") or "")),
        "maker": parts[1] if len(parts) > 1 else "",
        "model": parts[2] if len(parts) > 2 else "",
        "spec": spec,
        "full_name": full,
        "unit": "EA" if unit in ("개", "") else unit,
        "price": int(float(r.get("ctrtUprc") or r.get("srchUprc") or 0)),
        "image_url": (BASE + r["imgSrc"]) if r.get("imgSrc") else "",
    }


def download_image(item, folder):
    """품목 사진을 folder에 받아 경로를 돌려준다. 실패하면 None."""
    if not item.get("image_url"):
        return None
    os.makedirs(folder, exist_ok=True)
    ext = os.path.splitext(item["image_url"].split("?")[0])[1].lower() or ".jpg"
    path = os.path.join(folder, item["number"] + ext)
    if os.path.exists(path) and os.path.getsize(path) > 0:
        return path
    for _ in range(2):
        try:
            req = urllib.request.Request(item["image_url"], headers={"User-Agent": UA, "Referer": BASE + "/"})
            with urllib.request.urlopen(req, timeout=20) as res:
                data = res.read()
            with open(path, "wb") as f:
                f.write(data)
            return path
        except Exception:
            pass
    return None


if __name__ == "__main__":
    import sys
    for n in sys.argv[1:]:
        print(n, lookup(n))
