# -*- coding: utf-8 -*-
"""견적서 엑셀 파일을 만든다."""
import io
import os

from openpyxl import Workbook
from openpyxl.drawing.image import Image as XImage
from openpyxl.drawing.spreadsheet_drawing import AnchorMarker, OneCellAnchor
from openpyxl.drawing.xdr import XDRPositiveSize2D
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils.units import pixels_to_EMU

# 공급자 칸 (내용은 화면에서 입력한 값이 들어갑니다)
SUPPLIER_FIELDS = [
    ("등록번호", "regno"),
    ("상  호", "company"),
    ("대 표 자", "ceo"),
    ("주  소", "addr"),
    ("전화/팩스", "tel"),
    ("담 당 자", "manager"),
]

FONT = "맑은 고딕"
COLS = "ABCDEFGHIJ"
WIDTHS = [5, 12, 15, 18, 6, 6, 12, 13, 12, 12]
HEADERS = ["No", "이미지", "품  목", "규  격", "수량", "단위", "단가", "금액", "식별번호", "비  고"]
MIN_ROWS = 10
IMG = 58          # 사진 한 변(px)
IMG_CELL_W = 89   # 이미지 칸 폭(px): 열 너비 12
IMG_CELL_H = 76   # 품목 줄 높이(px): 57pt

_thin = Side(style="thin", color="000000")
BOX = Border(left=_thin, right=_thin, top=_thin, bottom=_thin)
LABEL = PatternFill("solid", fgColor="E4E9F2")
INPUT = PatternFill("solid", fgColor="FFF9E1")
NAVY = PatternFill("solid", fgColor="1F3864")
CENTER = Alignment(horizontal="center", vertical="center", wrap_text=True)
LEFT = Alignment(horizontal="left", vertical="center", wrap_text=True, indent=1)
RIGHT = Alignment(horizontal="right", vertical="center")


def korean_amount(n):
    """6096600 -> '육백구만육천육백'"""
    n = int(n)
    if n <= 0:
        return "영"
    digits = "영일이삼사오육칠팔구"
    small = ["", "십", "백", "천"]
    big = ["", "만", "억", "조"]
    res, g = "", 0
    while n > 0:
        chunk, n = n % 10000, n // 10000
        s = ""
        for i in range(4):
            d = chunk // (10 ** i) % 10
            if d:
                s = digits[d] + small[i] + s
        if s:
            res = s + big[g] + res
        g += 1
    return res


def build_quote(items, info=None, out=None):
    """items: [{name, spec, qty, unit, price, number, note, image}], info: {date, to, ref, valid, delivery}.
    out이 경로면 그 파일로 저장하고, 없으면 bytes를 돌려준다."""
    info = info or {}
    wb = Workbook()
    ws = wb.active
    ws.title = "견적서"
    ws.sheet_view.showGridLines = False
    for col, w in zip(COLS, WIDTHS):
        ws.column_dimensions[col].width = w

    def cell(ref, value=None, fill=None, align=CENTER, bold=False, size=10, color="000000", fmt=None):
        c = ws[ref]
        c.value = value
        c.font = Font(name=FONT, size=size, bold=bold, color=color)
        c.alignment = align
        if fill:
            c.fill = fill
        if fmt:
            c.number_format = fmt
        return c

    def merged(rng, value=None, border=True, **kw):
        ws.merge_cells(rng)
        c = cell(rng.split(":")[0], value, **kw)
        if border:
            for row in ws[rng]:
                for x in row:
                    x.border = BOX
                    if kw.get("fill"):
                        x.fill = kw["fill"]
        return c

    # 제목
    merged("A1:J1", "견   적   서", border=False, bold=True, size=26)
    ws.row_dimensions[1].height = 50
    for x in ws["A1:J1"][0]:
        x.border = Border(bottom=Side(style="medium", color="000000"))
    ws.row_dimensions[2].height = 6

    # 머리 정보
    to = (info.get("to") or "").strip()
    left_info = [
        ("견적일자", info.get("date") or ""),
        ("수  신", (to + "  귀하") if to else "귀하"),
        ("참  조", info.get("ref") or ""),
        ("유효기간", info.get("valid") or "견적일로부터 30일"),
        ("납  기", info.get("delivery") or "협의"),
    ]
    for i, (k, v) in enumerate(left_info):
        r = 3 + i
        ws.row_dimensions[r].height = 21
        merged(f"A{r}:B{r}", k, fill=LABEL, bold=True)
        merged(f"C{r}:E{r}", v, fill=INPUT, align=LEFT)
    ws.row_dimensions[8].height = 21
    merged("A8:E8", "아래와 같이 견적합니다.", border=False, align=LEFT, bold=True)
    for i, (k, key) in enumerate(SUPPLIER_FIELDS):
        r = 3 + i
        v = str(info.get(key) or "").strip()
        if key == "ceo":
            v = (v + "          (인)") if v else "(인)"
        merged(f"F{r}:G{r}", k, fill=LABEL, bold=True)
        merged(f"H{r}:J{r}", v, fill=INPUT, size=9 if len(v) > 16 else 10)
    ws.row_dimensions[9].height = 8

    # A4 한 장(위 10mm·아래 17mm 여백)을 꽉 채우도록 빈 줄 수와 줄 높이를 정한다
    width_pt = sum(int(w * 7 + 5) for w in WIDTHS) * 0.75
    scale = min(1.0, (190 / 25.4 * 72) / width_pt)
    fixed_pt = 50 + 6 + 21 * 6 + 8 + 32 + 8 + 24 + 26 + 20
    avail_pt = (270 / 25.4 * 72) / scale * 0.985 - fixed_pt
    rows = max(MIN_ROWS, len(items))
    row_h = 57
    if rows * row_h <= avail_pt:
        rows = max(rows, int(avail_pt // row_h))
        row_h = round(avail_pt / rows, 1)
    first, last = 13, 13 + rows - 1
    total = sum(int(it.get("qty") or 0) * int(it.get("price") or 0) for it in items)

    # 합계
    ws.row_dimensions[10].height = 32
    merged("A10:B10", "합계금액\n(VAT포함)", fill=LABEL, bold=True)
    merged("C10:F10", f"일금 {korean_amount(total)}원정", bold=True, size=12)
    merged("G10:J10", f"=SUM(H{first}:H{last})", bold=True, size=14, color="C00000", fmt='"₩"#,##0')
    ws.row_dimensions[11].height = 8

    # 표 머리
    ws.row_dimensions[12].height = 24
    for col, h in zip(COLS, HEADERS):
        cell(f"{col}12", h, fill=NAVY, bold=True, color="FFFFFF").border = BOX

    # 품목
    for i in range(rows):
        r = first + i
        ws.row_dimensions[r].height = row_h
        it = items[i] if i < len(items) else None
        cell(f"A{r}", i + 1)
        cell(f"B{r}")
        cell(f"C{r}", "\n".join(x for x in (it.get("name"), it.get("model")) if x) if it else None, fill=INPUT)
        cell(f"D{r}", it.get("spec") if it else None)
        cell(f"E{r}", int(it.get("qty") or 0) if it else None, fill=INPUT)
        cell(f"F{r}", (it.get("unit") or "EA") if it else None)
        cell(f"G{r}", int(it.get("price") or 0) if it else None, align=RIGHT, fmt="#,##0")
        cell(f"H{r}", f'=IF(E{r}="","",E{r}*G{r})', align=RIGHT, fmt="#,##0")
        cell(f"I{r}", it.get("number") if it else None, fill=INPUT)
        cell(f"J{r}", (it.get("note") or None) if it else None, size=9)
        for col in COLS:
            ws[f"{col}{r}"].border = BOX
        path = it.get("image") if it else None
        if path and os.path.exists(path):
            try:
                img = XImage(path)
                scale = IMG / max(img.width, img.height)
                w, h = max(1, int(img.width * scale)), max(1, int(img.height * scale))
                img.width, img.height = w, h
                cell_h = int(row_h * 4 / 3)
                marker = AnchorMarker(col=1, colOff=pixels_to_EMU((IMG_CELL_W - w) // 2),
                                      row=r - 1, rowOff=pixels_to_EMU(max(0, (cell_h - h) // 2)))
                img.anchor = OneCellAnchor(_from=marker, ext=XDRPositiveSize2D(pixels_to_EMU(w), pixels_to_EMU(h)))
                ws.add_image(img)
            except Exception:
                pass

    # 합계 줄 + 안내
    r = last + 1
    ws.row_dimensions[r].height = 26
    merged(f"A{r}:D{r}", "합  계", fill=LABEL, bold=True)
    cell(f"E{r}", f"=SUM(E{first}:E{last})", fill=LABEL, bold=True).border = BOX
    cell(f"F{r}", "EA", fill=LABEL, bold=True).border = BOX
    merged(f"G{r}:J{r}", f"=SUM(H{first}:H{last})", fill=LABEL, bold=True, size=11, align=RIGHT, fmt='"₩"#,##0')
    ws.row_dimensions[r + 1].height = 20
    merged(f"A{r+1}:J{r+1}", "※ 단가는 나라장터 종합쇼핑몰 등록 가격(VAT 포함)이며, 식별번호는 나라장터 물품식별번호입니다.",
           border=False, align=LEFT, size=9)

    ws.page_setup.paperSize = ws.PAPERSIZE_A4
    ws.page_setup.orientation = "portrait"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0 if len(items) > MIN_ROWS + 4 else 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.print_options.horizontalCentered = True
    ws.page_margins.left = ws.page_margins.right = 0.4  # 10mm
    ws.page_margins.top = 10 / 25.4      # 위 10mm
    ws.page_margins.bottom = 17 / 25.4   # 아래 17mm
    ws.print_area = f"A1:J{r+1}"

    if out:
        wb.save(out)
        return out
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


def _norm(v):
    return "".join(str(v or "").split())


_INFO_LABELS = {"견적일자": "date", "수신": "to", "참조": "ref", "유효기간": "valid", "납기": "delivery",
                "등록번호": "regno", "상호": "company", "대표자": "ceo", "주소": "addr",
                "전화/팩스": "tel", "담당자": "manager"}
_COL_LABELS = {"품목": "name", "품명": "name", "모델명": "model", "규격": "spec", "수량": "qty", "단위": "unit",
               "단가": "price", "식별번호": "number", "비고": "note"}


def _to_int(v):
    try:
        return int(float(str(v).replace(",", "")))
    except (TypeError, ValueError):
        return 0


def read_quote(data):
    """이 프로그램으로 저장한 견적서 엑셀(bytes)을 읽어 {info, items}로 돌려준다."""
    from openpyxl import load_workbook
    # read_only: 한셀 등 다른 프로그램에서 고쳐 저장한 파일도 서식 오류 없이 읽는다
    wb = load_workbook(io.BytesIO(data), read_only=True, data_only=True)
    ws = wb.worksheets[0]
    grid = []
    for row in ws.iter_rows(min_row=1, max_row=400, max_col=14, values_only=True):
        grid.append((list(row) + [None] * 14)[:14])
    wb.close()
    head = next((i for i, row in enumerate(grid) if _norm(row[0]).lower() == "no"), None)
    if head is None:
        raise ValueError("견적서 양식을 찾지 못했습니다. 이 프로그램으로 저장한 엑셀 파일인지 확인해 주세요.")

    info = {}
    for row in grid[:head]:
        for i, v in enumerate(row):
            key = _INFO_LABELS.get(_norm(v))
            if not key:
                continue
            val = ""
            for nxt in row[i + 1:]:
                if _norm(nxt) in _INFO_LABELS:
                    break
                if nxt not in (None, ""):
                    val = str(nxt).strip()
                    break
            info[key] = val
    to = info.get("to", "").replace("귀하", "").replace("(거래처명)", "").strip()
    info["to"] = to
    info["ceo"] = info.get("ceo", "").replace("(인)", "").strip()

    cols = {}
    for i, v in enumerate(grid[head]):
        key = _COL_LABELS.get(_norm(v))
        if key and key not in cols:
            cols[key] = i
    items = []
    for row in grid[head + 1:]:
        if not isinstance(row[0], (int, float)):
            break
        get = lambda k: row[cols[k]] if k in cols and row[cols[k]] is not None else ""
        name, model = str(get("name")).strip(), str(get("model")).strip()
        if "\n" in name and not model:
            name, model = [x.strip() for x in name.split("\n", 1)]
        number = "".join(ch for ch in str(get("number")) if ch.isdigit())
        if not (name or number):
            continue
        items.append({"name": name, "model": model, "spec": str(get("spec")).strip(), "qty": _to_int(get("qty")),
                      "unit": str(get("unit")).strip() or "EA", "price": _to_int(get("price")),
                      "number": number, "note": str(get("note")).strip()})
    return {"info": info, "items": items}
