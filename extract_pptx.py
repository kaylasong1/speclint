"""
화면정의서(PPTX) -> 중간 JSON 추출기 (1단계)

목표: 슬라이드 1장 = 화면 1개로 보고, 그 안의 모든 텍스트를 읽는 순서대로 뽑아
      뒷단(체크리스트 검토)이 쓸 수 있는 공통 JSON을 만든다.

사용법:
    pip install python-pptx
    python extract_pptx.py 화면정의서.pptx -o screens.json
"""

import argparse
import json
import re
from pathlib import Path

from pptx import Presentation
from pptx.enum.shapes import MSO_SHAPE_TYPE

# 1인치 = 914400 EMU. 같은 줄로 볼 세로 오차 범위(0.2인치).
ROW_TOLERANCE = 182880

# PPTX 텍스트 상자에서 줄바꿈이 \x0b(vertical tab)로 들어오는 경우가 있다.
VERTICAL_TAB_RE = re.compile("\x0b")

# 화면 코드: "F."로 시작해서 영문/숫자/기호(._-)와 "(ASIS)" 같은 괄호 구간까지.
# \w는 유니코드까지 잡아서 뒤에 조사가 붙은 한글(예: "...1.0를")까지
# 삼켜버리므로 ASCII로 제한한다.
# 괄호는 "(내용)" 형태로 코드 중간에 붙을 때만 삼키고, 문장에서 코드 전체를
# 감싸는 바깥쪽 닫는 괄호(...)는 삼키지 않도록 여는/닫는 괄호를 짝지어 매칭한다.
DEFAULT_SCREEN_CODE_PATTERN = r"F\.[A-Za-z0-9_.\-]+(?:\([A-Za-z0-9_.\-]+\)[A-Za-z0-9_.\-]*)*"
SCREEN_CODE_RE = re.compile(DEFAULT_SCREEN_CODE_PATTERN)
# URL 경로: "/segment"가 2번 이상 이어지고, 각 segment가 영문 소문자·숫자·
# 하이픈으로만 이뤄진 경우만 인정한다. "현역병사/복지/청소년/3G(CDMA)" 같은
# 한글 나열 중 "/3G"만 뚝 떼어 오탐하는 걸 막기 위함(segment 1개, 대문자 포함
# 이라 둘 다 걸러진다). 앞에 영문/숫자/슬래시/콜론이 있으면 날짜(2024/01/01),
# 분수(3/4), "https://" 안의 "//" 조각 같은 오탐이라 제외한다. 쿼리스트링은
# 흔히 camelCase라 이 규칙에서 제외하고 뒤에 그대로 붙여준다.
URL_RE = re.compile(r"(?<![A-Za-z0-9/:])(?:/[a-z0-9\-]+){2,}(?:\?[A-Za-z0-9_=&%.\-]+)?")

# 슬라이드 상단 이 비율 안에 있는 F. 코드를 그 화면의 "정식" 코드로 본다.
HEADER_RATIO = 0.2

# --- 표 분류 ---
# 표를 위치(EMU)만으로 나누면 문서마다 레이아웃이 달라 어긋난다.
# (실제로 백오피스 문서는 설명 표가 left 7.5M라 아래 8M 기준을 못 넘었다)
# 그래서 표의 '구조'로 먼저 판단하고, 판단이 안 되는 표만 위치 규칙으로 넘긴다.
#   설명 표   : 2열이고 첫 열에 항목 번호(1,2,3...)가 있거나 머리행이 화면 경로("A > B")
#   개정이력 표: 2열이고 첫 열이 전부 코드/버전 토큰(BO_M_0607, V1.4 ...)
#   그 외      : 본문(화면 목업 안의 데이터 그리드 등)
TABLE_CODE_CELL_RE = re.compile(r"^[A-Za-z0-9_.\-]+$")
TABLE_ITEM_NO_RE = re.compile(r"^\d{1,2}$")

# 이 left(EMU) 이상에 있는 표는 본문이 아니라 슬라이드 오른쪽에 붙는
# 설명/코멘트 표로 보고 description_blocks로 따로 뺀다.
DESCRIPTION_LEFT_MIN = 8_000_000
# 이 left(EMU) 이상은 개정이력 표(V0.1, V0.2 ... 변경사항 기록)라 기능 정의가
# 아니다. description_blocks에도 넣지 않고 통째로 제외한다.
REVISION_LEFT_MIN = 11_000_000


def classify_table(rows):
    """표를 'description' / 'revision' / 'body'로 나눈다. 판단 불가면 'body'."""
    if not rows:
        return "body"
    if max(len(r) for r in rows) != 2:
        return "body"

    firsts = [(r[0] or "").strip() for r in rows if (r[0] or "").strip()]
    if not firsts:
        return "body"
    if all(TABLE_CODE_CELL_RE.match(c) for c in firsts):
        return "revision"
    if any(TABLE_ITEM_NO_RE.match(c) for c in firsts) or ">" in firsts[0]:
        return "description"
    return "body"


def normalize_text(text):
    """세로탭(\\x0b)을 줄바꿈으로 치환."""
    return VERTICAL_TAB_RE.sub("\n", text)


def dedupe_preserve_order(items):
    seen = set()
    result = []
    for item in items:
        if item not in seen:
            seen.add(item)
            result.append(item)
    return result


def walk_shapes(shapes, inherited_pos=None):
    """
    그룹 도형 안까지 재귀로 들어가서 모든 도형을 평평하게 꺼낸다.
    이걸 안 하면 그룹 안 텍스트가 통째로 누락된다. (가장 흔한 사고)

    주의: 그룹 자식의 top/left는 슬라이드 좌표가 아니라 그룹 내부 좌표라
    정렬이 어긋난다. 그래서 자식은 부모 그룹의 위치를 물려받게 했다.
    (정확도가 필요해지면 나중에 좌표 변환을 넣는다)
    """
    for shape in shapes:
        if shape.shape_type == MSO_SHAPE_TYPE.GROUP:
            pos = (shape.top or 0, shape.left or 0)
            yield from walk_shapes(shape.shapes, inherited_pos or pos)
        else:
            pos = inherited_pos or (shape.top or 0, shape.left or 0)
            yield shape, pos


def read_order_key(item):
    """z-order(삽입 순서)가 아니라 화면 배치 순서대로 정렬한다."""
    _, (top, left) = item
    return (top // ROW_TOLERANCE, left)


def find_header_screen_code(blocks, block_tops, slide_height, code_re=None):
    """
    슬라이드 상단 HEADER_RATIO(20%) 안에 있는 블록 중 읽는 순서상 가장 먼저
    나오는 F. 코드를 그 화면의 screen_code로 본다. 라벨 텍스트에 기대지 않고
    순전히 위치로 판단한다.
    """
    code_re = code_re or SCREEN_CODE_RE
    header_limit = slide_height * HEADER_RATIO
    for bi, (block, top) in enumerate(zip(blocks, block_tops)):
        if top > header_limit:
            continue
        candidates = code_re.findall(block["text"])
        if candidates:
            return candidates[0], bi
    return "", None


def extract_table(shape):
    rows = []
    for row in shape.table.rows:
        rows.append([normalize_text(cell.text.strip()) for cell in row.cells])
    text = "\n".join(" | ".join(r) for r in rows)
    return rows, text


def extract_slide(slide, index, slide_height, code_re=None):
    code_re = code_re or SCREEN_CODE_RE

    blocks = []
    description_blocks = []  # 오른쪽에 붙는 설명/코멘트 표 (표 + left >= DESCRIPTION_LEFT_MIN)
    block_tops = []  # blocks와 같은 순서로 대응하는 슬라이드 세로 위치(top, EMU)
    unreadable = []  # 이미지/차트 등 텍스트를 읽을 수 없는 도형 (커버리지 확인용)

    items = sorted(walk_shapes(slide.shapes), key=read_order_key)

    for shape, (top, _left) in items:
        # 도형 자체의 위치/크기(EMU). 그룹 내부 도형은 그룹 좌표계 기준이라
        # 슬라이드 절대 좌표와는 어긋날 수 있다(정렬용 top과 별개로 원본 그대로 기록).
        geometry = {
            "left": shape.left or 0,
            "top": shape.top or 0,
            "width": shape.width or 0,
            "height": shape.height or 0,
        }

        # 표는 has_text_frame이 False라 반드시 따로 분기해야 한다.
        if getattr(shape, "has_table", False):
            rows, text = extract_table(shape)
            if text.strip():
                block = {"type": "table", "rows": rows, "text": text, **geometry}
                kind = classify_table(rows)
                if kind == "revision":
                    pass  # 개정이력 표: 기능 정의가 아니므로 완전히 제외
                elif kind == "description":
                    description_blocks.append(block)
                elif geometry["left"] >= REVISION_LEFT_MIN:
                    pass  # 구조로 판단 안 되는 표는 기존 위치 규칙으로 보조 판정
                elif geometry["left"] >= DESCRIPTION_LEFT_MIN:
                    description_blocks.append(block)
                else:
                    blocks.append(block)
                    block_tops.append(top)
            continue

        if shape.shape_type in (MSO_SHAPE_TYPE.PICTURE, MSO_SHAPE_TYPE.CHART):
            unreadable.append(shape.name)
            continue

        if shape.has_text_frame:
            paragraphs = [normalize_text(p.text.strip()) for p in shape.text_frame.paragraphs]
            text = "\n".join(p for p in paragraphs if p)
            if text:
                blocks.append({"type": "shape", "text": text, **geometry})
                block_tops.append(top)

    notes = ""
    if slide.has_notes_slide:
        notes = normalize_text(slide.notes_slide.notes_text_frame.text.strip())

    # screen_code: 상단 헤더 영역에서 찾은 F. 코드. screen_name: 그 코드
    # 바로 앞 텍스트 블록. 헤더에 코드가 없으면 둘 다 빈 값으로 둔다.
    screen_code, code_block_index = find_header_screen_code(
        blocks, block_tops, slide_height, code_re
    )
    if code_block_index is not None and code_block_index > 0:
        screen_name = blocks[code_block_index - 1]["text"].split("\n")[0].strip()
    else:
        screen_name = ""

    all_text = "\n".join(b["text"] for b in blocks + description_blocks) + "\n" + notes
    urls = dedupe_preserve_order(URL_RE.findall(all_text))

    # 헤더에서 뽑은 코드를 제외한 나머지 F. 코드(다른 화면 참조 등)는
    # referenced_codes로 분리한다.
    all_codes = dedupe_preserve_order(code_re.findall(all_text))
    referenced_codes = [c for c in all_codes if c != screen_code]

    return {
        "slide_id": f"S{index:02d}",
        "screen_name": screen_name,
        "source": {"slide": index},
        "blocks": blocks,
        "description_blocks": description_blocks,
        "notes": notes,
        "screen_code": screen_code,
        "referenced_codes": referenced_codes,
        "urls": urls,
        "unreadable_shapes": unreadable,
    }


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", help="화면정의서 .pptx 경로")
    parser.add_argument("-o", "--output", default="screens.json")
    parser.add_argument(
        "--code-pattern",
        default=DEFAULT_SCREEN_CODE_PATTERN,
        help="화면코드 정규식. detect_code_pattern.py가 추천한 값을 넣는다. "
             "기본값은 'F.'로 시작하는 코드 체계.",
    )
    args = parser.parse_args()

    code_re = re.compile(args.code_pattern)

    prs = Presentation(args.input)
    src = Path(args.input).name
    slide_height = prs.slide_height

    screens = []
    for i, slide in enumerate(prs.slides, start=1):
        screen = extract_slide(slide, i, slide_height, code_re)
        screen["source"]["file"] = src
        screens.append(screen)

    Path(args.output).write_text(
        json.dumps(screens, ensure_ascii=False, indent=2), encoding="utf-8"
    )

    # 원본과 눈으로 대조할 때 쓸 요약
    total_blocks = sum(len(s["blocks"]) for s in screens)
    total_unreadable = sum(len(s["unreadable_shapes"]) for s in screens)
    print(f"슬라이드 {len(screens)}개 / 텍스트 블록 {total_blocks}개 -> {args.output}")
    with_code = sum(1 for s in screens if s["screen_code"])
    print(f"텍스트를 못 읽은 도형(이미지·차트) {total_unreadable}개")
    print(f"화면코드가 잡힌 슬라이드 {with_code}개 (패턴: {args.code_pattern[:60]})")
    if with_code == 0:
        print("  [확인 필요] 화면코드를 하나도 못 찾았습니다. "
              "detect_code_pattern.py로 이 문서의 코드 패턴을 확인하세요.")
    for s in screens:
        if not s["blocks"]:
            print(f"  [확인 필요] {s['slide_id']} 텍스트가 하나도 안 잡힘")


if __name__ == "__main__":
    main()
