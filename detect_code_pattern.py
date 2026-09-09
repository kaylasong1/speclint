# -*- coding: utf-8 -*-
"""
detect_code_pattern.py
화면정의서에서 '화면코드'로 쓰였을 법한 토큰 패턴 후보를 자동으로 찾아 제시한다.

기획자가 정규식을 직접 입력하지 않아도 되도록, 문서를 훑어 후보 패턴을 뽑고
출현 슬라이드 수 / 고유값 수 / 예시 / 추천 정규식을 보여주는 것이 목적.

입력은 extract_pptx.py가 만든 screens.json을 쓰는 것을 권장한다.
(그룹 도형 재귀, 읽는 순서 정렬, 개정이력 표 제외가 이미 적용돼 있음)
.pptx를 직접 넣으면 그 전처리 없이 단순 추출한다.

사용법:
    python extract_pptx.py 화면정의서.pptx -o out/screens.json
    python detect_code_pattern.py out/screens.json --csv out/code_pattern_candidates.csv
    python detect_code_pattern.py out/screens_front.json out/screens_bo.json --top 8

옵션:
    --top N          레벨별로 보여줄 후보 개수 (기본 6)
    --min-slides N   최소 출현 슬라이드 수 (기본 3)
    --csv PATH       후보 목록을 CSV로 저장
    --exclude A,B    후보에서 제외할 토큰(대문자 무관), 쉼표 구분
"""

import argparse
import csv
import json
import os
import re
import sys
from collections import defaultdict

# ---------------------------------------------------------------------------
# 0. 설정
# ---------------------------------------------------------------------------

# 화면코드처럼 보이지만 실제로는 용어/규격인 토큰들
DEFAULT_STOPWORDS = {
    "MP3", "MP4", "H264", "H265", "MPEG2", "MPEG4", "AAC", "WAV",
    "3G", "4G", "5G", "6G", "LTE", "WIFI", "IPTV", "VOD", "OTT",
    "HD", "FHD", "UHD", "SD", "4K", "8K", "3D", "2D",
    "PNG", "JPG", "JPEG", "GIF", "SVG", "PDF", "XLS", "XLSX", "PPT", "PPTX",
    "B2B", "B2C", "O2O", "API", "URL", "URI", "CTA", "KPI", "FAQ", "SLA",
    "ID", "PW", "OTP", "USIM", "ESIM", "SMS", "MMS", "RCS",
    "IOS", "IPV4", "IPV6", "UTF-8", "EUC-KR", "CSS3", "HTML5", "ES6",
    "RGB", "CMYK", "AS-IS", "TO-BE", "QA", "UX", "UI", "GNB", "LNB",
}

# 숫자+단위(50GB, 250MB, 100Mbps)나 버전(v0.5, V1.2)은 화면코드가 아니다.
UNIT_RE = re.compile(r"^\d+(?:\.\d+)?(?:GB|MB|KB|TB|BPS|MBPS|KBPS|GHZ|MHZ|PX|EM|REM|"
                     r"KRW|USD|WON|EA|SEC|MIN|HR|DAY|MON|YR|G|M|K|B|P|일|월|년)$",
                     re.IGNORECASE)
VERSION_RE = re.compile(r"^[vV]\d+(?:\.\d+)*$")

# 토큰 후보: 영숫자 덩어리가 . _ - / 로 이어진 형태
# "(ASIS)"처럼 코드 중간에 끼는 괄호 구간까지 한 토큰으로 본다.
# 여는/닫는 괄호를 짝으로만 삼키므로 "(F.mMY.1.0)"처럼 문장이 코드를 감싼
# 바깥 괄호는 먹지 않는다. (extract_pptx.py의 SCREEN_CODE_RE와 같은 원리)
TOKEN_RE = re.compile(
    r"[A-Za-z0-9]+(?:[._\-/][A-Za-z0-9]+|\([A-Za-z0-9_.\-]+\))*")
SEP_CHARS = "._-/"


# ---------------------------------------------------------------------------
# 1. 문서에서 (슬라이드번호, 블록순서, 텍스트) 뽑기
# ---------------------------------------------------------------------------

def _iter_shapes(shapes):
    """그룹 도형 안쪽까지 재귀적으로 훑는다."""
    for shape in shapes:
        if shape.shape_type == 6 and hasattr(shape, "shapes"):  # GROUP
            for inner in _iter_shapes(shape.shapes):
                yield inner
        else:
            yield shape


def read_pptx(path):
    """.pptx에서 텍스트 블록을 순서대로 뽑는다."""
    try:
        from pptx import Presentation
    except ImportError:
        sys.exit("python-pptx가 필요합니다:  pip install python-pptx")

    prs = Presentation(path)
    blocks = []
    for slide_no, slide in enumerate(prs.slides, start=1):
        order = 0
        for shape in _iter_shapes(slide.shapes):
            if getattr(shape, "has_table", False):
                for row in shape.table.rows:
                    for cell in row.cells:
                        text = (cell.text or "").strip()
                        if text:
                            blocks.append((slide_no, order, text))
                            order += 1
                continue
            if not getattr(shape, "has_text_frame", False):
                continue
            for para in shape.text_frame.paragraphs:
                text = "".join(run.text for run in para.runs).strip()
                if text:
                    blocks.append((slide_no, order, text))
                    order += 1
    return blocks


def read_screens_json(path):
    """extract_pptx.py가 만든 screens.json에서 텍스트 블록을 읽는 순서대로 뽑는다.

    - blocks / description_blocks / notes 의 text만 본다.
    - slide_id, screen_code, referenced_codes, urls 는 이미 F. 정규식으로 걸러진
      결과라 패턴 감지에 쓰면 순환 참조가 된다. 전부 무시한다.
    - blocks는 extract_pptx.py가 이미 읽는 순서로 정렬해 둔 상태라
      그 인덱스를 슬라이드 내 위치 신호로 그대로 쓴다.
    """
    with open(path, "r", encoding="utf-8") as f:
        data = json.load(f)

    if not isinstance(data, list):
        return read_json_generic(data)

    out = []
    for i, screen in enumerate(data, start=1):
        if not isinstance(screen, dict):
            continue
        slide = screen.get("source", {}).get("slide") or i
        order = 0
        for key in ("blocks", "description_blocks"):
            for b in screen.get(key) or []:
                text = (b.get("text") or "").strip()
                if text:
                    out.append((slide, order, text))
                    order += 1
        notes = (screen.get("notes") or "").strip()
        if notes:
            out.append((slide, order, notes))
    return out


def read_json_generic(data):
    """스키마를 모르는 JSON용 best-effort 폴백."""
    slide_keys = {"slide", "slide_no", "slide_number", "page", "page_no"}
    skip_keys = {"file", "filename", "path", "type", "shape_type", "slide_id",
                 "screen_code", "referenced_codes", "urls", "unreadable_shapes"}
    blocks = []
    counter = {"order": 0}

    def walk(node, cur_slide):
        if isinstance(node, dict):
            for k, v in node.items():
                if k.lower() in slide_keys and isinstance(v, (int, str)):
                    try:
                        cur_slide = int(str(v).strip())
                    except ValueError:
                        pass
            for k, v in node.items():
                if k.lower() in skip_keys:
                    continue
                walk(v, cur_slide)
        elif isinstance(node, list):
            for item in node:
                walk(item, cur_slide)
        elif isinstance(node, str):
            text = node.strip()
            if text:
                blocks.append((cur_slide, counter["order"], text))
                counter["order"] += 1

    walk(data, 0)
    if all(b[0] == 0 for b in blocks):
        blocks = [(i // 20 + 1, i, t) for i, (_, _, t) in enumerate(blocks)]
    return blocks


def read_any(path):
    ext = os.path.splitext(path)[1].lower()
    if ext == ".pptx":
        return read_pptx(path)
    if ext == ".json":
        return read_screens_json(path)
    sys.exit("지원하지 않는 형식입니다(.pptx 또는 .json): %s" % path)


# ---------------------------------------------------------------------------
# 2. 토큰 → 마스크
# ---------------------------------------------------------------------------

def char_class(ch):
    if ch.isdigit():
        return "9"
    if ch.isupper():
        return "A"
    if ch.islower():
        return "a"
    return ch


def mask_exact(token):
    """BO_M_0607 -> AA_A_9999"""
    return "".join(char_class(c) for c in token)


def mask_run(token):
    """BO_M_0607 -> A+_A+_9+   (연속된 같은 종류를 하나로 압축)"""
    out = []
    prev = None
    for c in token:
        cls = char_class(c)
        if cls in "Aa9":
            if cls != prev:
                out.append(cls + "+")
        else:
            out.append(cls)
        prev = cls if cls in "Aa9" else None
    return "".join(out)


def masks_prefix(token):
    """F.mMY.PP_MOBILE.2.0 -> ['F.', 'F.mMY.', 'F.mMY.PP_MOBILE.', ...]

    실무 화면코드는 자릿수보다 '접두부'로 식별된다("F.로 시작", "BO_로 시작").
    구분자 경계마다 접두부를 만들어 후보로 올린다.
    """
    out = []
    for i, ch in enumerate(token):
        if ch in SEP_CHARS and 0 < i < len(token) - 1:
            out.append(token[: i + 1])
    return out[:4]  # 너무 긴 접두부(=거의 토큰 전체)는 의미 없음


def mask_shape(token):
    """BO_M_0607 -> '_ 구분 3칸, 숫자 포함'  (가장 느슨한 골격)"""
    seps = sorted({c for c in token if c in SEP_CHARS})
    n_seg = len([s for s in re.split("[%s]" % re.escape(SEP_CHARS), token) if s])
    has_digit = any(c.isdigit() for c in token)
    sep_label = "".join(seps) if seps else "없음"
    return "구분자 %s / %d칸 / 숫자 %s" % (sep_label, n_seg, "O" if has_digit else "X")


# ---------------------------------------------------------------------------
# 3. 마스크 → 추천 정규식
# ---------------------------------------------------------------------------

BOUND_L = r"(?<![A-Za-z0-9._\-/])"
BOUND_R = r"(?![A-Za-z0-9._\-/])"


def regex_from_exact(mask):
    out = []
    i = 0
    while i < len(mask):
        c = mask[i]
        if c in "Aa9":
            j = i
            while j < len(mask) and mask[j] == c:
                j += 1
            n = j - i
            base = {"A": "[A-Z]", "a": "[a-z]", "9": r"\d"}[c]
            out.append(base + ("{%d}" % n if n > 1 else ""))
            i = j
        else:
            out.append(re.escape(c))
            i += 1
    return BOUND_L + "".join(out) + BOUND_R


def regex_from_run(mask):
    out = []
    i = 0
    while i < len(mask):
        if mask[i] in "Aa9" and i + 1 < len(mask) and mask[i + 1] == "+":
            out.append({"A": "[A-Z]+", "a": "[a-z]+", "9": r"\d+"}[mask[i]])
            i += 2
        else:
            out.append(re.escape(mask[i]))
            i += 1
    return BOUND_L + "".join(out) + BOUND_R


def regex_from_prefix(prefix):
    """접두부 + 나머지. 괄호 구간((ASIS) 등)을 삼키도록 tail을 붙인다."""
    tail = r"[A-Za-z0-9_.\-]+(?:\([A-Za-z0-9_.\-]+\)[A-Za-z0-9_.\-]*)*"
    return BOUND_L + re.escape(prefix) + tail


def regex_from_shape(label, sample):
    seps = sorted({c for c in sample if c in SEP_CHARS})
    n_seg = len([s for s in re.split("[%s]" % re.escape(SEP_CHARS), sample) if s])
    sep_cls = "[%s]" % re.escape("".join(seps)) if seps else ""
    if not seps:
        body = "[A-Za-z0-9]+"
    else:
        body = sep_cls.join(["[A-Za-z0-9]+"] * n_seg)
    digit_req = r"(?=[A-Za-z0-9._\-/]*\d)" if "숫자 O" in label else ""
    return BOUND_L + digit_req + body + BOUND_R


# ---------------------------------------------------------------------------
# 4. 후보 수집 및 점수
# ---------------------------------------------------------------------------

class Group(object):
    def __init__(self, mask=""):
        self.mask = mask
        self.slides = set()
        self.values = defaultdict(int)
        self.first_pos = []      # 슬라이드별 첫 등장 블록순서(정규화 전)
        self.occurrences = 0

    @property
    def n_slides(self):
        return len(self.slides)

    @property
    def n_values(self):
        return len(self.values)

    def examples(self, k=3):
        ordered = sorted(self.values.items(), key=lambda kv: (-kv[1], kv[0]))
        return [v for v, _ in ordered[:k]]


def is_candidate(token, stopwords):
    if len(token) < 3:
        return False
    if token.upper() in stopwords:
        return False
    if token.isdigit():
        return False
    has_digit = any(c.isdigit() for c in token)
    n_sep = sum(1 for c in token if c in SEP_CHARS)
    # 숫자가 있거나, 구분자가 2개 이상인 토큰만 후보로 본다
    if not has_digit and n_sep < 2:
        return False
    # 연도/버전처럼 보이는 순수 숫자+점 (2026.09.09)
    if re.fullmatch(r"[0-9._\-/]+", token):
        return False
    if UNIT_RE.match(token):
        return False
    if VERSION_RE.match(token):
        return False
    return True


def collect(blocks, stopwords):
    total_slides = len({b[0] for b in blocks}) or 1
    max_order = defaultdict(int)
    for slide, order, _ in blocks:
        max_order[slide] = max(max_order[slide], order)

    levels = {
        "L0 정확한 형태": (lambda t: [mask_exact(t)], defaultdict(Group)),
        "L1 압축 형태": (lambda t: [mask_run(t)], defaultdict(Group)),
        "L2 느슨한 골격": (lambda t: [mask_shape(t)], defaultdict(Group)),
        "L3 접두부": (masks_prefix, defaultdict(Group)),
    }

    seen_first = {}  # (level, mask, slide) -> 최초 order
    for slide, order, text in blocks:
        for token in TOKEN_RE.findall(text):
            token = token.strip(SEP_CHARS)
            if not is_candidate(token, stopwords):
                continue
            for name, (fn, store) in levels.items():
                for mask in fn(token):
                    g = store[mask]
                    if not g.mask:
                        g.mask = mask
                    g.slides.add(slide)
                    g.values[token] += 1
                    g.occurrences += 1
                    key = (name, mask, slide)
                    if key not in seen_first:
                        seen_first[key] = order
                        denom = max_order[slide] or 1
                        g.first_pos.append(order / float(denom))
    return levels, total_slides


def score(group, total_slides):
    coverage = group.n_slides / float(total_slides)
    uniq_per_slide = group.n_values / float(group.n_slides)
    occ_per_slide = group.occurrences / float(group.n_slides)
    top_pos = sum(group.first_pos) / len(group.first_pos) if group.first_pos else 0.5

    s = coverage * 100
    if 0.6 <= uniq_per_slide <= 2.0:
        s += 30                      # 슬라이드마다 값이 거의 1:1 → 화면코드다움
    if occ_per_slide <= 3.0:
        s += 15                      # 한 슬라이드에 수십 번 나오면 코드가 아님
    s += 25 * (1 - min(top_pos, 1.0))  # 슬라이드 상단에 나올수록 가산
    if group.n_values <= 2:
        s -= 30                      # 값이 1~2개뿐이면 상수/용어
    return s, coverage, uniq_per_slide, occ_per_slide, top_pos


# ---------------------------------------------------------------------------
# 5. 출력
# ---------------------------------------------------------------------------

def _w(text):
    """한글 등 전각 문자를 2칸으로 계산한 표시 폭."""
    import unicodedata
    return sum(2 if unicodedata.east_asian_width(c) in "WF" else 1 for c in text)


def _pad(text, width):
    return text + " " * max(0, width - _w(text))


def report(path, levels, total_slides, top, min_slides, rows):
    print("\n" + "=" * 78)
    print("파일: %s   (텍스트가 있는 슬라이드 %d개)" % (path, total_slides))
    print("=" * 78)

    for name, (fn, store) in levels.items():
        cands = []
        for mask, g in store.items():
            if g.n_slides < min_slides:
                continue
            sc, cov, ups, ops, pos = score(g, total_slides)
            cands.append((sc, mask, g, cov, ups, ops, pos))
        cands.sort(key=lambda x: -x[0])

        print("\n[%s]" % name)
        if not cands:
            print("  후보 없음 (min-slides=%d)" % min_slides)
            continue
        print("  %s %8s %6s %7s  %s" % (_pad("패턴", 30), "슬라이드", "고유값", "점수", "예시"))
        for sc, mask, g, cov, ups, ops, pos in cands[:top]:
            ratio = g.n_values / float(g.n_slides)
            note = "  ← 값 종류가 적음(수정이력 표시일 수 있음)" if ratio < 0.3 else ""
            ex = ", ".join(g.examples(3)) + note
            print("  %s %8d %6d %7.1f  %s" % (_pad(mask[:30], 30), g.n_slides, g.n_values, sc, ex))

            sample = g.examples(1)[0]
            if name.startswith("L0"):
                rgx = regex_from_exact(mask)
            elif name.startswith("L1"):
                rgx = regex_from_run(mask)
            elif name.startswith("L3"):
                rgx = regex_from_prefix(mask)
            else:
                rgx = regex_from_shape(mask, sample)
            rows.append({
                "file": os.path.basename(path),
                "level": name,
                "pattern": mask,
                "slides": g.n_slides,
                "unique_values": g.n_values,
                "occurrences": g.occurrences,
                "coverage": round(cov, 3),
                "uniq_per_slide": round(ups, 2),
                "occ_per_slide": round(ops, 2),
                "avg_position": round(pos, 3),
                "score": round(sc, 1),
                "examples": " | ".join(g.examples(5)),
                "suggested_regex": rgx,
            })

        best = cands[0]
        sample = best[2].examples(1)[0]
        if name.startswith("L0"):
            rgx = regex_from_exact(best[1])
        elif name.startswith("L1"):
            rgx = regex_from_run(best[1])
        elif name.startswith("L3"):
            rgx = regex_from_prefix(best[1])
        else:
            rgx = regex_from_shape(best[1], sample)
        print("  → 추천 정규식: %s" % rgx)

    # 사용자에게 물어볼 문장
    l1 = levels["L3 접두부"][1]
    best = None
    for mask, g in l1.items():
        if g.n_slides < min_slides:
            continue
        sc = score(g, total_slides)[0]
        if best is None or sc > best[0]:
            best = (sc, mask, g)
    if best:
        _, mask, g = best
        print("\n  확인 문구 예시:")
        print("    \"이 문서에서 `%s`로 시작하는 코드를 %d개 슬라이드에서 "
              "%d종 발견했습니다(예: %s). 이걸 화면코드로 사용할까요?\""
              % (mask, g.n_slides, g.n_values, ", ".join(g.examples(2))))


def detect_from_screens(screens, top=5, min_slides=3, min_value_ratio=0.3,
                        levels_wanted=("L3 접두부",), extra_stopwords=None):
    """app.py 등에서 쓰는 함수 API.

    screens: extract_pptx.py가 만든 화면 dict 리스트(파일이 아니라 메모리 객체).
    반환: [{"pattern","slides","values","examples","regex","score"}, ...] 점수 내림차순.
    """
    stopwords = set(DEFAULT_STOPWORDS) | {w.upper() for w in (extra_stopwords or [])}

    blocks = []
    for i, screen in enumerate(screens, start=1):
        slide = (screen.get("source") or {}).get("slide") or i
        order = 0
        for key in ("blocks", "description_blocks"):
            for b in screen.get(key) or []:
                text = (b.get("text") or "").strip()
                if text:
                    blocks.append((slide, order, text))
                    order += 1

    levels, total_slides = collect(blocks, stopwords)
    out = []
    for name in levels_wanted:
        store = levels[name][1]
        for mask, g in store.items():
            if g.n_slides < min_slides:
                continue
            # 화면코드는 장표마다 값이 달라야 한다. 슬라이드는 많은데 값 종류가
            # 적으면 수정이력 표시(BO_M_0607을 여러 장표에 반복 표기) 쪽이다.
            if g.n_values / float(g.n_slides) < min_value_ratio:
                continue
            sc = score(g, total_slides)[0]
            out.append({
                "level": name,
                "pattern": mask,
                "slides": g.n_slides,
                "values": g.n_values,
                "examples": g.examples(3),
                "regex": regex_from_prefix(mask) if name.startswith("L3")
                         else regex_from_run(mask),
                "score": round(sc, 1),
            })
    out.sort(key=lambda r: -r["score"])

    # 같은 커버리지의 접두부가 여러 개면(F. / F.mMY. / F.mMY.PP_) 가장 짧은 것만 남긴다
    deduped, seen = [], set()
    for r in out:
        key = (r["slides"], r["values"])
        if key in seen:
            continue
        seen.add(key)
        deduped.append(r)
    return deduped[:top]


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("paths", nargs="+")
    ap.add_argument("--top", type=int, default=6)
    ap.add_argument("--min-slides", type=int, default=3)
    ap.add_argument("--csv", default=None)
    ap.add_argument("--exclude", default="")
    args = ap.parse_args()

    stopwords = set(DEFAULT_STOPWORDS)
    stopwords |= {w.strip().upper() for w in args.exclude.split(",") if w.strip()}

    rows = []
    for path in args.paths:
        if not os.path.exists(path):
            print("파일 없음: %s" % path)
            continue
        blocks = read_any(path)
        if not blocks:
            print("텍스트를 찾지 못했습니다: %s" % path)
            continue
        levels, total_slides = collect(blocks, stopwords)
        report(path, levels, total_slides, args.top, args.min_slides, rows)

    if args.csv and rows:
        os.makedirs(os.path.dirname(args.csv) or ".", exist_ok=True)
        with open(args.csv, "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=list(rows[0].keys()))
            w.writeheader()
            w.writerows(rows)
        print("\nCSV 저장: %s (%d행)" % (args.csv, len(rows)))


if __name__ == "__main__":
    main()
