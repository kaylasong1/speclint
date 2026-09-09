"""
screens.json 검사기 (2단계)

extract_pptx.py가 뽑은 out/screens.json을 훑어서 사람이 눈으로 찾기 힘든
불일치를 찾는다.

검사 1) 깨진 참조 - referenced_codes에 나오는데 어느 화면의 screen_code로도
        정의되지 않은 코드
검사 2) 중복 화면 - 같은 screen_code를 가진 장표가 둘 이상이고, screen_name까지
        똑같은 경우 (같은 화면을 실수로 두 번 정의했을 가능성)
검사 3) 화면코드 누락 - 화면정의 페이지로 보이는데 screen_code가 비어 있는 경우

사용법:
    python check_spec.py [out/screens.json] [-o out/issues.csv]
"""

import argparse
import csv
import json
import re
from collections import defaultdict
from pathlib import Path

FIELDNAMES = ["slide_id", "screen_name", "issue_type", "detail"]

# 이 어휘가 포함된 텍스트 블록이 있으면 "기능 정의로 보이는 페이지"로 판단한다.
FUNCTIONAL_KEYWORDS = ["클릭 시", "노출", "이동", "버튼", "팝업"]

# 본문(blocks) 앞쪽에 이 표시가 있으면 스코프 아웃된 장표로 본다.
# 설명 표(description_blocks)에는 원래 기획 내용이 그대로 남아 있어서
# 정상 화면과 구분이 안 되므로, 반드시 본문 좌측 상단을 봐야 한다.
DESCOPED_KEYWORDS = ["미진행", "미적용", "대상 제외", "제외되어", "보류",
                     "삭제됨", "해당 없음", "해당없음", "차수 이관",
                     "그대로 유지", "추가하지 않는"]

# 스코프 아웃 표시를 찾을 때 볼 본문 블록 수(좌측 상단 몇 개).
DESCOPED_HEAD_BLOCKS = 3

# 설명 표의 항목 번호("1 | ...", "2 | ...")를 읽기 위한 패턴.
# 설명이 길어 다음 장표로 넘긴 경우 번호가 이어지므로(1,2 -> 3),
# 같은 화면코드·화면명이 두 장표에 있어도 중복이 아니라 '분할'로 본다.
ITEM_NUMBER_RE = re.compile(r"^\s*(\d{1,2})\s*\|")

# 표지·목차 등 짧은 페이지를 걸러내기 위한 최소 텍스트 길이.
MIN_TEXT_LENGTH = 100

# detail에 붙이는 텍스트 미리보기 길이.
DETAIL_PREVIEW_LENGTH = 100


def load_screens(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def check_broken_references(screens):
    """referenced_codes 중 어느 화면의 screen_code로도 정의 안 된 코드를 찾는다."""
    defined = {s["screen_code"] for s in screens if s.get("screen_code")}

    issues = []
    for s in screens:
        for code in s.get("referenced_codes", []):
            if code in defined:
                continue
            issues.append({
                "slide_id": s["slide_id"],
                "screen_name": s.get("screen_name", ""),
                "issue_type": "broken_reference",
                "detail": f"참조한 코드 '{code}'가 어느 화면의 screen_code로도 정의되어 있지 않음",
            })
    return issues


def item_numbers(screen):
    """설명 표의 항목 번호 목록. 없으면 빈 리스트."""
    numbers = []
    for block in screen.get("description_blocks") or []:
        for line in block.get("text", "").split("\n"):
            match = ITEM_NUMBER_RE.match(line)
            if match:
                numbers.append(int(match.group(1)))
    return numbers


def is_split_pair(prev_screen, next_screen):
    """앞 장표에서 이어서 쓴 '분할 장표'인지 판단한다.

    조건 두 가지를 모두 만족해야 한다.
      1) 슬라이드가 바로 옆에 붙어 있다
      2) 설명 표의 항목 번호가 이어진다 (앞의 마지막 < 뒤의 처음)
    번호를 못 읽으면 판단하지 않고 중복으로 남긴다(놓치는 쪽보다 안전).
    """
    prev_no = (prev_screen.get("source") or {}).get("slide")
    next_no = (next_screen.get("source") or {}).get("slide")
    if prev_no is None or next_no is None or next_no != prev_no + 1:
        return False

    prev_nums = item_numbers(prev_screen)
    next_nums = item_numbers(next_screen)
    if not prev_nums or not next_nums:
        return False
    return max(prev_nums) < min(next_nums)


def check_duplicate_screens(screens):
    """screen_code와 screen_name이 둘 다 같은 장표가 2개 이상 있는지 찾는다.
    단, 설명이 길어 다음 장표로 넘긴 '분할 장표'는 중복이 아니다."""
    groups = defaultdict(list)
    for s in screens:
        code = s.get("screen_code")
        name = s.get("screen_name")
        if not code or not name:
            continue
        groups[(code, name)].append(s)

    issues = []
    for (code, name), members in groups.items():
        if len(members) < 2:
            continue

        members = sorted(members, key=lambda s: (s.get("source") or {}).get("slide", 0))
        # 앞뒤가 전부 '이어 쓴' 관계면 한 화면을 나눠 쓴 것이므로 넘어간다.
        if all(is_split_pair(a, b) for a, b in zip(members, members[1:])):
            continue

        slide_ids = [s["slide_id"] for s in members]
        for slide_id in slide_ids:
            others = ", ".join(sid for sid in slide_ids if sid != slide_id)
            issues.append({
                "slide_id": slide_id,
                "screen_name": name,
                "issue_type": "duplicate_screen",
                "detail": f"screen_code '{code}'와 화면명이 같은 장표가 더 있음: {others}",
            })
    return issues


def slide_visible_text(screen):
    """오른쪽 번호-설명 표(description_blocks)의 텍스트만 본다.
    화면정의 페이지는 이 표에 "클릭 시", "노출" 같은 기능 정의 문구가 들어있고,
    표지·플로우도·별첨 등 화면정의가 아닌 페이지는 이 표 자체가 비어 있다.
    본문(blocks)까지 포함하면 Action 정의표나 케이스 비교표처럼 화면코드가
    필요 없는 페이지까지 걸려 오탐이 크게 늘어난다."""
    blocks = screen.get("description_blocks", [])
    return "\n".join(b.get("text", "") for b in blocks)


def is_descoped(screen):
    """'해당 케이스 미진행' 같은 표시가 본문 앞쪽에 있으면 스코프 아웃된 장표."""
    head = "\n".join(
        b.get("text", "") for b in (screen.get("blocks") or [])[:DESCOPED_HEAD_BLOCKS]
    )
    return any(keyword in head for keyword in DESCOPED_KEYWORDS)


def check_missing_screen_code(screens):
    """screen_code가 없는데 화면정의 페이지로 보이는(기능 정의 어휘가 있고
    텍스트가 충분히 긴) 슬라이드를 찾는다."""
    issues = []
    for s in screens:
        if s.get("screen_code"):
            continue
        # 작성자가 의도적으로 코드를 뺀 장표(미진행/제외)는 누락이 아니다.
        if is_descoped(s):
            continue

        text = slide_visible_text(s)
        if len(text) < MIN_TEXT_LENGTH:
            continue
        if not any(keyword in text for keyword in FUNCTIONAL_KEYWORDS):
            continue

        preview = text[:DETAIL_PREVIEW_LENGTH]
        issues.append({
            "slide_id": s["slide_id"],
            "screen_name": s.get("screen_name", ""),
            "issue_type": "missing_screen_code",
            "detail": f"화면정의 페이지로 보이나 화면코드가 없음 (텍스트: {preview})",
        })
    return issues


def run_all_checks(screens):
    """모든 검사를 실행하고 이슈 목록을 합쳐서 반환한다.
    검사 항목이 늘어나면 이 함수 안에만 추가하면 된다."""
    return (
        check_broken_references(screens)
        + check_duplicate_screens(screens)
        + check_missing_screen_code(screens)
    )


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("input", nargs="?", default="out/screens.json", help="screens.json 경로")
    parser.add_argument("-o", "--output", default="out/issues.csv")
    args = parser.parse_args()

    screens = load_screens(args.input)

    issues = run_all_checks(screens)

    out_path = Path(args.output)
    out_path.parent.mkdir(parents=True, exist_ok=True)
    # utf-8-sig: 엑셀에서 그냥 열어도 한글이 깨지지 않도록 BOM을 붙인다.
    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(issues)

    broken = sum(1 for i in issues if i["issue_type"] == "broken_reference")
    dup = sum(1 for i in issues if i["issue_type"] == "duplicate_screen")
    missing = sum(1 for i in issues if i["issue_type"] == "missing_screen_code")
    print(f"검사 완료: {len(issues)}건 -> {out_path}")
    print(f"  깨진 참조(broken_reference): {broken}건")
    print(f"  중복 화면(duplicate_screen): {dup}건")
    print(f"  화면코드 누락(missing_screen_code): {missing}건")


if __name__ == "__main__":
    main()
