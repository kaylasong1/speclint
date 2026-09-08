"""
screens.json LLM 검사기

out/screens.json을 읽어서 화면 간 이동(transition) 정의 중 이동 대상 화면이
화면 목록에 없는 경우를 LLM에게 찾아달라고 요청하고, 결과를 CSV로 저장한다.

화면이 많으면 프롬프트 한 번에 다 넣기 어려우므로 25개씩 배치로 나눠 여러 번
호출하고, 배치별 결과를 하나의 CSV로 합친다.

사용법:
    python check_llm.py [--input out/screens.json] [-o out/llm_issues.csv]
    python check_llm.py --dry-run          # API 호출 없이 프롬프트만 저장

--input이 out/screens.json이 아니면 출력 파일명에 그 이름이 접미사로 붙는다.
    python check_llm.py --input out/screens_broken.json
        -> out/prompt_broken_1.txt, out/llm_issues_broken.csv
"""

import argparse
import csv
import json
import os
import re
import time
import urllib.error
import urllib.request
from pathlib import Path

FIELDNAMES = ["screen_code", "quote", "target", "reason", "batch"]

PROMPT_TEMPLATE_PATH = "prompts/review_transition.md"
BATCH_SIZE = 25
MODEL = "claude-sonnet-5"
MAX_TOKENS = 4000
ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages"
ANTHROPIC_VERSION = "2023-06-01"

CODE_FENCE_RE = re.compile(r"^```(?:json)?\s*|\s*```$", re.MULTILINE)


def load_screens(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def target_screens(screens):
    """screen_code가 있고 description_blocks가 있는, 실제 검토 대상 화면만 남긴다."""
    return [s for s in screens if s.get("screen_code") and s.get("description_blocks")]


def build_screen_list(screens):
    """문서 전체 화면 목록. 'screen_code | screen_name' 한 줄씩."""
    lines = [f"{s['screen_code']} | {s.get('screen_name', '')}" for s in screens]
    return "\n".join(lines)


def build_descriptions(batch_screens):
    """배치에 속한 화면들의 description만 '## code | name' 헤더와 함께 나열한다."""
    parts = []
    for s in batch_screens:
        header = f"## {s['screen_code']} | {s.get('screen_name', '')}"
        body = "\n".join(b.get("text", "") for b in s.get("description_blocks", []))
        parts.append(f"{header}\n{body}")
    return "\n\n".join(parts)


def build_prompt(template, screen_list_text, descriptions_text):
    return template.replace("{screen_list}", screen_list_text).replace(
        "{descriptions}", descriptions_text
    )


def chunk(items, size):
    for i in range(0, len(items), size):
        yield items[i : i + size]


def call_llm(prompt, api_key):
    body = json.dumps(
        {
            "model": MODEL,
            "max_tokens": MAX_TOKENS,
            "messages": [{"role": "user", "content": prompt}],
        }
    ).encode("utf-8")

    req = urllib.request.Request(
        ANTHROPIC_API_URL,
        data=body,
        headers={
            "x-api-key": api_key,
            "anthropic-version": ANTHROPIC_VERSION,
            "content-type": "application/json",
        },
        method="POST",
    )

    with urllib.request.urlopen(req) as resp:
        data = json.loads(resp.read().decode("utf-8"))

    return "".join(block.get("text", "") for block in data.get("content", []))


def strip_code_fence(text):
    return CODE_FENCE_RE.sub("", text.strip()).strip()


def input_suffix(input_path):
    """out/screens.json이면 접미사 없음. 그 외에는 파일명에서 접미사를 뽑는다.
    예: screens_broken.json -> 'broken', foo.json -> 'foo'"""
    stem = Path(input_path).stem
    if stem == "screens":
        return ""
    if stem.startswith("screens_"):
        return stem[len("screens_"):]
    return stem


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--input", default="out/screens.json", help="screens.json 경로")
    parser.add_argument("-o", "--output", default=None, help="기본값: out/llm_issues.csv (--input 이름에 따라 접미사 자동 부여)")
    parser.add_argument("--dry-run", action="store_true", help="API 호출 없이 프롬프트만 저장")
    args = parser.parse_args()

    if not Path(args.input).exists():
        print(f"에러: 입력 파일을 찾을 수 없습니다: {args.input}")
        return

    suffix = input_suffix(args.input)
    suffix_part = f"_{suffix}" if suffix else ""

    if args.output is None:
        args.output = f"out/llm_issues{suffix_part}.csv"

    out_path = Path(args.output)
    out_dir = out_path.parent
    out_dir.mkdir(parents=True, exist_ok=True)

    api_key = os.environ.get("ANTHROPIC_API_KEY")
    if not args.dry_run and not api_key:
        print("에러: 환경변수 ANTHROPIC_API_KEY가 설정되어 있지 않습니다.")
        return

    template = Path(PROMPT_TEMPLATE_PATH).read_text(encoding="utf-8")

    screens = load_screens(args.input)
    targets = target_screens(screens)
    screen_list_text = build_screen_list(targets)

    batches = list(chunk(targets, BATCH_SIZE))
    total_batches = len(batches)

    all_issues = []
    for batch_num, batch_screens in enumerate(batches, start=1):
        print(f"배치 {batch_num}/{total_batches} 처리 중... (화면 {len(batch_screens)}개)")

        descriptions_text = build_descriptions(batch_screens)
        prompt = build_prompt(template, screen_list_text, descriptions_text)

        if args.dry_run:
            prompt_path = out_dir / f"prompt{suffix_part}_{batch_num}.txt"
            prompt_path.write_text(prompt, encoding="utf-8")
            print(f"  -> {prompt_path} 저장")
            continue

        try:
            raw_text = call_llm(prompt, api_key)
        except (urllib.error.URLError, urllib.error.HTTPError) as e:
            print(f"  경고: 배치 {batch_num} API 호출 실패 ({e}), 다음 배치로 진행합니다.")
            if batch_num < total_batches:
                time.sleep(1)
            continue

        try:
            issues = json.loads(strip_code_fence(raw_text))
        except json.JSONDecodeError:
            raw_path = out_dir / f"llm_raw{suffix_part}_{batch_num}.txt"
            raw_path.write_text(raw_text, encoding="utf-8")
            print(f"  경고: 배치 {batch_num} 응답 파싱 실패, 원문을 {raw_path}에 저장했습니다.")
            if batch_num < total_batches:
                time.sleep(1)
            continue

        for issue in issues:
            issue["batch"] = batch_num
            all_issues.append(issue)
        print(f"  -> {len(issues)}건 검출")

        if batch_num < total_batches:
            time.sleep(1)

    if args.dry_run:
        print(f"dry-run 완료: 프롬프트 {total_batches}개를 {out_dir}에 저장했습니다.")
        return

    with out_path.open("w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=FIELDNAMES)
        writer.writeheader()
        writer.writerows(all_issues)

    print(f"검사 완료: {len(all_issues)}건 -> {out_path}")


if __name__ == "__main__":
    main()
