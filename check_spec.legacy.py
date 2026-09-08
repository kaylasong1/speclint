# check_spec.py
# 기획 문서(SOT JSON)를 읽어서 사람이 눈으로 찾기 힘든 오류를 대신 찾아준다.
# 사용법:  python check_spec.py 파일이름.json
#
# 검사 1) 고아 기능   - 정의만 되고 어느 화면에도 안 쓰이는 기능
# 검사 2) 깨진 참조   - 화면이 가리키는데 정의가 없는 기능
# 검사 3) 도달 불가   - 시작 화면에서 출발해 갈 수 없는 페이지

import json
import sys


def collect_page_refs(pages, found):
    """IA 페이지들을 훑으면서 refs에 적힌 기능 ID를 모은다.
    페이지 안에 children(하위 페이지)이 또 있어서 자기 자신을 다시 호출한다."""
    for page in pages:
        for ref in page.get("refs", []):
            # refs는 "F1" 또는 "F1:0"(기능:수용조건 번호) 두 가지 모양이다.
            # ":" 앞부분만 잘라서 기능 ID만 남긴다.
            found.add(ref.split(":")[0])
        collect_page_refs(page.get("children", []), found)


def collect_page_ids(pages, found):
    """IA에 정의된 모든 페이지 ID와 제목을 모은다. 구조는 위 함수와 똑같다."""
    for page in pages:
        found[page["id"]] = page["title"]
        collect_page_ids(page.get("children", []), found)


def find_reachable(start, transitions):
    """시작 페이지에서 화살표를 따라가며 갈 수 있는 페이지를 전부 찾는다.

    지하철 노선도에서 출발역부터 손가락으로 짚어가는 것과 같다.
    더 이상 새로 갈 곳이 없을 때까지 반복한다."""
    reachable = {start}      # 이미 도착한 역
    to_visit = [start]       # 아직 출구를 안 살펴본 역

    while to_visit:
        current = to_visit.pop()
        # 지금 역에서 나가는 화살표를 모두 본다
        for t in transitions:
            if t["from"] == current and t["to"] not in reachable:
                reachable.add(t["to"])
                to_visit.append(t["to"])   # 이 역도 나중에 출구를 살펴봐야 한다

    return reachable


def main():
    # 1) 파일 읽기 ------------------------------------------------------
    filename = sys.argv[1]
    with open(filename, encoding="utf-8") as f:
        doc = json.load(f)

    all_pages_list = [p for section in doc["ia"]["sections"] for p in section["pages"]]

    # 2) 문서에 "정의된" 기능 ID를 모두 모은다 (F1, F2, ... )
    defined = set()
    for req in doc["requirements"]:
        for feature in req["features"]:
            defined.add(feature["id"])

    # 3) 문서에서 "참조된" 기능 ID를 모두 모은다
    used = set()
    collect_page_refs(all_pages_list, used)
    for transition in doc["flow"]["transitions"]:
        if transition.get("ref"):
            used.add(transition["ref"].split(":")[0])

    # 4) 기능 ID 두 집합을 비교한다 -------------------------------------
    orphans = defined - used   # 정의는 했는데 아무 화면에서도 안 쓰는 기능
    broken = used - defined    # 화면이 가리키는데 정의가 없는 기능

    # 5) 페이지 도달 가능 여부를 본다 -----------------------------------
    pages = {}                                    # {페이지ID: 제목}
    collect_page_ids(all_pages_list, pages)

    start = doc["flow"]["start"]
    reachable = find_reachable(start, doc["flow"]["transitions"])
    unreachable = set(pages) - reachable          # 정의됨 - 도달가능 = 갈 수 없는 페이지

    # 6) 결과 출력 ------------------------------------------------------
    print(f"\n문서: {doc['title']}")
    print(f"정의된 기능 {len(defined)}개 / 참조된 기능 {len(used)}개")
    print(f"정의된 페이지 {len(pages)}개 / 시작({start})에서 도달 가능 {len(reachable)}개\n")

    if orphans:
        print(f"[고아 기능] {len(orphans)}개 — 정의만 되고 어느 화면에도 없습니다")
        for fid in sorted(orphans):
            print(f"   - {fid}")
    else:
        print("[고아 기능] 없음")

    if broken:
        print(f"\n[깨진 참조] {len(broken)}개 — 없는 기능을 가리키고 있습니다")
        for fid in sorted(broken):
            print(f"   - {fid}")
    else:
        print("[깨진 참조] 없음")

    if unreachable:
        print(f"\n[도달 불가] {len(unreachable)}개 — 들어갈 경로가 없는 페이지입니다")
        for pid in sorted(unreachable):
            print(f"   - {pid} {pages[pid]}")
    else:
        print("\n[도달 불가] 없음")

    print()


main()
