#!/usr/bin/env python3
"""K-MOOC 강좌 크롤러 — K-MOOC 시트 양식으로 엑셀 출력."""

from __future__ import annotations

import argparse
import json
import logging
import re
import time
from concurrent.futures import ThreadPoolExecutor, as_completed
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path

import openpyxl
import requests
from bs4 import BeautifulSoup

BASE_URL = "https://www.kmooc.kr"
SEARCH_API = f"{BASE_URL}/json/course/category/search"
DETAIL_URL = f"{BASE_URL}/view/course/detail/{{cid}}"
HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) "
        "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "ko-KR,ko;q=0.9",
    "X-Requested-With": "XMLHttpRequest",
    "Referer": f"{BASE_URL}/view/course",
}
EXCEL_COLUMNS = (
    "순번",
    "출처(제작사)",
    "대메뉴",
    "중분류",
    "세분류(전공분류)",
    "상태",
    "콘텐츠명",
    "강좌수",
    "강의목차",
    "강사명",
    "URL",
    "비고",
)
SHEET_NAME = "K-MOOC"
PROVIDER = "K-MOOC"
# API open_status id → 표시명
OPEN_STATUSES = (
    ("1", "진행중"),
    ("2", "개강예정"),
    ("4", "종료(청강)"),
    ("3", "종료"),
)

log = logging.getLogger(__name__)


@dataclass
class CourseStub:
    cid: str
    title: str
    professor: str
    status: str
    url: str
    section_hint: int = 0


@dataclass
class CourseRow:
    provider: str
    menu: str | None
    mid_category: str
    sub_category: str
    status: str
    title: str
    lecture_count: int
    syllabus: str
    professor: str
    url: str
    detail_fetched: bool = True


class KmoocClient:
    def __init__(self, delay: float = 0.3, retries: int = 3):
        self.session = requests.Session()
        self.session.headers.update(HEADERS)
        self.delay = delay
        self.retries = retries

    def get(self, url: str) -> str:
        last_error: Exception | None = None
        for attempt in range(1, self.retries + 1):
            try:
                time.sleep(self.delay)
                resp = self.session.get(url, timeout=30)
                resp.raise_for_status()
                resp.encoding = "utf-8"
                return resp.text
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                log.warning("요청 실패 (%s/%s): %s — %s", attempt, self.retries, url, exc)
                time.sleep(attempt)
        raise RuntimeError(f"요청 실패: {url}") from last_error

    def post_json(self, url: str, data: dict) -> dict:
        last_error: Exception | None = None
        for attempt in range(1, self.retries + 1):
            try:
                time.sleep(self.delay)
                resp = self.session.post(url, data=data, timeout=30)
                resp.raise_for_status()
                return resp.json()
            except Exception as exc:  # noqa: BLE001
                last_error = exc
                log.warning("요청 실패 (%s/%s): %s — %s", attempt, self.retries, url, exc)
                time.sleep(attempt)
        raise RuntimeError(f"요청 실패: {url}") from last_error


def parse_categories(soup: BeautifulSoup) -> tuple[str, str]:
    crumbs = [a.get_text(strip=True) for a in soup.select(".breadcrumb a")]
    if "강좌" in crumbs:
        i = crumbs.index("강좌")
        mid = crumbs[i + 1] if len(crumbs) > i + 1 else ""
        sub = crumbs[i + 2] if len(crumbs) > i + 2 else ""
        if mid and mid not in {"공유하기", "홈"}:
            return mid, sub if sub not in {"공유하기"} else ""

    text = soup.get_text("\n", strip=True)
    m = re.search(r"분야\s*\n?\s*([^\n(]+)\s*\(([^)]+)\)", text)
    if m:
        return m.group(1).strip(), m.group(2).strip()
    return "", ""


def parse_status(soup: BeautifulSoup, fallback: str = "") -> str:
    for el in soup.select("small.audit, small.ing, small.prev, small.end"):
        parent = el.parent.get_text(" ", strip=True) if el.parent else ""
        status = el.get_text(strip=True)
        if status in {"진행중", "개강예정", "종료(청강)", "종료", "개강중"} and status in parent:
            return status
    return fallback


def parse_syllabus(soup: BeautifulSoup) -> tuple[int, str]:
    items: list[str] = []
    for table in soup.select("table"):
        head = table.get_text(" ", strip=True)
        if "주차" not in head or "강의명" not in head:
            continue
        for tr in table.select("tr")[1:]:
            tds = tr.select("td")
            if not tds:
                continue
            title = tds[-1].get_text(strip=True)
            if title and not re.match(r"^\d+주차$", title):
                items.append(title)
        break
    syllabus = "\n".join(f"{i}.\t{title}" for i, title in enumerate(items, 1))
    return len(items), syllabus


def parse_professor(soup: BeautifulSoup, fallback: str = "") -> str:
    for h in soup.find_all(["h2", "h3", "strong"]):
        if h.get_text(strip=True) == "담당 교수":
            nxt = h.find_next(["h2", "h3", "p", "strong", "span"])
            if nxt:
                name = nxt.get_text(strip=True)
                if name and name != "담당 교수":
                    return name
    return fallback


def parse_course_detail(html: str, stub: CourseStub) -> CourseRow:
    soup = BeautifulSoup(html, "lxml")
    mid, sub = parse_categories(soup)
    status = parse_status(soup, stub.status)
    lecture_count, syllabus = parse_syllabus(soup)
    if lecture_count == 0 and stub.section_hint:
        lecture_count = stub.section_hint
    professor = parse_professor(soup, stub.professor)

    return CourseRow(
        provider=PROVIDER,
        menu=None,
        mid_category=mid,
        sub_category=sub,
        status=status,
        title=stub.title,
        lecture_count=lecture_count,
        syllabus=syllabus,
        professor=professor or stub.professor,
        url=stub.url,
        detail_fetched=True,
    )


def stub_to_row(stub: CourseStub) -> CourseRow:
    return CourseRow(
        provider=PROVIDER,
        menu=None,
        mid_category="",
        sub_category="",
        status=stub.status,
        title=stub.title,
        lecture_count=stub.section_hint,
        syllabus="",
        professor=stub.professor,
        url=stub.url,
        detail_fetched=False,
    )


def cache_entry_fetched(entry: dict) -> bool:
    return entry.get("detail_fetched", True)


def row_from_cache(entry: dict) -> CourseRow:
    data = dict(entry)
    data.setdefault("detail_fetched", True)
    data.setdefault("menu", None)
    return CourseRow(**{k: data[k] for k in CourseRow.__dataclass_fields__})


def load_cache(path: Path) -> dict[str, dict]:
    if not path.exists():
        return {}
    with path.open(encoding="utf-8") as f:
        return json.load(f)


def save_cache(path: Path, cache: dict[str, dict]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open("w", encoding="utf-8") as f:
        json.dump(cache, f, ensure_ascii=False, indent=2)


def save_stubs(path: Path, stubs: list[CourseStub]) -> None:
    path.parent.mkdir(parents=True, exist_ok=True)
    payload = {
        "collected_at": datetime.now().isoformat(timespec="seconds"),
        "stubs": [asdict(s) for s in stubs],
    }
    with path.open("w", encoding="utf-8") as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)


def load_stubs(path: Path) -> list[CourseStub] | None:
    if not path.exists():
        return None
    with path.open(encoding="utf-8") as f:
        payload = json.load(f)
    stubs = [CourseStub(**item) for item in payload.get("stubs", [])]
    if not stubs:
        return None
    log.info(
        "목록 캐시 사용: %s건 (%s, 수집일 %s)",
        len(stubs),
        path,
        payload.get("collected_at", "알 수 없음"),
    )
    return stubs


def collect_stubs(client: KmoocClient, limit: int = 0) -> list[CourseStub]:
    stubs: list[CourseStub] = []
    seen: set[str] = set()

    for status_id, status_name in OPEN_STATUSES:
        page = 1
        page_total = 1
        while page <= page_total:
            data = client.post_json(
                SEARCH_API,
                {
                    "page": page,
                    "open_status": status_id,
                    "sby": "enrol_end",
                    "sorder": "desc",
                },
            )
            if data.get("result") != "success":
                raise RuntimeError(f"목록 API 실패: {data}")

            page_total = int(data.get("paging", {}).get("pageTotal") or 1)
            if page == 1 or page % 20 == 0 or page == page_total:
                log.info(
                    "목록: %s %s/%s페이지 (누적 %s건)",
                    status_name,
                    page,
                    page_total,
                    len(stubs),
                )

            for item in data.get("list") or []:
                cid = str(item.get("id") or "")
                if not cid or cid in seen:
                    continue
                seen.add(cid)
                profs = item.get("profs") or []
                professor = ", ".join(p for p in profs if p)
                stubs.append(
                    CourseStub(
                        cid=cid,
                        title=(item.get("fullname") or "").strip(),
                        professor=professor,
                        status=status_name,
                        url=DETAIL_URL.format(cid=cid),
                        section_hint=int(item.get("numsections") or 0),
                    )
                )
                if limit > 0 and len(stubs) >= limit:
                    return stubs
            page += 1

    return stubs


def get_stubs(
    client: KmoocClient,
    stubs_path: Path,
    limit: int,
    refresh_listing: bool,
) -> list[CourseStub]:
    if limit == 0 and not refresh_listing:
        cached = load_stubs(stubs_path)
        if cached:
            return cached

    stubs = collect_stubs(client, limit=limit)
    if limit == 0:
        save_stubs(stubs_path, stubs)
        log.info("목록 캐시 저장: %s", stubs_path)
    return stubs


def fetch_details(
    client: KmoocClient,
    stubs: list[CourseStub],
    cache_path: Path,
    workers: int,
) -> list[CourseRow]:
    cache = load_cache(cache_path)
    pending: list[CourseStub] = []

    for stub in stubs:
        entry = cache.get(stub.cid)
        if entry is None or not cache_entry_fetched(entry):
            pending.append(stub)

    log.info("상세 페이지: 캐시 %s건, 신규 %s건", len(stubs) - len(pending), len(pending))

    def fetch_one(stub: CourseStub) -> tuple[str, CourseRow, bool]:
        try:
            html = client.get(DETAIL_URL.format(cid=stub.cid))
            return stub.cid, parse_course_detail(html, stub), True
        except RuntimeError:
            log.error("상세 수집 실패(목록 정보만 사용): %s", stub.url)
            return stub.cid, stub_to_row(stub), False

    failed = 0
    if pending:
        with ThreadPoolExecutor(max_workers=workers) as executor:
            futures = {executor.submit(fetch_one, stub): stub for stub in pending}
            done = 0
            for future in as_completed(futures):
                cid, row, ok = future.result()
                if ok:
                    cache[cid] = asdict(row)
                else:
                    failed += 1
                done += 1
                if done % 50 == 0:
                    save_cache(cache_path, cache)
                    log.info("상세 수집 진행: %s/%s", done, len(pending))
        save_cache(cache_path, cache)

    if failed:
        log.warning("상세 수집 실패 %s건 — 엑셀에는 기본 정보만 기록, 재실행 시 재시도", failed)

    rows: list[CourseRow] = []
    for stub in stubs:
        entry = cache.get(stub.cid)
        if entry and cache_entry_fetched(entry):
            rows.append(row_from_cache(entry))
        else:
            rows.append(stub_to_row(stub))
    return rows


def write_excel(output: Path, rows: list[CourseRow]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = SHEET_NAME
    ws.append(list(EXCEL_COLUMNS))

    for idx, row in enumerate(rows, start=1):
        ws.append(
            [
                idx,
                row.provider,
                row.menu,
                row.mid_category,
                row.sub_category,
                row.status,
                row.title,
                row.lecture_count,
                row.syllabus,
                row.professor,
                row.url,
                None if row.detail_fetched else "상세수집실패",
            ]
        )

    wb.save(output)
    wb.close()


def main() -> None:
    parser = argparse.ArgumentParser(description="K-MOOC 강좌 크롤러")
    parser.add_argument(
        "--output",
        default=None,
        help="결과 엑셀 파일 경로 (기본: KMOOC_크롤링결과_날짜.xlsx)",
    )
    parser.add_argument("--limit", type=int, default=0, help="테스트용 수집 건수 제한")
    parser.add_argument("--delay", type=float, default=0.3, help="요청 간격(초)")
    parser.add_argument("--workers", type=int, default=5, help="상세 페이지 병렬 수")
    parser.add_argument(
        "--cache",
        default=".cache/kmooc_details.json",
        help="상세 페이지 캐시 파일",
    )
    parser.add_argument(
        "--stubs-cache",
        default=".cache/kmooc_stubs.json",
        help="강좌 목록 캐시 파일",
    )
    parser.add_argument(
        "--refresh-listing",
        action="store_true",
        help="저장된 목록 캐시를 무시하고 목록을 다시 수집",
    )
    parser.add_argument("-v", "--verbose", action="store_true")
    args = parser.parse_args()

    logging.basicConfig(
        level=logging.DEBUG if args.verbose else logging.INFO,
        format="%(asctime)s %(levelname)s %(message)s",
    )

    output = Path(
        args.output
        or f"KMOOC_크롤링결과_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
    )
    client = KmoocClient(delay=args.delay)
    stubs = get_stubs(
        client,
        Path(args.stubs_cache),
        limit=args.limit,
        refresh_listing=args.refresh_listing,
    )
    log.info("목록 준비 완료: %s건", len(stubs))

    rows = fetch_details(client, stubs, Path(args.cache), args.workers)
    write_excel(output, rows)
    log.info("저장 완료: %s (%s건)", output, len(rows))


if __name__ == "__main__":
    main()
