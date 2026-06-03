from __future__ import annotations

import html
import json
import os
import re
import urllib.parse
import urllib.request
from concurrent.futures import ThreadPoolExecutor, as_completed
from pathlib import Path

from openpyxl import load_workbook


ROOT = Path(__file__).resolve().parents[1]
WORKBOOK = next(ROOT.glob("*.xlsx"))
OUTPUT = ROOT / "assets" / "app-data.js"
CACHE = ROOT / "assets" / "bilibili-cache.json"


def clean(value) -> str:
    if value is None:
        return ""
    text = str(value).replace("\n", " ").replace("\r", " ").strip()
    return re.sub(r"\s+", " ", text)


def clean_major(value) -> str:
    text = clean(value)
    text = text.replace("\uff08", "(").replace("\uff09", ")")
    text = re.sub(r"\s+", "", text)
    return text.replace("(", "\uff08").replace(")", "\uff09")


def strip_highlight(value: str) -> str:
    text = re.sub(r"</?em[^>]*>", "", value or "")
    return clean(html.unescape(text))


def is_real_book(title: str) -> bool:
    if not title:
        return False
    skip_words = ["\u65e0\u9700\u6559\u6750", "\u7528\u4e0a\u5b66\u671f\u7684\u6559\u6750"]
    return not any(word in title for word in skip_words)


def search_bilibili(query: str) -> dict | None:
    keyword = urllib.parse.quote(query)
    url = (
        "https://api.bilibili.com/x/web-interface/search/type"
        f"?search_type=video&keyword={keyword}"
    )
    req = urllib.request.Request(
        url,
        headers={
            "User-Agent": (
                "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
                "AppleWebKit/537.36 (KHTML, like Gecko) Chrome/125 Safari/537.36"
            ),
            "Referer": "https://search.bilibili.com/",
        },
    )
    try:
        with urllib.request.urlopen(req, timeout=20) as response:
            payload = json.loads(response.read().decode("utf-8", errors="ignore"))
    except Exception:
        return None

    for item in payload.get("data", {}).get("result", []):
        if item.get("type") != "video" or not item.get("bvid"):
            continue
        title = strip_highlight(item.get("title", ""))
        if not title:
            continue
        return {
            "title": title,
            "author": clean(item.get("author", "")),
            "bvid": clean(item.get("bvid", "")),
            "url": f"https://www.bilibili.com/video/{item.get('bvid')}",
            "play": int(item.get("play") or 0),
            "duration": clean(item.get("duration", "")),
        }
    return None


def load_rows() -> list[dict]:
    workbook = load_workbook(WORKBOOK, data_only=True)
    rows = []
    seen_rows = set()

    for sheet in workbook.worksheets:
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if not row or len(row) < 4:
                continue
            major = clean_major(row[0])
            year = clean(row[1])
            course = clean(row[2])
            title = clean(row[3])
            author = clean(row[4] if len(row) > 4 else "")
            publisher = clean(row[5] if len(row) > 5 else "")
            if not (major and year and course and is_real_book(title)):
                continue
            key = (major, year, course, title, author, publisher)
            if key in seen_rows:
                continue
            seen_rows.add(key)
            rows.append(
                {
                    "majorName": major,
                    "year": year,
                    "course": course,
                    "title": title,
                    "author": author,
                    "publisher": publisher,
                }
            )
    return rows


def load_cache() -> dict:
    if not CACHE.exists():
        return {}
    try:
        return json.loads(CACHE.read_text(encoding="utf-8"))
    except json.JSONDecodeError:
        return {}


def write_cache(cache: dict) -> None:
    CACHE.write_text(json.dumps(cache, ensure_ascii=False, indent=2), encoding="utf-8")


def main() -> None:
    rows = load_rows()

    major_order = []
    for row in rows:
        if row["majorName"] not in major_order:
            major_order.append(row["majorName"])

    majors = []
    major_ids = {}
    for index, name in enumerate(major_order, start=1):
        major_rows = [row for row in rows if row["majorName"] == name]
        year_count = len({row["year"] for row in major_rows})
        book_count = len(major_rows)
        major_id = f"major-{index}"
        major_ids[name] = major_id
        majors.append(
            {
                "id": major_id,
                "name": name,
                "desc": f"{year_count} \u4e2a\u5e74\u7ea7 \u00b7 {book_count} \u672c\u6559\u6750",
            }
        )

    cache = load_cache()
    unique_queries = []
    for row in rows:
        query = f"{row['title']} {row['course']}".strip()
        if query not in unique_queries:
            unique_queries.append(query)

    pending = [] if os.environ.get("SKIP_BILI_SEARCH") else [
        query for query in unique_queries if query not in cache
    ]
    with ThreadPoolExecutor(max_workers=8) as executor:
        futures = {executor.submit(search_bilibili, query): query for query in pending}
        for index, future in enumerate(as_completed(futures), start=1):
            query = futures[future]
            try:
                cache[query] = future.result()
            except Exception:
                cache[query] = None
            if index % 40 == 0:
                write_cache(cache)
    write_cache(cache)

    books = []
    for index, row in enumerate(rows, start=1):
        query = f"{row['title']} {row['course']}".strip()
        video = cache.get(query)
        routes = []
        if video:
            routes.append(
                {
                    "title": "B\u7ad9\u76f8\u5173\u89c6\u9891",
                    "search": (
                        f"\u89c6\u9891\uff1a{video['title']}\uff5cUP\uff1a{video['author']}\uff5c"
                        f"{video['bvid']}\uff5c\u64ad\u653e\uff1a{video['play']}"
                    ),
                    "url": video["url"],
                    "query": query,
                }
            )
        routes.append(
            {
                "title": "B\u7ad9\u641c\u7d22\u5173\u952e\u8bcd",
                "search": f"\u5efa\u8bae\u641c\u7d22\uff1a{query}",
                "query": query,
            }
        )

        tags = [value for value in [row["author"], row["publisher"]] if value][:2]
        note_parts = []
        if row["author"]:
            note_parts.append(f"\u4f5c\u8005\uff1a{row['author']}")
        if row["publisher"]:
            note_parts.append(f"\u51fa\u7248\u793e\uff1a{row['publisher']}")
        note_parts.append("\u6309 Excel \u6559\u6750\u8868\u6c47\u603b\u5c55\u793a\u3002")

        books.append(
            {
                "id": f"book-{index}",
                "major": major_ids[row["majorName"]],
                "year": row["year"],
                "course": row["course"],
                "title": row["title"],
                "tags": tags,
                "note": "\uff1b".join(note_parts),
                "routes": routes,
            }
        )

    payload = {"majors": majors, "books": books}
    OUTPUT.write_text(
        "window.APP_DATA = "
        + json.dumps(payload, ensure_ascii=False, indent=2)
        + ";\n",
        encoding="utf-8",
    )
    print(f"majors={len(majors)} books={len(books)} unique_queries={len(unique_queries)}")


if __name__ == "__main__":
    main()
