# -*- coding: utf-8 -*-
"""Attach Ukrainian product URLs without downloading or reprocessing images."""
from __future__ import annotations

import argparse
import json
import re
import urllib.parse
from collections import defaultdict
from pathlib import Path

from scan_catalog_colors import BASE, DETAIL_RE, CachedFetcher, absolute_url, discover_products


def image_key(url: str) -> str:
    path = urllib.parse.unquote(urllib.parse.urlsplit(url).path).lower()
    return path.replace("/resized/", "/")


def product_id(url: str) -> str | None:
    name = urllib.parse.unquote(urllib.parse.urlsplit(url).path.rsplit("/", 1)[-1])
    match = re.search(r"-(\d+)-detail\.html$", name, re.I)
    return match.group(1) if match else None


def locale_neutral_url(url: str) -> str:
    parts = urllib.parse.urlsplit(url)
    path = parts.path[3:] if parts.path.startswith("/uk/") else parts.path
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, path, parts.query, ""))


def ukrainian_switch_url(page_html: str, page_url: str) -> str | None:
    for tag in re.findall(r"<a\b[^>]*>", page_html, re.I):
        if not re.search(r"aria-label=[\"']Українська[\"']", tag, re.I):
            continue
        href = re.search(r"href=[\"']([^\"']+)[\"']", tag, re.I)
        if href:
            url = absolute_url(href.group(1), page_url)
            if url.startswith(BASE + "/uk/") and DETAIL_RE.search(url):
                return url
    return None


def choose_match(row: dict, candidates: list[dict]) -> dict | None:
    unique = {item["product_url"]: item for item in candidates}
    candidates = list(unique.values())
    if not candidates:
        return None
    if len(candidates) == 1:
        return candidates[0]

    direct = BASE + "/uk" + urllib.parse.urlsplit(row["product_url"]).path
    for candidate in candidates:
        if urllib.parse.urlsplit(candidate["product_url"]).path == urllib.parse.urlsplit(direct).path:
            return candidate

    wanted_id = product_id(row["product_url"])
    if wanted_id:
        by_id = [item for item in candidates if product_id(item["product_url"]) == wanted_id]
        if len(by_id) == 1:
            return by_id[0]

    ru_dir = urllib.parse.unquote(urllib.parse.urlsplit(row["product_url"]).path).rsplit("/", 1)[0]
    by_dir = [item for item in candidates
              if urllib.parse.unquote(urllib.parse.urlsplit(locale_neutral_url(item["product_url"])).path).rsplit("/", 1)[0] == ru_dir]
    return by_dir[0] if len(by_dir) == 1 else None


def main() -> None:
    parser = argparse.ArgumentParser()
    parser.add_argument("catalog")
    parser.add_argument("--work-dir", default=".cache/catalog_scan")
    parser.add_argument("--pause", type=float, default=1.5)
    args = parser.parse_args()

    path = Path(args.catalog)
    rows = json.loads(path.read_text(encoding="utf-8"))
    fetcher = CachedFetcher(Path(args.work_dir) / "uk-pages", args.pause)
    for row in rows:
        url = row.get("product_url_uk") or ""
        if url and not (url.startswith(BASE + "/uk/") and DETAIL_RE.search(url)):
            row["product_url_uk"] = None
            row["title_uk"] = None
    pending = [row for row in rows if not row.get("product_url_uk")]
    uk_products = discover_products(fetcher, language_prefix="/uk") if len(pending) >= 100 else []

    by_image = defaultdict(list)
    for product in uk_products:
        for image_url in product["images"]:
            by_image[image_key(image_url)].append(product)

    matched = sum(bool(row.get("product_url_uk")) for row in rows)
    ambiguous = 0
    for row in pending:
        candidates = by_image.get(image_key(row.get("image_url", "")), [])
        match = choose_match(row, candidates)
        if match:
            row["product_url_uk"] = match["product_url"]
            row["title_uk"] = match["title"]
            matched += 1
        else:
            if candidates:
                ambiguous += 1

    switch_matched = 0
    for row in rows:
        if row.get("product_url_uk"):
            continue
        try:
            url = ukrainian_switch_url(fetcher.text(row["product_url"]), row["product_url"])
        except Exception as exc:
            print("language switch error:", row["product_url"], exc)
            continue
        if url:
            row["product_url_uk"] = url
            row.setdefault("title_uk", None)
            matched += 1
            switch_matched += 1

    temporary = path.with_suffix(path.suffix + ".tmp")
    temporary.write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    temporary.replace(path)
    print("rows %d, ukrainian links %d, unmatched %d, ambiguous %d, switch links %d" %
          (len(rows), matched, len(rows) - matched, ambiguous, switch_matched))


if __name__ == "__main__":
    main()
