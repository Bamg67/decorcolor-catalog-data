# -*- coding: utf-8 -*-
"""Build a product-photo color index for decor-opt.com.ua.

The crawler is intentionally conservative: it reads the public sitemap, expands
catalog/category pages, caches every response, and sleeps between network calls.
Output is JSON so a separate presentation step can create CSV/XLSX without
mixing crawling, image analysis, and workbook formatting.
"""
from __future__ import annotations

import argparse
import colorsys
import hashlib
import html as html_lib
import json
import math
import os
import re
import sys
import time
import urllib.parse
import urllib.request
import urllib.error
import xml.etree.ElementTree as ET
from pathlib import Path

import numpy as np
from PIL import Image, ImageOps


BASE = "https://decor-opt.com.ua"
SITEMAP = BASE + "/sitemap.xml"
UA = {"User-Agent": "Mozilla/5.0 (DecorColor product palette; owner-authorized)"}
DETAIL_RE = re.compile(r"(?:-detail\.html|/details\.html)(?:[?#].*)?$", re.I)
IMAGE_RE = re.compile(r"(?:https?:)?(?://decor-opt\.com\.ua)?(/images/stories/virtuemart/product/[^\"'<>]+)", re.I)


def clean_text(value: str) -> str:
    value = re.sub(r"<[^>]+>", " ", value)
    return re.sub(r"\s+", " ", html_lib.unescape(value)).strip()


def absolute_url(value: str, base: str = BASE) -> str:
    value = html_lib.unescape(value).replace("\\/", "/")
    if value.startswith("//"):
        value = "https:" + value
    joined = urllib.parse.urljoin(base, value)
    parts = urllib.parse.urlsplit(joined)
    path = urllib.parse.quote(urllib.parse.unquote(parts.path), safe="/%:@,;=-._~!$&'()*+")
    query = urllib.parse.quote(urllib.parse.unquote(parts.query), safe="=&?/:@,;+-._~!$'()*")
    return urllib.parse.urlunsplit((parts.scheme, parts.netloc, path, query, parts.fragment))


class CachedFetcher:
    def __init__(self, cache_dir: Path, pause: float = 0.35):
        self.cache_dir = cache_dir
        self.pause = pause
        self.cache_dir.mkdir(parents=True, exist_ok=True)

    def _path(self, url: str, suffix: str) -> Path:
        digest = hashlib.sha256(url.encode("utf-8")).hexdigest()
        return self.cache_dir / (digest + suffix)

    def bytes(self, url: str, suffix: str = ".bin") -> bytes:
        path = self._path(url, suffix)
        if path.exists() and path.stat().st_size:
            return path.read_bytes()
        req = urllib.request.Request(url, headers=UA)
        last_error = None
        for attempt in range(3):
            try:
                with urllib.request.urlopen(req, timeout=60) as response:
                    body = response.read()
                path.write_bytes(body)
                time.sleep(self.pause)
                return body
            except urllib.error.HTTPError as exc:
                if exc.code in (400, 401, 403, 404, 410):
                    raise RuntimeError("HTTP %d for %s" % (exc.code, url)) from exc
                last_error = exc
            except Exception as exc:  # network errors are recorded per product
                last_error = exc
                time.sleep(1.5 * (attempt + 1))
        raise RuntimeError("download failed for %s: %s" % (url, last_error))

    def text(self, url: str) -> str:
        return self.bytes(url, ".html").decode("utf-8", "replace")


def sitemap_urls(fetcher: CachedFetcher) -> list[str]:
    root = ET.fromstring(fetcher.bytes(SITEMAP, ".xml"))
    return [node.text.strip() for node in root.iter() if node.tag.endswith("loc") and node.text]


def detail_links(page_html: str, page_url: str) -> set[str]:
    links = set()
    for href in re.findall(r"href\s*=\s*[\"']([^\"']+)[\"']", page_html, re.I):
        url = absolute_url(href, page_url).split("#", 1)[0]
        if url.startswith(BASE + "/") and DETAIL_RE.search(url):
            links.add(url)
    return links


def catalog_page_candidates(urls: list[str]) -> list[str]:
    skip = ("registration", "reset", "remind", "korzina", "kontakty", "kak-sdelat-zakaz")
    out = []
    for url in urls:
        low = url.lower()
        if not low.startswith(BASE + "/novinki") or DETAIL_RE.search(low):
            continue
        if any(part in low for part in skip):
            continue
        out.append(url)
    return sorted(set(out))


def parse_product_cards(page_html: str, page_url: str) -> list[dict]:
    products = []
    for block in page_html.split("product-container")[1:]:
        block = block[:12000]
        href = re.search(r'href\s*=\s*[\"\']([^\"\']+)[\"\']', block, re.I)
        image = re.search(r'(?:data-src|src)\s*=\s*[\"\']([^\"\']*?/images/stories/virtuemart/product/[^\"\']+)[\"\']', block, re.I)
        title = re.search(r'vm-product-title[^>]*>(.*?)</h\d>', block, re.I | re.S)
        if not (href and image):
            continue
        url = absolute_url(href.group(1), page_url).split("#", 1)[0]
        image_url = absolute_url(image.group(1), page_url)
        name = clean_text(title.group(1)) if title else ""
        if not name:
            alt = re.search(r'alt\s*=\s*[\"\']([^\"\']+)', block, re.I)
            name = clean_text(alt.group(1)) if alt else ""
        if url.startswith(BASE + "/") and image_url.startswith(BASE + "/"):
            products.append({"title": name, "sku": "", "product_url": url,
                             "images": [image_url]})
    return products


def pagination_links(page_html: str, page_url: str) -> set[str]:
    links = set()
    for href in re.findall(r'href\s*=\s*[\"\']([^\"\']+)[\"\']', page_html, re.I):
        if "results," not in href.lower():
            continue
        url = absolute_url(href, page_url).split("#", 1)[0]
        if url.startswith(BASE + "/novinki/"):
            links.add(url)
    return links


def discover_products(fetcher: CachedFetcher, max_pages: int | None = None) -> list[dict]:
    urls = sitemap_urls(fetcher)
    products = {}
    pages = catalog_page_candidates(urls)
    if max_pages is not None:
        pages = pages[:max_pages]
    queue = list(pages)
    seen_pages = set()
    while queue:
        page = queue.pop(0)
        if page in seen_pages:
            continue
        seen_pages.add(page)
        try:
            page_html = fetcher.text(page)
            for product in parse_product_cards(page_html, page):
                previous = products.get(product["product_url"])
                if previous and product["images"][0] not in previous["images"]:
                    previous["images"].append(product["images"][0])
                elif not previous:
                    products[product["product_url"]] = product
            for next_page in pagination_links(page_html, page):
                if next_page not in seen_pages:
                    queue.append(next_page)
        except Exception as exc:
            print("category error:", page, exc, file=sys.stderr)
        count = len(seen_pages)
        if count % 25 == 0 or not queue:
            print("catalog pages %d, queued %d, products %d" % (count, len(queue), len(products)))
    return sorted(products.values(), key=lambda item: (item["title"].lower(), item["product_url"]))


def meta_content(page_html: str, key: str) -> str | None:
    patterns = [
        r'<meta[^>]+(?:property|name)=[\"\']%s[\"\'][^>]+content=[\"\']([^\"\']+)' % re.escape(key),
        r'<meta[^>]+content=[\"\']([^\"\']+)[\"\'][^>]+(?:property|name)=[\"\']%s[\"\']' % re.escape(key),
    ]
    for pattern in patterns:
        match = re.search(pattern, page_html, re.I)
        if match:
            return clean_text(match.group(1))
    return None


def parse_product(page_html: str, url: str) -> dict:
    title = meta_content(page_html, "og:title")
    if not title:
        match = re.search(r"<h1[^>]*>(.*?)</h1>", page_html, re.I | re.S)
        title = clean_text(match.group(1)) if match else urllib.parse.unquote(url.rsplit("/", 1)[-1])

    sku = ""
    for pattern in (
        r'itemprop=[\"\']sku[\"\'][^>]*>(.*?)</',
        r'(?:Артикул|SKU)\s*:?\s*</?[^>]*>\s*([^<\r\n]+)',
    ):
        match = re.search(pattern, page_html, re.I | re.S)
        if match:
            sku = clean_text(match.group(1))
            break

    images = []
    og_image = meta_content(page_html, "og:image")
    if og_image:
        images.append(absolute_url(og_image, url))
    for rel in IMAGE_RE.findall(page_html):
        candidate = absolute_url("/" + rel, url)
        candidate = candidate.replace("/resized/", "/")
        if candidate not in images:
            images.append(candidate)

    # Product pages also contain recommendation thumbnails. Keep the explicit
    # social image first and gallery/full-size images belonging to this page.
    if og_image:
        og_name = urllib.parse.unquote(urllib.parse.urlsplit(images[0]).path.rsplit("/", 1)[-1])
        stem = re.sub(r"(?:_\d+x\d+)?\.[^.]+$", "", og_name).lower()
        related = [img for img in images if stem and stem in urllib.parse.unquote(img).lower()]
        images = [images[0]] + [img for img in related if img != images[0]]
    return {"title": title, "sku": sku, "product_url": url, "images": images[:12]}


def srgb_to_lab(rgb: np.ndarray) -> np.ndarray:
    a = rgb.astype(np.float64) / 255.0
    lin = np.where(a <= 0.04045, a / 12.92, ((a + 0.055) / 1.055) ** 2.4)
    xyz = lin @ np.array([[0.4124, 0.3576, 0.1805],
                          [0.2126, 0.7152, 0.0722],
                          [0.0193, 0.1192, 0.9505]]).T
    t = xyz / np.array([0.95047, 1.0, 1.08883])
    f = np.where(t > 0.008856, np.cbrt(t), 7.787 * t + 16.0 / 116.0)
    return np.stack([116 * f[..., 1] - 16,
                     500 * (f[..., 0] - f[..., 1]),
                     200 * (f[..., 1] - f[..., 2])], axis=-1)


def lab_to_rgb(lab: np.ndarray) -> np.ndarray:
    L, A, B = lab
    fy = (L + 16) / 116.0
    fx, fz = fy + A / 500.0, fy - B / 200.0
    inv = lambda t: t ** 3 if t ** 3 > 0.008856 else (t - 16.0 / 116.0) / 7.787
    xyz = np.array([inv(fx) * 0.95047, inv(fy), inv(fz) * 1.08883])
    lin = xyz @ np.array([[3.2406, -1.5372, -0.4986],
                          [-0.9689, 1.8758, 0.0415],
                          [0.0557, -0.2040, 1.0570]]).T
    srgb = np.where(lin <= 0.0031308, lin * 12.92,
                    1.055 * np.power(np.clip(lin, 0, None), 1 / 2.4) - 0.055)
    return np.clip(np.round(srgb * 255), 0, 255).astype(int)


def kmeans(points: np.ndarray, k: int, iters: int = 30) -> tuple[np.ndarray, np.ndarray]:
    rng = np.random.default_rng(67)
    # Quantile-spread seeds are more stable than purely random seeds.
    first = int(rng.integers(len(points)))
    indices = [first]
    distances = ((points - points[first]) ** 2).sum(axis=1)
    while len(indices) < k:
        indices.append(int(np.argmax(distances)))
        distances = np.minimum(distances, ((points - points[indices[-1]]) ** 2).sum(axis=1))
    centers = points[indices].copy()
    labels = np.zeros(len(points), dtype=int)
    for _ in range(iters):
        new_labels = ((points[:, None, :] - centers[None, :, :]) ** 2).sum(-1).argmin(1)
        if np.array_equal(labels, new_labels):
            break
        labels = new_labels
        for j in range(k):
            selected = points[labels == j]
            if len(selected):
                centers[j] = np.median(selected, axis=0)
    return centers, labels


def color_record(lab: np.ndarray, share: float) -> dict:
    rgb = lab_to_rgb(lab)
    hue, saturation, value = colorsys.rgb_to_hsv(*(rgb / 255.0))
    chroma = float(math.hypot(float(lab[1]), float(lab[2])))
    lab_hue = (math.degrees(math.atan2(float(lab[2]), float(lab[1]))) + 360.0) % 360.0
    hue_deg = round(hue * 360.0, 1) if saturation >= 0.08 and chroma >= 6.0 else None
    hue_from = round((hue * 360.0 - 12.0) % 360.0, 1) if hue_deg is not None else None
    hue_to = round((hue * 360.0 + 12.0) % 360.0, 1) if hue_deg is not None else None
    return {
        "hex": "#%02X%02X%02X" % tuple(rgb),
        "hue_deg": hue_deg,
        "hue_from_deg": hue_from,
        "hue_to_deg": hue_to,
        "hue_wraps_zero": bool(hue_from > hue_to) if hue_deg is not None else False,
        "delta_e_likely": 10.0,
        "delta_e_review": 15.0,
        "lightness_from": round(max(0.0, float(lab[0]) - 8.0), 1),
        "lightness_to": round(min(100.0, float(lab[0]) + 8.0), 1),
        "lab_hue_deg": round(lab_hue, 1) if chroma >= 4.0 else None,
        "share": round(share, 4),
        "rgb": [int(x) for x in rgb],
        "lab": [round(float(x), 2) for x in lab],
    }


def dominant_colors(image_path: Path, maximum: int = 3) -> tuple[list[dict], str | None]:
    with Image.open(image_path) as source:
        image = ImageOps.exif_transpose(source).convert("RGB")
        image.thumbnail((360, 360))
        rgb = np.asarray(image)
    if min(rgb.shape[:2]) < 24:
        return [], "image too small"

    lab = srgb_to_lab(rgb)
    h, w, _ = lab.shape
    edge = max(2, int(min(h, w) * 0.055))
    ring = np.concatenate((lab[:edge].reshape(-1, 3), lab[-edge:].reshape(-1, 3),
                           lab[:, :edge].reshape(-1, 3), lab[:, -edge:].reshape(-1, 3)))
    bg = np.median(ring, axis=0)

    # Favor the central product while retaining enough area for wide garlands,
    # ribbons, and floral arrangements. The edge-derived mask removes studio
    # backgrounds but still permits genuinely white and black objects.
    yy, xx = np.ogrid[:h, :w]
    ellipse = ((xx - w / 2) / (w * 0.49)) ** 2 + ((yy - h / 2) / (h * 0.49)) ** 2 <= 1
    distance = np.sqrt(((lab - bg) ** 2).sum(axis=2))
    ring_distance = np.sqrt(((ring - bg) ** 2).sum(axis=1))
    bg_threshold = max(12.0, float(np.percentile(ring_distance, 90)) + 3.0)
    chroma = np.hypot(lab[:, :, 1], lab[:, :, 2])
    background_chroma = float(np.median(np.hypot(ring[:, 1], ring[:, 2])))
    neutral_limit = max(10.0, min(14.0, background_chroma + 5.0))
    gray_background = (
        (chroma < neutral_limit)
        & (lab[:, :, 0] > 20.0)
        & (lab[:, :, 0] < 82.0)
    )
    keep = ellipse & (distance >= bg_threshold) & ~gray_background
    pixels = lab[keep]
    if len(pixels) < 250:
        pixels = lab[ellipse & ~gray_background]
    if len(pixels) > 18000:
        step = max(1, len(pixels) // 18000)
        pixels = pixels[::step]
    if len(pixels) < 100:
        return [], "review: no non-gray product pixels"

    centers, labels = kmeans(pixels, min(7, max(2, len(pixels) // 500)))
    sizes = np.bincount(labels, minlength=len(centers)).astype(float)
    candidates = []
    for idx in np.argsort(-sizes):
        share = sizes[idx] / sizes.sum()
        if share < 0.02:
            continue
        center = np.median(pixels[labels == idx], axis=0)
        # Reject a pale cluster only when it is close to the measured frame.
        if np.linalg.norm(center - bg) < bg_threshold + 2.0:
            continue
        candidates.append((center, share))

    if not candidates:
        # Pixels are already free of the neutral gray background, so this
        # fallback cannot reintroduce it as a product color.
        idx = int(np.argmax(sizes))
        candidates = [(np.median(pixels[labels == idx], axis=0), float(sizes[idx] / sizes.sum()))]

    # Merge lighting variants of one material. A dark and a light green are
    # still one product color; neutral black and white remain distinct.
    merged = []
    for center, share in candidates:
        chroma = math.hypot(float(center[1]), float(center[2]))
        hue = (math.degrees(math.atan2(float(center[2]), float(center[1]))) + 360.0) % 360.0
        match = None
        for group in merged:
            old_chroma = math.hypot(float(group[0][1]), float(group[0][2]))
            old_hue = (math.degrees(math.atan2(float(group[0][2]), float(group[0][1]))) + 360.0) % 360.0
            hue_gap = abs(hue - old_hue)
            hue_gap = min(hue_gap, 360.0 - hue_gap)
            both_colored = chroma >= 7.0 and old_chroma >= 7.0
            both_neutral = chroma < 7.0 and old_chroma < 7.0
            if (both_colored and hue_gap <= 16.0) or (both_neutral and abs(center[0] - group[0][0]) <= 38.0):
                match = group
                break
        if match is None:
            merged.append([center.copy(), share, float(np.linalg.norm(center - bg))])
        else:
            new_share = match[1] + share
            if both_neutral:
                distance_from_bg = float(np.linalg.norm(center - bg))
                if distance_from_bg > match[2]:
                    match[0] = center.copy()
                    match[2] = distance_from_bg
            else:
                match[0] = (match[0] * match[1] + center * share) / new_share
            match[1] = new_share

    merged = sorted(merged, key=lambda item: item[1], reverse=True)
    selected = [(item[0], item[1]) for item in merged if item[1] >= 0.10][:maximum]
    if not selected and candidates:
        selected = [candidates[0]]
    total = sum(share for _, share in selected) or 1.0
    return [color_record(center, share / total) for center, share in selected], None


def image_suffix(url: str) -> str:
    suffix = Path(urllib.parse.urlsplit(url).path).suffix.lower()
    return suffix if suffix in (".jpg", ".jpeg", ".png", ".webp") else ".img"


def scan(args: argparse.Namespace) -> list[dict]:
    work = Path(args.work_dir)
    fetcher = CachedFetcher(work / "cache", args.pause)
    products = discover_products(fetcher, args.max_category_pages)
    if args.offset:
        products = products[args.offset:]
    if args.limit:
        products = products[:args.limit]
    (work / "product_index.json").write_text(json.dumps(products, ensure_ascii=False, indent=2), encoding="utf-8")
    if args.index_only:
        print("index only, products %d" % len(products))
        return []
    previous = {}
    if args.previous and Path(args.previous).exists():
        old_rows = json.loads(Path(args.previous).read_text(encoding="utf-8"))
        previous = {(row.get("product_url"), row.get("image_url")): row
                    for row in old_rows if row.get("colors")}
    rows = []
    for index, product in enumerate(products, 1):
        url = product["product_url"]
        for photo_index, image_url in enumerate(product["images"], 1):
            old = previous.get((url, image_url))
            if old:
                rows.append({**old, "title": product["title"],
                             "photo_index": photo_index, "status": "ok"})
                continue
            try:
                image_path = fetcher._path(image_url, image_suffix(image_url))
                if not image_path.exists():
                    image_path.write_bytes(fetcher.bytes(image_url, image_suffix(image_url)))
                colors, error = dominant_colors(image_path)
                rows.append({
                    "title": product["title"], "sku": product["sku"],
                    "product_url": url, "image_url": image_url,
                    "photo_index": photo_index, "colors": colors,
                    "status": error or "ok",
                })
            except Exception as exc:
                rows.append({
                    "title": product["title"], "sku": product["sku"],
                    "product_url": url, "image_url": image_url,
                    "photo_index": photo_index, "colors": [],
                    "status": "image error: %s" % exc,
                })
        if index % 20 == 0 or index == len(products):
            print("products %d/%d, rows %d" % (index, len(products), len(rows)))
            Path(args.output).write_text(json.dumps(rows, ensure_ascii=False, indent=2), encoding="utf-8")
    return rows


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser()
    parser.add_argument("--work-dir", default="catalog_scan")
    parser.add_argument("--output", default="catalog_scan/product_colors.json")
    parser.add_argument("--pause", type=float, default=0.35)
    parser.add_argument("--limit", type=int)
    parser.add_argument("--offset", type=int, default=0)
    parser.add_argument("--max-category-pages", type=int)
    parser.add_argument("--index-only", action="store_true")
    parser.add_argument("--previous", help="Reuse colors when product and image URLs are unchanged")
    return parser


if __name__ == "__main__":
    scan(build_parser().parse_args())
