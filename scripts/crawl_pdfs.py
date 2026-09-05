#!/usr/bin/env python3
"""Crawl a list of HTML pages on jf-cmca.jp and emit absolute PDF URLs.

Usage: crawl_pdfs.py <pages_file> > pdf_urls.txt
  <pages_file>: one absolute page URL per line.
"""
import re
import sys
import time
import urllib.request
from urllib.parse import urljoin

UA = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"
HREF = re.compile(r'href="([^"]+\.pdf)"', re.IGNORECASE)


def fetch(url: str) -> str:
    req = urllib.request.Request(url, headers={"User-Agent": UA})
    with urllib.request.urlopen(req, timeout=30) as resp:
        return resp.read().decode("utf-8", "ignore")


def main() -> None:
    pages_file = sys.argv[1]
    seen = set()
    with open(pages_file) as fh:
        pages = [ln.strip() for ln in fh if ln.strip()]
    for page in pages:
        try:
            html = fetch(page)
        except Exception as exc:  # noqa: BLE001
            print(f"# ERR fetching {page}: {exc}", file=sys.stderr)
            continue
        for m in HREF.findall(html):
            absolute = urljoin(page, m)
            if absolute not in seen:
                seen.add(absolute)
                print(absolute)
        time.sleep(0.2)
    print(f"# pages={len(pages)} pdfs={len(seen)}", file=sys.stderr)


if __name__ == "__main__":
    main()
