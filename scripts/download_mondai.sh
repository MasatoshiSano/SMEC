#!/usr/bin/env bash
# Download all official 中小企業診断士 exam question PDFs from J-SMECA (now jf-cmca.jp).
# Source listing page: https://www.jf-cmca.jp/contents/010_c_/shikenmondai.html
# Usage: scripts/download_mondai.sh <url_list_file> <dest_root>
set -uo pipefail

URL_LIST="${1:?url list file required}"
DEST_ROOT="${2:?dest root required}"
UA="Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120 Safari/537.36"

ok=0; fail=0
while IFS= read -r url; do
  [ -z "$url" ] && continue
  # Map .../shikenmondai/<dir>/<file>.pdf  ->  <DEST_ROOT>/<1ji|2ji sorted>/<dir>/<file>.pdf
  rel="${url#*/shikenmondai/}"          # e.g. 1ji2026/A1JC2026.pdf
  case "$rel" in
    1ji*|1JI*) sub="1ji" ;;
    2ji*|2JI*) sub="2ji" ;;
    *) sub="other" ;;
  esac
  out="$DEST_ROOT/$sub/$rel"
  mkdir -p "$(dirname "$out")"
  if [ -s "$out" ]; then
    echo "SKIP (exists): $rel"; ok=$((ok+1)); continue
  fi
  code=$(curl -sS -A "$UA" -w "%{http_code}" -o "$out" "$url")
  if [ "$code" = "200" ] && [ -s "$out" ]; then
    echo "OK  [$code] $rel ($(wc -c < "$out") bytes)"; ok=$((ok+1))
  else
    echo "ERR [$code] $rel"; rm -f "$out"; fail=$((fail+1))
  fi
  sleep 0.3
done < "$URL_LIST"

echo "-----"
echo "done: ok=$ok fail=$fail"
