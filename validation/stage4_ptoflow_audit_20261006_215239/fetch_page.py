"""Explicit bounded HTTPS reads with persistent cache and mandatory HTTP 429 stop."""
import hashlib
import json
import re
import sys
import urllib.error
import urllib.request
from datetime import datetime, timezone
from pathlib import Path

BASE = Path(__file__).resolve().parent
CACHE = BASE / "cache"
CACHE.mkdir(exist_ok=True)
LEDGER = BASE / "pages.json"
sys.stdout.reconfigure(encoding="utf-8")


def fetch(label, url, kind="category"):
    entries = json.loads(LEDGER.read_text(encoding="utf-8")) if LEDGER.exists() else []
    if (BASE / "HTTP_429_STOP.json").exists():
        raise RuntimeError("HTTP 429 stop is active: no further PTOflow requests")
    previous = next((x for x in entries if x["url"] == url), None)
    if previous:
        print(json.dumps({"offline_cached": True, **previous}, ensure_ascii=False))
        return previous
    if not re.fullmatch(r"[a-z0-9_-]+", label) or not url.startswith("https://ptoflow.ru/"):
        raise ValueError("Explicit PTOflow URL and safe cache label required")
    if kind == "detail" and sum(x["kind"] == "detail" for x in entries) >= 20:
        raise RuntimeError("20 detail-page ceiling reached")
    stamp = datetime.now(timezone.utc).isoformat()
    status, content_type, final_url, error = None, None, url, None
    data = b""
    try:
        with urllib.request.urlopen(url, timeout=40) as response:
            status = response.status
            content_type = response.headers.get("Content-Type")
            final_url = response.url
            data = response.read()
    except urllib.error.HTTPError as exc:
        status, content_type, final_url = exc.code, exc.headers.get("Content-Type"), exc.url
        data, error = exc.read(), str(exc)
    except Exception as exc:
        error = str(exc)
    target = CACHE / f"{len(entries):02d}_{label}.html"
    target.write_bytes(data)
    entry = {"url": url, "final_url": final_url, "timestamp_utc": stamp, "http_status": status,
             "content_type": content_type, "size_bytes": len(data),
             "sha256": hashlib.sha256(data).hexdigest(), "local_path": str(target.relative_to(BASE)),
             "kind": kind, "error": error}
    entries.append(entry)
    LEDGER.write_text(json.dumps(entries, ensure_ascii=False, indent=2), encoding="utf-8")
    if status == 429:
        (BASE / "HTTP_429_STOP.json").write_text(json.dumps(entry, ensure_ascii=False, indent=2), encoding="utf-8")
    print(json.dumps(entry, ensure_ascii=False))
    if status != 200:
        raise RuntimeError(f"Request failed, HTTP {status}: {error}")
    return entry


if __name__ == "__main__":
    fetch(sys.argv[1], sys.argv[2], sys.argv[3] if len(sys.argv) > 3 else "category")
