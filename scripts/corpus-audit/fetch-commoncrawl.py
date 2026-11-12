#!/usr/bin/env python3
"""Fetch the original bytes of spreadsheet records from Common Crawl, one HTTP range request per record.

Input is JSON lines with the WARC file name, the gzip member's offset and length, and the WAT path FUSE recorded (the
WARC lives next to it under warc/ instead of wat/). Each record is inflated, its WARC and HTTP headers are stripped,
chunked transfer encoding is undone, and the payload is written under its sha256 (never a name from the crawl).
docProps/app.xml's Application and AppVersion are recorded so the corpus can be grouped by writer. Truncated records
(Common Crawl cut payloads at 1 MiB in 2014) and payloads that are not a zip are recorded and skipped.

    python3 -I scripts/corpus-audit/fetch-commoncrawl.py <records.jsonl> <out-dir> [--jobs 4] [--limit N]

Resumable: records already in <out-dir>/fetched.jsonl are skipped.
"""
import argparse
import concurrent.futures
import hashlib
import io
import json
import os
import re
import subprocess
import sys
import threading
import time
import zipfile
import zlib

BASE = "https://data.commoncrawl.org/"
USER_AGENT = "simple-excel-corpus-audit/1.0 (compatibility research)"


def warc_url(record):
    path = record["path"].removeprefix("common-crawl/")
    folder = path.rsplit("/wat/", 1)[0]
    return f"{BASE}{folder}/warc/{record['file']}"


def fetch(record, attempts=5):
    """The system curl, for its CA store: python.org builds on macOS ship without one."""
    byte_range = f"{record['offset']}-{record['offset'] + record['length'] - 1}"
    for attempt in range(attempts):
        result = subprocess.run(
            ["curl", "-sS", "--fail", "--max-time", "60", "-A", USER_AGENT, "-r", byte_range, warc_url(record)],
            capture_output=True,
            check=False,
        )
        if result.returncode == 0:
            return result.stdout
        # 503 slow-down and transient network errors: back off and retry
        if attempt == attempts - 1:
            raise RuntimeError(result.stderr.decode("utf-8", "replace").strip()[:200])
        time.sleep(2**attempt + 1)
    return b""


CHUNK_SIZE = re.compile(rb"[0-9A-Fa-f]+")


def dechunk(body):
    """Undo chunked transfer encoding. The framing is crawl content: a size that is not plain hex (int() would take a
    sign, and a negative size moves backwards for ever) or a body that ends before the last chunk raises ValueError."""
    out = bytearray()
    position = 0
    while True:
        if position >= len(body):
            raise ValueError("chunked body ends before its last chunk")
        end = body.find(b"\r\n", position)
        if end < 0:
            raise ValueError("chunked body ends inside a chunk-size line")
        size_text = body[position:end].split(b";")[0].strip()
        if not CHUNK_SIZE.fullmatch(size_text):
            raise ValueError(f"bad chunk size {size_text[:20]!r}")
        size = int(size_text, 16)
        if size == 0:
            return bytes(out)
        data_end = end + 2 + size
        if data_end > len(body):
            raise ValueError("chunked body ends inside a chunk")
        out += body[end + 2 : data_end]
        position = data_end + 2


def payload(member):
    # One gzip member only: the range can run a few bytes into the next record, which gzip.decompress would try to read.
    raw = zlib.decompressobj(16 + zlib.MAX_WBITS).decompress(member)
    warc_end = raw.index(b"\r\n\r\n")
    warc_headers = raw[:warc_end].decode("latin-1")
    http = raw[warc_end + 4 :]
    http_end = http.index(b"\r\n\r\n")
    http_headers = http[:http_end].decode("latin-1").lower()
    body = http[http_end + 4 :]
    if "transfer-encoding: chunked" in http_headers:
        body = dechunk(body)
    truncated = re.search(r"WARC-Truncated:\s*(\S+)", warc_headers, re.I)
    return body, truncated.group(1) if truncated else None


def application(data):
    try:
        with zipfile.ZipFile(io.BytesIO(data)) as archive:
            app = archive.read("docProps/app.xml").decode("utf-8", "replace")
    except Exception:
        return None, None
    name = re.search(r"<(?:\w+:)?Application>([^<]*)<", app)
    version = re.search(r"<(?:\w+:)?AppVersion>([^<]*)<", app)
    return (name.group(1).strip() if name else None), (version.group(1).strip() if version else None)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("records")
    parser.add_argument("out")
    parser.add_argument("--jobs", type=int, default=4)
    parser.add_argument("--limit", type=int)
    args = parser.parse_args()

    files_dir = os.path.join(args.out, "files")
    os.makedirs(files_dir, exist_ok=True)
    log_path = os.path.join(args.out, "fetched.jsonl")
    done = set()
    if os.path.exists(log_path):
        with open(log_path) as handle:
            done = {json.loads(line)["key"] for line in handle if line.strip()}
    with open(args.records) as handle:
        records = [json.loads(line) for line in handle if line.strip()]
    todo = [record for record in records if f"{record['file']}@{record['offset']}" not in done and record.get("path")]
    if args.limit:
        todo = todo[: args.limit]
    print(f"{len(done)} done, {len(todo)} to fetch", file=sys.stderr)

    lock = threading.Lock()
    log = open(log_path, "a")
    counts = {}

    def work(record):
        entry = {"key": f"{record['file']}@{record['offset']}", "uri": record["uri"], "tika": record["tika"]}
        try:
            data, truncated = payload(fetch(record))
            entry["bytes"] = len(data)
            if truncated:
                entry["status"] = f"truncated:{truncated}"
            elif not data.startswith(b"PK"):
                entry["status"] = "not-zip"
            else:
                digest = hashlib.sha256(data).hexdigest()
                with open(os.path.join(files_dir, digest + ".xlsx"), "wb") as handle:
                    handle.write(data)
                entry.update(status="ok", sha256=digest)
                entry["application"], entry["appVersion"] = application(data)
        except Exception as error:
            entry["status"] = f"error:{type(error).__name__}:{str(error)[:120]}"
        with lock:
            log.write(json.dumps(entry) + "\n")
            log.flush()
            key = entry["status"].split(":")[0]
            counts[key] = counts.get(key, 0) + 1
            total = sum(counts.values())
            if total % 250 == 0:
                print(f"{total}/{len(todo)} {counts}", file=sys.stderr)

    with concurrent.futures.ThreadPoolExecutor(max_workers=args.jobs) as pool:
        list(pool.map(work, todo))
    print(f"finished {counts}", file=sys.stderr)


if __name__ == "__main__":
    main()
