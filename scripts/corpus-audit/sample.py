#!/usr/bin/env python3
"""Seeded random sample of the .xlsx members of a (gzipped) tar archive, streamed twice: once to list, once to extract.

Members are written under their sha256 (flat, never the archive path) so a hostile member name cannot escape the output
directory, and a manifest records the original name, size and sha256 of each one.

    python3 -I scripts/corpus-audit/sample.py <archive.tar.gz> <out-dir> [--count 3000] [--seed 20261111]
"""
import argparse
import hashlib
import json
import os
import random
import sys
import tarfile


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("archive")
    parser.add_argument("out")
    parser.add_argument("--count", type=int, default=3000)
    parser.add_argument("--seed", type=int, default=20261111)
    args = parser.parse_args()

    names = []
    with tarfile.open(args.archive, "r|gz") as archive:
        for member in archive:
            if member.isfile() and member.name.lower().endswith((".xlsx", ".xlsm")) and not os.path.basename(member.name).startswith("._"):
                names.append(member.name)
    print(f"{len(names)} xlsx members", file=sys.stderr)
    chosen = set(random.Random(args.seed).sample(names, min(args.count, len(names))))

    os.makedirs(args.out, exist_ok=True)
    manifest = []
    with tarfile.open(args.archive, "r|gz") as archive:
        for member in archive:
            if member.name not in chosen or not member.isfile():
                continue
            data = archive.extractfile(member).read()
            digest = hashlib.sha256(data).hexdigest()
            extension = ".xlsm" if member.name.lower().endswith(".xlsm") else ".xlsx"
            path = os.path.join(args.out, digest + extension)
            with open(path, "wb") as handle:
                handle.write(data)
            manifest.append({"file": digest + extension, "source": member.name, "bytes": len(data), "sha256": digest})
    with open(os.path.join(args.out, "manifest.json"), "w") as handle:
        json.dump({"archive": os.path.basename(args.archive), "seed": args.seed, "population": len(names), "files": manifest}, handle, indent=1)
    print(f"wrote {len(manifest)} files to {args.out}", file=sys.stderr)


if __name__ == "__main__":
    main()
