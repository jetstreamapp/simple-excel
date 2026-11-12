#!/usr/bin/env python3
"""Group fetched spreadsheets by the application that wrote them and copy up to --per files of each into
<out>/<writer>/, so run.ts reads a corpus balanced across writers and report.ts can break results down by folder.

The writer key starts from docProps/app.xml's Application and AppVersion, which several libraries fill with
"Microsoft Excel" (PHPExcel, some .NET writers), so two more signals from xl/workbook.xml separate the impostors:
Excel writes <fileVersion appName="xl" ... rupBuild="..."/>, and most libraries write no fileVersion at all.
Nothing in the files is executed; they are only unzipped in memory to read those two parts.

    python3 -I scripts/corpus-audit/select-by-writer.py <fetched-dir> <out> [--per 300] [--seed 20261111]
"""
import argparse
import collections
import json
import os
import random
import re
import shutil
import zipfile


def read_part(archive, name):
    try:
        return archive.read(name).decode("utf-8", "replace")
    except KeyError:
        return ""


def library_fingerprint(names, workbook):
    """Writers that leave Application out or fake it, recognised by what they put in the package (seen in FUSE)."""
    if any("\\" in name for name in names):
        return "backslash-entry-names"
    if re.search(r'r:id="R[0-9a-f]{16}"', workbook):
        return "openxml-sdk-dotnet"
    if "schemas.microsoft.com/office/mac/excel/2008/main" in workbook and "docProps/app.xml" not in names:
        return "google-sheets-2014"
    view = re.search(r"<(?:\w+:)?workbookView\s+(\w+)=", workbook)
    if view and view.group(1) == "activeTab" and 'rupBuild="4505"' in workbook:
        return "phpexcel"
    return None


def writer_key(path):
    try:
        with zipfile.ZipFile(path) as archive:
            names = archive.namelist()
            app = read_part(archive, "docProps/app.xml")
            workbook_name = next((name for name in names if name.replace("\\", "/").lower() == "xl/workbook.xml"), "")
            workbook = read_part(archive, workbook_name) if workbook_name else ""
    except Exception:
        return "unreadable-zip"
    library = library_fingerprint(names, workbook)
    if library:
        return library
    application = re.search(r"<(?:\w+:)?Application>([^<]*)<", app)
    version = re.search(r"<(?:\w+:)?AppVersion>([^<]*)<", app)
    name = application.group(1).strip() if application else "no-application"
    major = version.group(1).split(".")[0] if version else ""
    file_version = re.search(r"<(?:\w+:)?fileVersion\b[^>]*>", workbook)
    if file_version is None:
        excel_mark = "nofv"
    elif "rupBuild" in file_version.group(0):
        excel_mark = "rup"
    else:
        excel_mark = "fv"
    key = f"{name} {major} {excel_mark}".strip() if name.lower().startswith("microsoft") else name
    return re.sub(r"[^A-Za-z0-9.+-]+", "_", key).strip("_") or "unknown"


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("fetched")
    parser.add_argument("out")
    parser.add_argument("--per", type=int, default=300)
    parser.add_argument("--seed", type=int, default=20261111)
    args = parser.parse_args()

    files_dir = os.path.join(args.fetched, "files")
    groups = collections.defaultdict(list)
    for name in sorted(os.listdir(files_dir)):
        groups[writer_key(os.path.join(files_dir, name))].append(name)

    rng = random.Random(args.seed)
    summary = {}
    for key, names in sorted(groups.items(), key=lambda item: -len(item[1])):
        chosen = names if len(names) <= args.per else rng.sample(names, args.per)
        target = os.path.join(args.out, key)
        os.makedirs(target, exist_ok=True)
        for name in chosen:
            shutil.copyfile(os.path.join(files_dir, name), os.path.join(target, name))
        summary[key] = {"population": len(names), "selected": len(chosen)}
        print(f"{len(names):6d} {len(chosen):5d}  {key}")
    with open(os.path.join(args.out, "writers.json"), "w") as handle:
        json.dump({"per": args.per, "seed": args.seed, "writers": summary}, handle, indent=1)


if __name__ == "__main__":
    main()
