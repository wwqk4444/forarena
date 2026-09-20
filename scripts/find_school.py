#!/usr/bin/env python3
"""
从 docx / xlsx / htm(html) / csv 格式的获奖名单中筛出包含关键词（默认“曹杨”）的表格行。
不依赖第三方库（docx/xlsx 直接按 zip+XML 解析）。

用法:
    python3 scripts/find_school.py data/raw/NOI2025冬令营获奖名单.docx [更多文件...] [--kw 曹杨] [--all]
    --all  输出全部行（不过滤），便于人工核对
"""
import csv
import html
import io
import re
import sys
import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
S = "{http://schemas.openxmlformats.org/spreadsheetml/2006/main}"


def rows_from_docx(path):
    with zipfile.ZipFile(path) as z:
        root = ET.fromstring(z.read("word/document.xml"))
    for tbl in root.iter(W + "tbl"):
        for tr in tbl.iter(W + "tr"):
            cells = []
            for tc in tr.iter(W + "tc"):
                cells.append("".join(t.text or "" for t in tc.iter(W + "t")).strip())
            yield cells


def rows_from_xlsx(path):
    with zipfile.ZipFile(path) as z:
        shared = []
        if "xl/sharedStrings.xml" in z.namelist():
            sroot = ET.fromstring(z.read("xl/sharedStrings.xml"))
            for si in sroot.iter(S + "si"):
                shared.append("".join(t.text or "" for t in si.iter(S + "t")))
        sheets = sorted(n for n in z.namelist() if re.match(r"xl/worksheets/sheet\d+\.xml", n))
        for name in sheets:
            root = ET.fromstring(z.read(name))
            for row in root.iter(S + "row"):
                cells = []
                for c in row.iter(S + "c"):
                    v = c.find(S + "v")
                    if v is None:
                        is_ = c.find(S + "is")
                        cells.append("".join(t.text or "" for t in is_.iter(S + "t")) if is_ is not None else "")
                        continue
                    val = v.text or ""
                    if c.get("t") == "s":
                        val = shared[int(val)]
                    cells.append(val.strip())
                yield cells


def rows_from_html(path):
    text = Path(path).read_text(encoding="utf-8", errors="ignore")
    for tr in re.findall(r"<tr[^>]*>(.*?)</tr>", text, flags=re.S | re.I):
        cells = re.findall(r"<t[dh][^>]*>(.*?)</t[dh]>", tr, flags=re.S | re.I)
        cells = [html.unescape(re.sub(r"<[^>]+>", "", c)).strip() for c in cells]
        yield cells


def rows_from_csv(path):
    with open(path, encoding="utf-8-sig", newline="") as f:
        for row in csv.reader(f):
            yield [c.strip() for c in row]


READERS = {".docx": rows_from_docx, ".xlsx": rows_from_xlsx, ".htm": rows_from_html,
           ".html": rows_from_html, ".csv": rows_from_csv}


def main(argv):
    kw, show_all, files = "曹杨", False, []
    it = iter(argv)
    for a in it:
        if a == "--kw":
            kw = next(it)
        elif a == "--all":
            show_all = True
        else:
            files.append(a)
    if not files:
        print(__doc__)
        return 1
    out = csv.writer(sys.stdout)
    for f in files:
        reader = READERS.get(Path(f).suffix.lower())
        if reader is None:
            print(f"# 跳过不支持的格式: {f}", file=sys.stderr)
            continue
        n = 0
        for cells in reader(f):
            if show_all or any(kw in c for c in cells):
                out.writerow([Path(f).name] + cells)
                n += 1
        print(f"# {f}: 命中 {n} 行", file=sys.stderr)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv[1:]))
