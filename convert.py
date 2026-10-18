#!/usr/bin/env python3
"""รวม data/*.xlsx (เหตุการณ์ T/R, T/L) -> data.js   ใช้: python convert.py [โฟลเดอร์] [data.js]"""
import glob, json, os, re, sys, datetime as dt
from openpyxl import load_workbook

SRC = sys.argv[1] if len(sys.argv) > 1 else "data"
OUT = sys.argv[2] if len(sys.argv) > 2 else "data.js"

A = {  # คอลัมน์ -> ชื่อหัวที่รองรับ
    "t": ["dateoutage", "date", "datetime", "วันที่", "วันเวลา"],
    "f": ["equiptment", "equipment", "feeder", "อุปกรณ์", "ฟีดเดอร์"],
    "k": ["tripreclose", "type", "ประเภท"],
    "r": ["reason", "สาเหตุ", "remark"],
}
KEY = {a: k for k, v in A.items() for a in v}


def when(v):
    """'dd-mm-yyyy hh:mm' (พ.ศ. หรือ ค.ศ.) หรือ datetime -> datetime ค.ศ."""
    if isinstance(v, dt.datetime):
        d = v
    else:
        m = re.match(r"\s*(\d{1,2})[-/.](\d{1,2})[-/.](\d{4})(?:\D+(\d{1,2}):(\d{2}))?", str(v or ""))
        if not m:
            return None
        dd, mm, yy, h, mi = m.groups()
        try:
            d = dt.datetime(int(yy), int(mm), int(dd), int(h or 0), int(mi or 0))
        except ValueError:
            return None
    return d.replace(year=d.year - 543) if d.year > 2400 else d


def kind(k, r):
    for x in (str(k or ""), str(r or "")[:4]):
        if re.search(r"t\s*/\s*l|lock", x, re.I):
            return "T/L"
        if re.search(r"t\s*/\s*r|reclos", x, re.I):
            return "T/R"


rows, seen, skipped = [], set(), 0
files = sorted(glob.glob(os.path.join(SRC, "**", "*.xlsx"), recursive=True))
for f in files:
    fb = re.search(r"Event_Data_(.+?)\.xlsx$", os.path.basename(f), re.I)
    wb = load_workbook(f, read_only=True, data_only=True)
    for ws in wb.worksheets:
        it = ws.iter_rows(values_only=True)
        cols = [KEY.get(str(h or "").strip().lower()) for h in next(it, [])]
        if "t" not in cols:
            print(f"ข้ามชีต {os.path.basename(f)}/{ws.title}: ไม่พบคอลัมน์วันที่")
            continue
        for r in it:
            o = {k: c for k, c in zip(cols, r) if k and c is not None}
            if not o:
                continue
            d, k = when(o.get("t")), kind(o.get("k"), o.get("r"))
            fd = str(o.get("f") or (fb.group(1) if fb else "")).strip()
            if not (d and k and fd):
                skipped += 1
                continue
            key = (fd, d, k)
            if key in seen:
                continue
            seen.add(key)
            rows.append([d.strftime("%Y-%m-%dT%H:%M"), fd, k, re.sub(r"\s+", " ", str(o.get("r") or "")).strip()])

rows.sort(key=lambda x: (x[0], x[1]), reverse=True)
upd = dt.datetime.now(dt.timezone(dt.timedelta(hours=7))).strftime("%Y-%m-%dT%H:%M")
body = ",\n".join(json.dumps(r, ensure_ascii=False) for r in rows)
open(OUT, "w", encoding="utf-8").write(f'window.DATA={{"updated":"{upd}","rows":[\n{body}\n]}};\n')
print(f"{OUT}: {len(rows)} เหตุการณ์ จาก {len(files)} ไฟล์ (ข้ามแถวที่อ่านไม่ได้ {skipped})")
