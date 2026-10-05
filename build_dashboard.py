# -*- coding: utf-8 -*-
"""
สร้าง EQ STAYING Dashboard (AV / DMG / LONG STAYING) จากไฟล์ STAYING .xls

วิธีใช้:
    python build_dashboard.py                      # ใช้ไฟล์ *STAYING*.xls ล่าสุดในโฟลเดอร์นี้
    python build_dashboard.py "C:/Users/HAL-USER/Desktop/10-5-STAYING.xls"
ผลลัพธ์: index.html (เปิดด้วย browser ได้เลย ไม่ต้องต่อเน็ต)
และเก็บไฟล์ของรอบนี้ (ไฟล์ต้นฉบับ + dashboard + LONG STAYING .xlsx) ไว้ในโฟลเดอร์วันที่ เช่น 2026-10-05/
"""
import base64
import glob
import json
import os
import shutil
import sys
from datetime import datetime

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "index.html")

AV_CODES = {"IEC", "IED", "IEP", "IER", "IEW", "VED"}
DMG_CODES = {"OER"}
BKK_OVERRIDE_LOCATIONS = {"LCH55", "LCH54"}
REEFER_TYPES = {"22RE", "45RE"}


def pick_source():
    if len(sys.argv) > 1:
        return sys.argv[1]
    files = [f for f in glob.glob(os.path.join(HERE, "*STAYING*.xls*")) if not os.path.basename(f).startswith("~$")]
    if not files:
        sys.exit("ไม่พบไฟล์ *STAYING*.xls ในโฟลเดอร์")
    return max(files, key=os.path.getmtime)


def group_of(code):
    if code in AV_CODES:
        return "AV"
    if code in DMG_CODES:
        return "DMG"
    return "OTHER"


def area_of(location, area):
    if location in BKK_OVERRIDE_LOCATIONS:
        return "BKK"
    if location.startswith("BKK"):
        return "BKK"
    if location.startswith("LCH"):
        return "LCH"
    return area or "-"


def clean(v):
    if pd.isna(v):
        return ""
    return str(v).strip()


def fmt_vl(v):
    """202604211133 -> 2026-04-21 11:33"""
    if pd.isna(v):
        return ""
    t = str(int(v))
    return f"{t[:4]}-{t[4:6]}-{t[6:8]} {t[8:10]}:{t[10:12]}" if len(t) == 12 else t


def logo_data_uri():
    path = os.path.join(HERE, "heunga_logo.png")
    if not os.path.exists(path):
        return ""
    with open(path, "rb") as f:
        return "data:image/png;base64," + base64.b64encode(f.read()).decode()


def main():
    src = pick_source()
    df = pd.read_excel(src)
    # ตัดแถวสรุปท้ายไฟล์ / แถวว่าง
    df = df[df["Location"].notna() & df["Size/Type"].notna()].copy()

    rows = []
    for r in df.itertuples(index=False):
        row = dict(zip(df.columns, r))
        loc = clean(row["Location"])
        code = clean(row["Move Code"])
        size = clean(row["Size/Type"])
        by = row["Built Year"]
        rows.append({
            "cn": clean(row["Container No"]),
            "st": size,
            "loc": loc,
            "ar": area_of(loc, clean(row["Area"])),
            "mc": code,
            "gp": group_of(code),
            "d": int(row["Days"]) if pd.notna(row["Days"]) else 0,
            "eq": clean(row["EQ Date"])[:10],
            "by": int(by) if pd.notna(by) else None,
            "rf": clean(row["RF Brand"]).upper() if size in REEFER_TYPES else "",
            "vl": fmt_vl(row["VL Date"]),
            "fe": clean(row["Full/Empty"]),
        })

    rows.sort(key=lambda x: -x["d"])
    eq_dates = [x["eq"] for x in rows if x["eq"]]
    meta = {
        "source": os.path.basename(src),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "lastEq": max(eq_dates) if eq_dates else "",
    }

    with open(os.path.join(HERE, "dashboard_template.html"), encoding="utf-8") as f:
        html = f.read()
    html = html.replace("/*__DATA__*/null", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
    html = html.replace("__LOGO__", logo_data_uri())
    html = html.replace("/*__META__*/null", json.dumps(meta, ensure_ascii=False))
    with open(OUT, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"Source : {meta['source']}  ({len(rows)} containers)")
    for g in ("AV", "DMG", "OTHER"):
        print(f"  {g:5}: {sum(1 for x in rows if x['gp'] == g)}")
    print(f"  >=100 days: {sum(1 for x in rows if x['d'] >= 100)}   >=200 days: {sum(1 for x in rows if x['d'] >= 200)}")
    print(f"Output : {OUT}")

    archive(src, rows, meta)


def write_long_staying_xlsx(rows, path):
    """รายการแจ้งเตือน LONG STAYING (>= 100 วัน) แบ่งสีเหมือนใน dashboard"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    ws_rows = [r for r in rows if r["d"] >= 100]
    wb = Workbook()
    ws = wb.active
    ws.title = "LONG STAYING"
    head = ["#", "Area", "Location", "Days", "Aging", "Container No", "Size/Type", "Group",
            "Move Code", "EQ Date", "Built Year", "RF Brand"]
    widths = [6, 8, 11, 8, 12, 16, 10, 8, 11, 12, 11, 13]
    ws.append(head)
    for i, r in enumerate(ws_rows, 1):
        lo = r["d"] // 50 * 50
        ws.append([i, r["ar"], r["loc"], r["d"], f"{lo}–{lo + 49} วัน", r["cn"], r["st"], r["gp"],
                   r["mc"], r["eq"], r["by"] or "", r["rf"]])
    thin = Side(style="thin", color="BFC5CE")
    border = Border(top=thin, left=thin, bottom=thin, right=thin)
    red_fill = PatternFill("solid", fgColor="FEE2E2")
    yellow_fill = PatternFill("solid", fgColor="FFF59D")
    for c, w in zip("ABCDEFGHIJKL", widths):
        ws.column_dimensions[c].width = w
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill("solid", fgColor="1F5FBF")
        cell.alignment = Alignment(horizontal="center")
        cell.border = border
    for row in ws.iter_rows(min_row=2):
        fill = yellow_fill if row[3].value >= 200 else red_fill
        for cell in row:
            cell.fill = fill
            cell.border = border
        row[3].font = Font(bold=True, color="B91C1C")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = f"A1:L{max(len(ws_rows) + 1, 1)}"
    wb.save(path)
    return len(ws_rows)


def archive(src, rows, meta):
    """เก็บไฟล์ของรอบนี้ไว้ในโฟลเดอร์วันที่ เช่น EQ STAYING/2026-10-05/"""
    day = datetime.now().strftime("%Y-%m-%d")
    folder = os.path.join(HERE, day)
    os.makedirs(folder, exist_ok=True)
    # ไฟล์ต้นฉบับ: ย้ายเข้าโฟลเดอร์วัน (ถ้ายังไม่อยู่ในนั้น)
    dest_src = os.path.join(folder, os.path.basename(src))
    if os.path.abspath(src) != os.path.abspath(dest_src):
        shutil.move(src, dest_src)
    shutil.copy2(OUT, os.path.join(folder, f"EQ_STAYING_Dashboard_{day}.html"))
    n = write_long_staying_xlsx(rows, os.path.join(folder, f"LONG_STAYING_{day}.xlsx"))
    print(f"Archive: {folder}")
    print(f"  - {os.path.basename(src)} (ไฟล์ต้นฉบับ)")
    print(f"  - EQ_STAYING_Dashboard_{day}.html")
    print(f"  - LONG_STAYING_{day}.xlsx ({n} ตู้)")


if __name__ == "__main__":
    main()
