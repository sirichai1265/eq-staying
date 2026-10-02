# -*- coding: utf-8 -*-
"""
สร้าง EQ STAYING Dashboard (AV / DMG / LONG STAYING) จากไฟล์ STAYING .xls

วิธีใช้:
    python build_dashboard.py                      # ใช้ไฟล์ *STAYING*.xls ล่าสุดในโฟลเดอร์นี้
    python build_dashboard.py "10-2-STAYING - Copy.xls"
ผลลัพธ์: index.html (เปิดด้วย browser ได้เลย ไม่ต้องต่อเน็ต)
"""
import base64
import glob
import json
import os
import sys
from datetime import datetime

import pandas as pd

HERE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(HERE, "index.html")

AV_CODES = {"IEC", "IED", "IEP", "IER", "VED"}
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


if __name__ == "__main__":
    main()
