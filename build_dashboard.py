# -*- coding: utf-8 -*-
"""
สร้าง EQ STAYING Dashboard (AV / DMG / LONG STAYING + ตู้ปล่อยออก Gate Out)

วิธีใช้ (ใส่ไฟล์ STAYING และ/หรือ GATE OUT กี่ไฟล์ก็ได้ — แยกประเภทจากชื่อไฟล์):
    python build_dashboard.py "C:/Users/HAL-USER/Desktop/10-15-STAYING.xls"
    python build_dashboard.py "C:/Users/HAL-USER/Desktop/10-15-GATE OUT.xls"
    python build_dashboard.py "...STAYING.xls" "...GATE OUT.xls"
    python build_dashboard.py                      # สร้างใหม่จากไฟล์ที่เก็บไว้แล้ว
ผลลัพธ์: index.html
และเก็บไฟล์ของแต่ละวันไว้ในโฟลเดอร์วันที่ เช่น 2026-10-15/
    STAYING   -> ไฟล์ต้นฉบับ + EQ_STAYING_Dashboard_<วัน>.html + EQ_STAYING_Summary_<วัน>.xlsx
    GATE OUT  -> ไฟล์ต้นฉบับ + GATE_OUT_Report_<วัน>.xlsx (ตู้ปล่อยออก OEP/OEV/OEW พร้อมวันค้างก่อนปล่อย)
"""
import base64
import glob
import json
import os
import re
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


RELEASE_MOVES = {"OEP", "OEV", "OEW"}   # ปล่อยตู้ออกจากลาน
DAY_DIR = re.compile(r"\d{4}-\d{2}-\d{2}")


def is_gateout(path):
    return "GATE OUT" in os.path.basename(path).upper().replace("_", " ")


def xls_in(folder, gate):
    return sorted(f for f in glob.glob(os.path.join(folder, "*.xls*"))
                  if not os.path.basename(f).startswith(("~$", "EQ_STAYING_", "GATE_OUT_")) and is_gateout(f) == gate
                  and (gate or "STAYING" in os.path.basename(f).upper()))


def day_folders():
    return sorted(d for d in os.listdir(HERE) if DAY_DIR.fullmatch(d) and os.path.isdir(os.path.join(HERE, d)))


def latest_archived_staying():
    for d in reversed(day_folders()):
        files = xls_in(os.path.join(HERE, d), gate=False)
        if files:
            return files[-1]
    sys.exit("ไม่พบไฟล์ STAYING ที่เก็บไว้")


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
    args = sys.argv[1:]
    gate_files = [a for a in args if is_gateout(a)]
    stay_files = [a for a in args if not is_gateout(a)]

    # 1) ไฟล์ GATE OUT: ย้ายเข้าโฟลเดอร์ตามวันที่ Gate แล้วทำรายงาน
    for g in gate_files:
        archive_gateout(g)

    # 2) STAYING: ใช้ไฟล์ที่ให้มา หรือไฟล์ล่าสุดที่เก็บไว้
    if stay_files:
        for s in stay_files[:-1]:
            build_staying(s, archive_it=True, write_index=False)
        build_staying(stay_files[-1], archive_it=True)
    else:
        build_staying(latest_archived_staying(), archive_it=False)


def load_staying_rows(src):
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
    return rows


def build_staying(src, archive_it=True, write_index=True):
    rows = load_staying_rows(src)
    eq_dates = [x["eq"] for x in rows if x["eq"]]
    meta = {
        "source": os.path.basename(src),
        "generated": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "lastEq": max(eq_dates) if eq_dates else "",
    }
    gate = load_gateout_history()

    print(f"Source : {meta['source']}  ({len(rows)} containers)")
    for g in ("AV", "DMG", "OTHER"):
        print(f"  {g:5}: {sum(1 for x in rows if x['gp'] == g)}")
    print(f"  >=100 days: {sum(1 for x in rows if x['d'] >= 100)}   >=200 days: {sum(1 for x in rows if x['d'] >= 200)}")
    print(f"Gate out history: {len(gate)} ตู้ปล่อยออก ({', '.join(sorted({g['gd'] for g in gate})) or '-'})")
    if write_index or archive_it:
        with open(os.path.join(HERE, "dashboard_template.html"), encoding="utf-8") as f:
            html = f.read()
        html = html.replace("/*__DATA__*/null", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
        html = html.replace("/*__GATE__*/null", json.dumps(gate, ensure_ascii=False, separators=(",", ":")))
        html = html.replace("__LOGO__", logo_data_uri())
        html = html.replace("/*__META__*/null", json.dumps(meta, ensure_ascii=False))
        with open(OUT, "w", encoding="utf-8") as f:
            f.write(html)
        print(f"Output : {OUT}")
    if archive_it:
        archive(src, rows, meta)


# ===================== GATE OUT (ตู้ปล่อยออก) =====================
def read_gateout(path):
    g = pd.read_excel(path, header=0)
    g = g.rename(columns={g.columns[0]: "Loc"})
    g = g[g["Container No"].notna() & g["Move"].notna()].copy()
    return g


def gate_day(path):
    d = read_gateout(path)["GateDate"].dropna().astype(str).str[:10]
    return max(d) if len(d) else datetime.now().strftime("%Y-%m-%d")


def staying_lookup(gate_date):
    """ข้อมูล EQ Date / Move Code ก่อนปล่อย จากไฟล์ STAYING ที่เก็บไว้ (วันเดียวกันหรือก่อนหน้า ใหม่สุดก่อน)"""
    found = {}
    for d in reversed(day_folders()):
        if d > gate_date:
            continue
        for f in xls_in(os.path.join(HERE, d), gate=False):
            cols = ["Container No", "EQ Date", "Move Code", "Location"]
            s = pd.read_excel(f, usecols=cols).dropna(subset=["Container No", "EQ Date"])[cols]
            for cn, eq, mc, loc in s.itertuples(index=False):
                found.setdefault(str(cn).strip(), (str(eq)[:10], clean(mc), clean(loc)))
    return found


def gateout_rows(path):
    g = read_gateout(path)
    g = g[g["Move"].isin(RELEASE_MOVES)]
    if g.empty:
        return []
    look = staying_lookup(gate_day(path))
    out = []
    for r in g.itertuples(index=False):
        row = dict(zip(g.columns, r))
        cn = clean(row["Container No"])
        gd = clean(row["GateDate"])[:10]
        loc = clean(row["Loc"])
        eq, pmc, sloc = look.get(cn, ("", "", ""))
        days = (datetime.strptime(gd, "%Y-%m-%d") - datetime.strptime(eq, "%Y-%m-%d")).days if eq and eq <= gd else None
        t = clean(row["GateTime"]).split(".")[0].zfill(4)
        by = row.get("Built Year")
        out.append({
            "cn": cn, "mv": clean(row["Move"]), "gd": gd, "gt": f"{t[:2]}:{t[2:]}",
            "loc": loc, "ar": area_of(loc, ""), "st": clean(row["TPSZ"]),
            "eq": eq, "d": days, "pmc": pmc, "pgp": group_of(pmc) if pmc else "",
            "by": int(by) if pd.notna(by) and str(by).strip().isdigit() else None,
            "pod": clean(row.get("POD")), "voy": clean(row.get("Voyage")),
        })
    out.sort(key=lambda x: -(x["d"] if x["d"] is not None else -1))
    return out


def load_gateout_history():
    rows = []
    for d in day_folders():
        for f in xls_in(os.path.join(HERE, d), gate=True):
            rows += gateout_rows(f)
    seen, uniq_rows = set(), []
    for r in sorted(rows, key=lambda x: (x["gd"], x["gt"]), reverse=True):
        k = (r["cn"], r["gd"], r["gt"])
        if k not in seen:
            seen.add(k)
            uniq_rows.append(r)
    return uniq_rows


def archive_gateout(src):
    day = gate_day(src)
    folder = os.path.join(HERE, day)
    os.makedirs(folder, exist_ok=True)
    dest = os.path.join(folder, os.path.basename(src))
    if os.path.abspath(src) != os.path.abspath(dest):
        shutil.move(src, dest)
    rows = gateout_rows(dest)
    path = os.path.join(folder, f"GATE_OUT_Report_{day}.xlsx")
    counts = write_gateout_xlsx(rows, day, os.path.basename(dest), path)
    print(f"Gate out {day}: {len(rows)} ตู้ปล่อยออก (OEP/OEV/OEW)")
    print(f"  - {os.path.basename(dest)} (ไฟล์ต้นฉบับ)")
    print(f"  - GATE_OUT_Report_{day}.xlsx  " + " · ".join(f"{k}: {v:,}" for k, v in counts.items()))


# กลุ่มวันตามสี (เหมือนใน dashboard): ชื่อชีท, ช่วงวัน, สีพื้น, สีตัวเลข Days
DAY_BUCKETS = [
    ("0-30 วัน", 0, 30, "DCFCE7", "15803D"),
    ("31-99 วัน", 31, 99, "FFEDD5", "C2410C"),
    ("100-199 วัน", 100, 199, "FEE2E2", "B91C1C"),
    ("200+ วัน", 200, 10**9, "FFF59D", "B91C1C"),
]


def write_summary_xlsx(rows, meta, path):
    """Excel สรุป: ชีท 'สรุป' + แยกชีทละกลุ่มวัน (0-30 / 31-99 / 100-199 / 200+)"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    thin = Side(style="thin", color="BFC5CE")
    border = Border(top=thin, left=thin, bottom=thin, right=thin)
    head_fill = PatternFill("solid", fgColor="1F5FBF")
    total_fill = PatternFill("solid", fgColor="EEF1F5")

    def style_head(cells):
        for c in cells:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = head_fill
            c.alignment = Alignment(horizontal="center")
            c.border = border

    in_bucket = {b[0]: [r for r in rows if b[1] <= r["d"] <= b[2]] for b in DAY_BUCKETS}

    wb = Workbook()
    ws = wb.active
    ws.title = "สรุป"
    ws["A1"] = "EQ STAYING — สรุปตามกลุ่มวัน"
    ws["A1"].font = Font(bold=True, size=14)
    ws["A2"] = f"ไฟล์: {meta['source']} · EQ Date ล่าสุด: {meta['lastEq']} · ทั้งหมด {len(rows):,} ตู้"
    ws["A2"].font = Font(color="6B7684")

    def table(start_row, title, keys, key_fn):
        ws.cell(start_row, 1, title).font = Font(bold=True, size=12)
        r0 = start_row + 1
        hdr = [""] + [b[0] for b in DAY_BUCKETS] + ["รวม"]
        for j, h in enumerate(hdr, 1):
            ws.cell(r0, j, h)
        style_head(ws[r0][:len(hdr)])
        for b, cell in zip(DAY_BUCKETS, ws[r0][1:len(DAY_BUCKETS) + 1]):
            cell.fill = PatternFill("solid", fgColor=b[3])
            cell.font = Font(bold=True, color=b[4])
        for i, k in enumerate(keys + ["รวม"], 1):
            rr = r0 + i
            vals = [sum(1 for x in in_bucket[b[0]] if k == "รวม" or key_fn(x) == k) for b in DAY_BUCKETS]
            for j, v in enumerate([k] + vals + [sum(vals)], 1):
                c = ws.cell(rr, j, v)
                c.border = border
                if k == "รวม" or j in (1, len(hdr)):
                    c.font = Font(bold=True)
                if k == "รวม":
                    c.fill = total_fill
        return r0 + len(keys) + 3

    nxt = table(4, "แยกตาม Area", ["BKK", "LCH"], lambda x: x["ar"])
    nxt = table(nxt, "แยกตาม Group", ["AV", "DMG", "OTHER"], lambda x: x["gp"])
    sizes = sorted({x["st"] for x in rows}, key=lambda s: -sum(1 for x in rows if x["st"] == s))
    nxt = table(nxt, "แยกตาม Size/Type", sizes, lambda x: x["st"])
    locs = sorted({(x["ar"], x["loc"]) for x in rows})
    table(nxt, "แยกตาม Location", [l for _, l in locs], lambda x: x["loc"])
    ws.column_dimensions["A"].width = 16
    for c in "BCDEF":
        ws.column_dimensions[c].width = 13

    head = ["#", "Area", "Location", "Days", "Container No", "Size/Type", "Group",
            "Move Code", "EQ Date", "Built Year", "RF Brand"]
    widths = [6, 8, 11, 8, 16, 10, 8, 11, 12, 11, 13]
    for name, lo, hi, fill_hex, font_hex in DAY_BUCKETS:
        sh = wb.create_sheet(name)
        sh.append(head)
        style_head(sh[1])
        fill = PatternFill("solid", fgColor=fill_hex)
        for i, r in enumerate(sorted(in_bucket[name], key=lambda x: -x["d"]), 1):
            sh.append([i, r["ar"], r["loc"], r["d"], r["cn"], r["st"], r["gp"],
                       r["mc"], r["eq"], r["by"] or "", r["rf"]])
            for c in sh[i + 1]:
                c.fill = fill
                c.border = border
            sh.cell(i + 1, 4).font = Font(bold=True, color=font_hex)
        for c, w in zip("ABCDEFGHIJK", widths):
            sh.column_dimensions[c].width = w
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = f"A1:K{len(in_bucket[name]) + 1}"
        sh.sheet_properties.tabColor = font_hex if name != "200+ วัน" else "FDE047"
    wb.save(path)
    return {b[0]: len(in_bucket[b[0]]) for b in DAY_BUCKETS}


def write_gateout_xlsx(rows, day, source, path):
    """รายงานตู้ปล่อยออก (OEP/OEV/OEW): ชีทสรุป + แยกชีทตามกลุ่มวันค้างก่อนปล่อย"""
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Border, Font, PatternFill, Side

    thin = Side(style="thin", color="BFC5CE")
    border = Border(top=thin, left=thin, bottom=thin, right=thin)

    def style_head(cells):
        for c in cells:
            c.font = Font(bold=True, color="FFFFFF")
            c.fill = PatternFill("solid", fgColor="1F5FBF")
            c.alignment = Alignment(horizontal="center")
            c.border = border

    known = [r for r in rows if r["d"] is not None]
    in_bucket = {b[0]: [r for r in known if b[1] <= r["d"] <= b[2]] for b in DAY_BUCKETS}
    unknown = [r for r in rows if r["d"] is None]

    wb = Workbook()
    ws = wb.active
    ws.title = "สรุป"
    ws["A1"] = f"รายงานตู้ปล่อยออกจากลาน (OEP / OEV / OEW) — {day}"
    ws["A1"].font = Font(bold=True, size=14)
    avg = round(sum(r["d"] for r in known) / len(known), 1) if known else 0
    long_n = sum(1 for r in known if r["d"] >= 100)
    ws["A2"] = (f"ไฟล์: {source} · ปล่อยออก {len(rows):,} ตู้ · วันค้างเฉลี่ย {avg} วัน · "
                f"ค้าง ≥100 วัน {long_n:,} ตู้" + (f" · ไม่พบวันเข้าลาน {len(unknown)} ตู้" if unknown else ""))
    ws["A2"].font = Font(color="6B7684")

    def table(start_row, title, keys, key_fn):
        ws.cell(start_row, 1, title).font = Font(bold=True, size=12)
        r0 = start_row + 1
        hdr = [""] + [b[0] for b in DAY_BUCKETS] + ["รวม"]
        for j, h in enumerate(hdr, 1):
            ws.cell(r0, j, h)
        style_head(ws[r0][:len(hdr)])
        for b, cell in zip(DAY_BUCKETS, ws[r0][1:len(DAY_BUCKETS) + 1]):
            cell.fill = PatternFill("solid", fgColor=b[3])
            cell.font = Font(bold=True, color=b[4])
        for i, k in enumerate(keys + ["รวม"], 1):
            rr = r0 + i
            vals = [sum(1 for x in in_bucket[b[0]] if k == "รวม" or key_fn(x) == k) for b in DAY_BUCKETS]
            for j, v in enumerate([k] + vals + [sum(vals)], 1):
                c = ws.cell(rr, j, v)
                c.border = border
                if k == "รวม" or j in (1, len(hdr)):
                    c.font = Font(bold=True)
                if k == "รวม":
                    c.fill = PatternFill("solid", fgColor="EEF1F5")
        return r0 + len(keys) + 3

    nxt = table(4, "แยกตาม Move (การปล่อย)", sorted({r["mv"] for r in known}), lambda x: x["mv"])
    nxt = table(nxt, "แยกตาม Group ก่อนปล่อย (AV / DMG)", ["AV", "DMG", "OTHER"], lambda x: x["pgp"])
    nxt = table(nxt, "แยกตาม Size/Type", sorted({r["st"] for r in known}), lambda x: x["st"])
    table(nxt, "แยกตาม Location", sorted({r["loc"] for r in known}), lambda x: x["loc"])
    ws.column_dimensions["A"].width = 18
    for c in "BCDEF":
        ws.column_dimensions[c].width = 13

    head = ["#", "Gate Date", "Time", "Move", "Area", "Location", "Days ค้าง", "Container No", "Size/Type",
            "EQ Date (เข้าลาน)", "Move ก่อนปล่อย", "Group", "Built Year", "POD", "Voyage"]
    widths = [6, 12, 7, 7, 7, 10, 9, 16, 10, 16, 14, 8, 10, 9, 9]
    sheets = [(b[0], in_bucket[b[0]], b[3], b[4]) for b in DAY_BUCKETS]
    if unknown:
        sheets.append(("ไม่พบวันเข้าลาน", unknown, "EEF1F5", "6B7684"))
    for name, lst, fill_hex, font_hex in sheets:
        sh = wb.create_sheet(name)
        sh.append(head)
        style_head(sh[1])
        fill = PatternFill("solid", fgColor=fill_hex)
        for i, r in enumerate(lst, 1):
            sh.append([i, r["gd"], r["gt"], r["mv"], r["ar"], r["loc"], r["d"] if r["d"] is not None else "",
                       r["cn"], r["st"], r["eq"], r["pmc"], r["pgp"], r["by"] or "", r["pod"], r["voy"]])
            for c in sh[i + 1]:
                c.fill = fill
                c.border = border
            sh.cell(i + 1, 7).font = Font(bold=True, color=font_hex)
        for col, w in zip("ABCDEFGHIJKLMNO", widths):
            sh.column_dimensions[col].width = w
        sh.freeze_panes = "A2"
        sh.auto_filter.ref = f"A1:O{len(lst) + 1}"
        sh.sheet_properties.tabColor = "FDE047" if name == "200+ วัน" else font_hex
    wb.save(path)
    counts = {b[0]: len(in_bucket[b[0]]) for b in DAY_BUCKETS}
    if unknown:
        counts["ไม่พบวันเข้าลาน"] = len(unknown)
    return counts


def archive(src, rows, meta):
    """เก็บไฟล์ของรอบนี้ไว้ในโฟลเดอร์วันที่ เช่น EQ STAYING/2026-10-05/
    (ถ้าไฟล์ต้นฉบับอยู่ในโฟลเดอร์วันที่อยู่แล้ว จะใช้โฟลเดอร์นั้น)"""
    parent = os.path.dirname(os.path.abspath(src))
    if os.path.dirname(parent) == HERE and re.fullmatch(r"\d{4}-\d{2}-\d{2}", os.path.basename(parent)):
        day = os.path.basename(parent)
    else:
        day = datetime.now().strftime("%Y-%m-%d")
    folder = os.path.join(HERE, day)
    os.makedirs(folder, exist_ok=True)
    # ไฟล์ต้นฉบับ: ย้ายเข้าโฟลเดอร์วัน (ถ้ายังไม่อยู่ในนั้น)
    dest_src = os.path.join(folder, os.path.basename(src))
    if os.path.abspath(src) != os.path.abspath(dest_src):
        shutil.move(src, dest_src)
    shutil.copy2(OUT, os.path.join(folder, f"EQ_STAYING_Dashboard_{day}.html"))
    counts = write_summary_xlsx(rows, meta, os.path.join(folder, f"EQ_STAYING_Summary_{day}.xlsx"))
    print(f"Archive: {folder}")
    print(f"  - {os.path.basename(src)} (ไฟล์ต้นฉบับ)")
    print(f"  - EQ_STAYING_Dashboard_{day}.html")
    print(f"  - EQ_STAYING_Summary_{day}.xlsx  " + " · ".join(f"{k}: {v:,}" for k, v in counts.items()))


if __name__ == "__main__":
    main()
