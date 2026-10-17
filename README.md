# EQ STAYING Dashboard — HEUNG A LINE

Dashboard ติดตามตู้ AV / DMG, LONG STAYING และตู้ปล่อยออกจากลาน (BKK / LCH)

แท็บในหน้าเว็บ
- **📊 ภาพรวม** — KPI, กราฟอายุตู้ (Aging) ตาม Area, สรุป Area × Group, Aging ตาม Location
- **⚠ LONG STAYING** — ตู้ ≥100 วัน แบ่งกลุ่มละ 50 วัน, ค้นหาเบอร์ตู้, Export Excel
- **🚚 ตู้ปล่อยออก** — ตู้ที่ปล่อยออก (OEP / OEV / OEW) พร้อมวันค้างก่อนปล่อย, สรุปรายวัน, Export Excel
- **📋 รายการตู้** — ตู้ทั้งหมดในลาน, ค้นหา, Export CSV

เกณฑ์
- **AV** = Move Code IEC, IED, IEP, IER, IEW, VED · **DMG** = OER · **OTHER** = Move Code อื่น
- **Days**: 0–30 เขียว · 31–99 ส้ม · 100–199 แดง · 200+ แดงไฮไลท์เหลือง
- **Area**: แบ่งจาก Location โดย LCH54 / LCH55 นับเป็น BKK
- **วันค้างก่อนปล่อย** = Gate Date (ไฟล์ GATE OUT) − EQ Date (จากไฟล์ STAYING ที่เก็บไว้ วันเดียวกันหรือก่อนหน้า)
- ตู้ในรายการ STAYING ที่ปล่อยออกแล้วจะมีป้าย "ปล่อยแล้ว"

## ลิงก์
- Dashboard: https://sirichai1265.github.io/eq-staying/ — กรองล่วงหน้าได้ เช่น `?area=BKK&group=AV&bucket=LONG&size=REEFER`, เปิดแท็บตรง `#long` `#gate` `#list`
- หน้ารวม EQUIPMENT: https://sirichai1265.github.io/equipment/ (repo `equipment`)

## อัปเดตข้อมูล
1. `python build_dashboard.py "<ไฟล์ STAYING.xls>" "<ไฟล์ GATE OUT.xls>"` (ใส่ไฟล์เดียวหรือทั้งสองไฟล์ก็ได้)
   - ไฟล์จะถูกย้ายไปเก็บในโฟลเดอร์วันที่ เช่น `2026-10-15/` พร้อมรายงาน Excel
     (`EQ_STAYING_Summary_<วัน>.xlsx`, `GATE_OUT_Report_<วัน>.xlsx`)
2. `git add index.html && git commit -m "Update data" && git push`
