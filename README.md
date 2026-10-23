# EQ STAYING Dashboard — HEUNG A LINE

Dashboard ติดตามตู้ AV / DMG, LONG STAYING และตู้ปล่อยออกจากลาน (BKK / LCH)

แท็บในหน้าเว็บ
- **📊 ภาพรวม** — KPI, กราฟอายุตู้ (Aging) ตาม Area, สรุป Area × Group, Aging ตาม Location
- **⚠ LONG STAYING** — ตู้ ≥100 วัน แบ่งกลุ่มละ 50 วัน, ค้นหาเบอร์ตู้, Export Excel
- **🚚 ตู้ปล่อยออก** — GATE OUT จากรีพอตลาน (`GATE MOVE\GATE MOVE DATABASE.xlsx`) ตั้งแต่ 1/10 พร้อมวันค้างก่อนปล่อย · เลือกสัปดาห์ (ISO จันทร์–อาทิตย์) / ช่วงวันที่ / ลาน · สรุปแต่ละลาน · สรุปรายวัน/รายสัปดาห์ · Export Excel

เกณฑ์
- **AV** = Move Code IEC, IED, IEP, IER, IEW, VED · **DMG** = OER · **OTHER** = Move Code อื่น
- **Days**: 0–30 เขียว · 31–99 ส้ม · 100–199 แดง · 200+ แดงไฮไลท์เหลือง
- **Area**: แบ่งจาก Location โดย LCH54 / LCH55 นับเป็น BKK
- **วันค้างก่อนปล่อย** = Gate Date − EQ Date (จากไฟล์ STAYING รายวันใน `STOCK DAILY` / โฟลเดอร์วันที่ ฉบับล่าสุดที่ไม่เกินวันปล่อย)
- ช่วงวันที่ของตู้ปล่อยออกกำหนดที่ `GATE_FROM` / `GATE_TO` ใน `build_dashboard.py`
- ตู้ในรายการ STAYING ที่ปล่อยออกแล้วจะมีป้าย "ปล่อยแล้ว"

## ลิงก์
- Dashboard: https://sirichai1265.github.io/eq-staying/ — กรองล่วงหน้าได้ เช่น `?area=BKK&group=AV&bucket=LONG&size=REEFER`, เปิดแท็บตรง `#long` `#gate` `#list`
- หน้ารวม EQUIPMENT: https://sirichai1265.github.io/equipment/ (repo `equipment`)

## อัปเดตข้อมูล
1. `python build_dashboard.py "<ไฟล์ STAYING.xls>" "<ไฟล์ GATE OUT.xls>"` (ใส่ไฟล์เดียวหรือทั้งสองไฟล์ก็ได้)
   - ไฟล์จะถูกย้ายไปเก็บในโฟลเดอร์วันที่ เช่น `2026-10-15/` พร้อมรายงาน Excel
     (`EQ_STAYING_Summary_<วัน>.xlsx`, `GATE_OUT_Report_<วัน>.xlsx`)
2. `git add index.html && git commit -m "Update data" && git push`

## ตัดยอดตู้ปล่อยออก (อัตโนมัติหลังรัน GATE MOVE)
- `GATE MOVE\yard_gate_summary.py` เรียก `build_dashboard.py` ต่อท้ายให้เอง (ยกเว้นโหมด `--check`)
- ใช้ไฟล์ STAYING ตั้งต้นเดิม (`.current_staying`) แล้วตัดตู้ที่ปล่อยออกแล้วตาม GATE MOVE DATABASE
- ผลลัพธ์: `GATE OUT\GATE_OUT_Summary.xlsx` (รายสัปดาห์ / รายวัน + ยอดคงเหลือ / แต่ละลาน / รายการ) และ `<วันที่>\GATE_OUT_<วันที่>.xlsx`
- อัปเดตหน้าเว็บออนไลน์: `git add index.html && git commit -m "Update gate out" && git push`

