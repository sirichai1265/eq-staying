# EQ STAYING Dashboard — HEUNG A LINE

Dashboard ติดตามตู้ AV / DMG และ LONG STAYING (BKK / LCH)

- **AV** = Move Code IEC, IED, IEP, IER, IEW, VED · **DMG** = OER · **OTHER** = Move Code อื่น
- **Days**: 0–30 เขียว · 31–99 ส้ม · 100–199 แดง · 200+ แดงไฮไลท์เหลือง
- **Area**: แบ่งจาก Location โดย LCH54 / LCH55 นับเป็น BKK
- ตัวกรอง Size/Type, Built Year, Location, Move Code, Lessor และ RF Brand (22RE / 45RE)
- คลิกเบอร์ตู้ หรือพิมพ์ในช่องค้นหาแล้วกด Enter เพื่อดูรายละเอียด · Dashboard อยู่ที่ `dashboard.html`

## หน้ารวม EQUIPMENT
`index.html` (หน้าแรก) — การ์ดสรุปตัวเลข (EQ Staying, Long Staying, Reefer, AV ตาม Area) กด **Open Dashboard** เพื่อเปิด Dashboard `dashboard.html` พร้อมตัวกรอง
(รองรับลิงก์ เช่น `dashboard.html?area=BKK&group=AV&bucket=LONG&size=REEFER`) · เพิ่มการ์ดใหม่ได้ที่ `CARDS` ใน `equipment_template.html`

## อัปเดตข้อมูล
1. วางไฟล์ `*STAYING*.xls` ใหม่ไว้ในโฟลเดอร์นี้
2. `python build_dashboard.py` (สร้าง `index.html` และ `dashboard.html` ใหม่)
3. `git add index.html dashboard.html && git commit -m "Update data" && git push`
