# EQ STAYING Dashboard — HEUNG A LINE

Dashboard ติดตามตู้ AV / DMG และ LONG STAYING (BKK / LCH)

- **AV** = Move Code IEC, IED, IEP, IER, IEW, VED · **DMG** = OER · **OTHER** = Move Code อื่น
- **Days**: 0–30 เขียว · 31–99 ส้ม · 100–199 แดง · 200+ แดงไฮไลท์เหลือง
- **Area**: แบ่งจาก Location โดย LCH54 / LCH55 นับเป็น BKK
- ตัวกรอง Size/Type, Built Year, Location, Move Code, Lessor และ RF Brand (22RE / 45RE)
- คลิกเบอร์ตู้ หรือพิมพ์ในช่องค้นหาแล้วกด Enter เพื่อดูรายละเอียด · ลิงก์ตรงตัวกรองได้ (ดูด้านล่าง)

## ลิงก์
- Dashboard: https://sirichai1265.github.io/eq-staying/ — กรองล่วงหน้าได้ เช่น `?area=BKK&group=AV&bucket=LONG&size=REEFER`
- หน้ารวม EQUIPMENT: https://sirichai1265.github.io/equipment/ (repo `equipment`)

## อัปเดตข้อมูล
1. วางไฟล์ `*STAYING*.xls` ใหม่ไว้ในโฟลเดอร์นี้
2. `python build_dashboard.py` (สร้าง `index.html` ใหม่)
3. `git add index.html && git commit -m "Update data" && git push`
