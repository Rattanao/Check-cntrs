---
name: check-cntrs
description: >
  ตรวจเทียบ MANIFEST กับ CNTRS (.xls/.xlsx/.pdf ใน input/) ด้วย check_cntrs.py แล้วออก
  output/CNTRS_EDI.xlsx (SHED, STATUS, TEMP, VENT, DG CLASS/UN, REMARK). ใช้ทุกครั้งที่ผู้ใช้พูดว่า
  "ตรวจสอบ CNTRS", "เช็ค CNTRS", "check CNTRS", "เช็คใหม่", "ตรวจใหม่", "เทียบ MANIFEST กับ CNTRS",
  "ทำ EDI" หรือเปลี่ยน/วางไฟล์ใหม่ใน input/ แล้วให้รันซ้ำ
---

# ตรวจสอบ MANIFEST vs CNTRS

1. ดูไฟล์ใน `input/` ว่ามี MANIFEST และ CNTRS (xls/xlsx/pdf ได้หลายไฟล์) ถ้าขาดฝั่งใด บอกผู้ใช้ให้วางไฟล์
   (โปรแกรมแยกประเภทจากเนื้อหา ไม่ดูชื่อไฟล์)
2. รัน `python check_cntrs.py` (ต้องมี `pip install -r requirements.txt`)
3. เปิด `output/CNTRS_EDI.xlsx` อ่านคอลัมน์ NOTE สรุปตู้ที่ไม่ตรงให้ผู้ใช้ แล้วส่งไฟล์
4. ถ้าจำนวนตู้ไม่ตรงกันสูงผิดปกติ (เช่น > 50%) ให้สงสัยว่าอ่านไฟล์ผิดรูปแบบ (คอลัมน์เลื่อน) ก่อน
   แล้วตรวจข้อมูลดิบ ไม่ใช่รายงานทันที

## กฎที่โปรแกรมใช้ (รายละเอียดใน README.md)
- SHED เทียบตามที่เห็น + กฎล็อก DG=2826, THBMT=0110, THLKR ทั่วไป=0332
- REMARK ตามปลายทาง: BMT=BY BARGE, SCT=BY TRUCK, LKR=BY TRUCK (ตู้เย็น/DG) นอกนั้น BY TRAIN
- ตู้เย็น (xxRx) ต้องมี VENT ใน MANIFEST
- ตรวจทุก PORT ยกเว้น BKK, UCT
- ช่องที่ตรง = "-", ไม่ตรง = พื้นแดง
- ห้ามเอาไฟล์ใน input/ output/ ขึ้น GitHub (ข้อมูลลูกค้า)
