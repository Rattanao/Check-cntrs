# Check-cntrs

โปรแกรมตรวจเทียบ **MANIFEST** กับ **CNTRS** (Container List for Feeder) ของ Heung-A Line
ก่อนทำใบขนขาเข้า แล้วออกรายงาน Excel สีเขียว ✓ / แดง ⚠ ที่ `output/CNTRS_EDI.xlsx`

## ตรวจอะไรบ้าง

| หัวข้อ | วิธีตรวจ |
|---|---|
| **SHED NUMBER** | เทียบ MANIFEST กับ CNTRS ตามที่เห็นจริง (ไม่ตัดเลข 0 นำหน้า, CNTRS แสดง 4 หลัก) และตรวจกฎ SHED ที่ล็อกไว้: DG = 2826, ปลายทาง THBMT = 0110, ปลายทาง THLKR (สินค้าทั่วไป) = 0332 |
| **STATUS** | CY/FCL = FCL, LCL/LCL(AGENT)/LCL-CFS = LCL |
| **TEMP** | ตู้เย็น (ประเภท xxRx เช่น 22R1/45R1) เทียบอุณหภูมิ MANIFEST กับ `REEFER` ใน CNTRS |
| **VENT** | ตู้เย็นต้องมี `VENT:...` ใน MANIFEST ถ้าไม่มีขึ้น ❗ |
| **DG (CLASS/UN)** | เทียบ UN NUMBER เป็นหลัก, CLASS เทียบเลขหลักก่อน `+` |
| **REMARK** | ตามปลายทาง: THBMT = BY BARGE, SCT = BY TRUCK, THLKR = BY TRUCK (ตู้เย็น/DG) นอกนั้น BY TRAIN |

- ตรวจทุก Port of Discharge ยกเว้น BKK และ UCT (แก้ที่ `EXCLUDED_PORTS`)
- ช่องที่ตรงกันแสดง `-` ช่องที่ไม่ตรงแสดงค่าจริงพร้อมพื้นแดง และอธิบายในคอลัมน์ NOTE
- ตู้ที่มีในฝั่งเดียว (ไม่พบใน MANIFEST/CNTRS) ขึ้นแถวสีเหลือง

## ใช้งานผ่านเว็บ (ไม่ต้องติดตั้งอะไร)

เปิด **https://rattanao.github.io/Check-cntrs/** ลากไฟล์ MANIFEST และ CNTRS (`.xls` / `.xlsx` / `.pdf` วางพร้อมกันได้หลายไฟล์)
แล้วกด **ตรวจสอบ** ดูผลบนหน้าเว็บ ค้นหา/กรองเฉพาะที่ไม่ตรง และดาวน์โหลด `CNTRS_EDI.xlsx` ได้เลย
ไฟล์ถูกประมวลผลในเบราว์เซอร์ของคุณเท่านั้น ไม่มีการอัปโหลดไปที่ใด

หน้าเว็บคือ `index.html` + `core.js` (ตรรกะเดียวกับ `check_cntrs.py`) เปิดใช้งานด้วย GitHub Pages
(Settings → Pages → Branch `main` / root)

## วิธีใช้แบบโปรแกรม Python


1. ติดตั้ง Python 3.10+ แล้ว
   ```bash
   pip install -r requirements.txt
   ```
2. วางไฟล์ในโฟลเดอร์ `input/`
   - **MANIFEST**: `.xls` / `.xlsx` (รายงาน CARGO MANIFEST)
   - **CNTRS**: `.xls` / `.xlsx` หรือ `.pdf` (Container List for Feeder) วางได้หลายไฟล์
     (เช่น แยกตาม BMT/LCH/LKR) โปรแกรมรวมให้เป็นชุดเดียว
   - ตั้งชื่อไฟล์อะไรก็ได้ โปรแกรมแยกประเภทจากเนื้อหาในไฟล์
3. รัน
   ```bash
   python check_cntrs.py
   ```
4. เปิดผลที่ `output/CNTRS_EDI.xlsx`

## Skill สำหรับ Claude Code

โฟลเดอร์ [`.claude/skills/check-cntrs/`](.claude/skills/check-cntrs/SKILL.md) เป็น skill ระดับโปรเจกต์
เปิด Claude Code ในโฟลเดอร์นี้แล้วพิมพ์ **"ตรวจสอบ CNTRS"** (หรือ "เช็คใหม่") Claude จะรันโปรแกรมให้และสรุปผล

## หมายเหตุ

- `input/` และ `output/` ถูกตั้งเป็น ignore ไม่เอาข้อมูลลูกค้าขึ้น GitHub
- MANIFEST/CNTRS เป็นรายงานที่แปลงมาจากงานพิมพ์ ตำแหน่งคอลัมน์ไม่คงที่ โปรแกรมจึงหาข้อมูลจากข้อความ/หัวตารางแทนเลขคอลัมน์ตายตัว
