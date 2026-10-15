# Balance Waste — ระบบบริหารจัดการของเสีย (Waste Management)

**บริษัท แน็ป นิวตริซายส์ (ประเทศไทย) จำกัด** · ทะเบียนโรงงาน 82260007625693 · หนังสืออนุญาต สก.2 เลขที่ 2569-8409 (ปี 2569)

ใช้ลงข้อมูลปริมาณของเสียที่จัดเก็บในโรงงาน: เลือกประเภทของเสีย กรอกน้ำหนัก (กก.) และภาชนะบรรจุ (ถัง 200 ลิตร / IBC / Bigbag)
แล้วติดตามยอดคงเหลือ การส่งกำจัด และการใช้โควตาตามปริมาณที่ได้รับอนุญาต

- รีโป: https://github.com/watanathep8-dotcom/balance-waste-nap
- หน้าเว็บ: GitHub Pages จากโฟลเดอร์ `docs/` → https://watanathep8-dotcom.github.io/balance-waste-nap/
- Backend: Google Apps Script (`apps-script/Code.gs` ไฟล์เดียว) เก็บข้อมูลใน Google Sheets

## ความสามารถ

| หน้า | รายละเอียด |
|---|---|
| แดชบอร์ด | KPI (คงเหลือรวม, จัดเก็บเดือนนี้, ส่งกำจัดปีนี้, % โควตา), เกจเข็มโควตารวม, วงแหวนโควตาแยกประเภท, กราฟ 12 เดือน, โดนัทภาชนะ, โควตาตามผู้รับ, แจ้งเตือน, รายการล่าสุด |
| บันทึกจัดเก็บ | เลือกประเภท (ค้นหาได้ แยกตามผู้รับ) · น้ำหนัก กก. (ทศนิยม 2 ตำแหน่ง) · ภาชนะ + จำนวน · สถานที่จัดเก็บ · หมายเหตุ · โหมด **ยอดยกมา** สำหรับของที่มีอยู่ก่อนเริ่มใช้ระบบ |
| ส่งกำจัด / นำออก | แสดงยอดคงเหลือสด · ส่งเกินยอดคงเหลือ/จำนวนภาชนะไม่ได้ · เลขที่ใบกำกับการขนส่ง · ทะเบียนรถ |
| คงเหลือ & โควตา | คงเหลือ = รับเข้า − ส่งกำจัด, ภาชนะในคลัง, % โควตาปีนี้ (เตือน ≥ 80 %, เกิน > 100 %), อายุล็อตเก่าสุด (FIFO, เตือนเกิน 90 วัน), สรุปตามผู้รับ |
| ประวัติรายการ | ตัวกรอง (ประเภท, รายการ, ช่วงวันที่, ผู้รับ, ค้นหา), ส่งออก Excel (CSV UTF-8 BOM), แก้ไข / ลบ (ต้องใช้รหัสผ่านผู้ดูแล) |
| บันทึกประวัติ (Log) | ทุกการกระทำ ใคร ทำอะไร เมื่อไร (ต้องใช้รหัสผ่านผู้ดูแล) |
| ตั้งค่าประเภทของเสีย | แก้ไข / เพิ่มประเภท (รหัส, ชื่อ, ปริมาณอนุญาต (ตัน), รหัสการจัดการ, ผู้รับ, เปิด/ปิดใช้งาน) และปุ่มรีเซ็ตข้อมูล (ต้องใช้รหัสผ่านผู้ดูแล) |

- ไม่มีระบบล็อกอิน: ต้องตั้ง **ชื่อผู้ใช้งานที่มุมขวาบน** ก่อนบันทึก — ชื่อถูกบันทึกเป็น “ผู้บันทึก” ทุกรายการ
- เซิร์ฟเวอร์ตรวจยอดย้อนหลังทุกครั้ง (ส่งกำจัด / แก้ไข / ลบ): น้ำหนักของทุกประเภท และจำนวนภาชนะแต่ละชนิด ต้องไม่ติดลบ ณ เวลาใดเลย
- ประเภทของเสีย 13 รายการตามหนังสืออนุญาต ปี 2569 (รวม 194 ตัน) ถูกสร้างอัตโนมัติตอน `setupSystem()` — รหัสเดียวกันที่ผู้รับต่างกันเป็นคนละรายการ
- ระบบ **ไม่สร้างยอดคงเหลือให้เอง** — ของเสียที่จัดเก็บอยู่ก่อนแล้ว (เช่น น้ำล้างเครื่องจักร, ตะกอนบำบัดน้ำเสีย) ให้บันทึกเป็น “ยอดยกมา” ตามที่ตรวจนับจริง

## ติดตั้ง (ครั้งแรก)

1. สร้างโปรเจกต์ใหม่ที่ https://script.google.com ตั้งชื่อโปรเจกต์ **Balance Waste Nap** แล้ววางเนื้อหา `apps-script/Code.gs`
   (และตั้งค่า `appsscript.json` ตามไฟล์ในรีโป: เขตเวลา Asia/Bangkok, V8)
2. **Project Settings → Script properties** เพิ่ม:

   | Property | ค่า |
   |---|---|
   | `WASTE_RESET_PASSWORD` | รหัสผ่านผู้ดูแล (ตั้งเอง ห้ามใส่ในโค้ด) — ใช้แก้ไข/ลบรายการ ดู Log ตั้งค่าประเภทของเสีย และรีเซ็ตข้อมูล |

   `WASTE_SPREADSHEET_ID` ถูกสร้างให้อัตโนมัติโดย `setupSystem()` (ไม่ต้องตั้งเอง)
3. ในตัวแก้ไข Apps Script เลือกฟังก์ชันแล้วกด **Run**:
   - `setupSystem` — สร้างสเปรดชีต “Balance Waste Database” (ชีต `waste_types`, `movements`, `audit_logs`), ใส่ประเภทของเสีย 13 รายการ และบันทึก ID ไว้ใน Script Properties (รันซ้ำได้ ไม่สร้างซ้ำ)
   - `installKeepWarmTrigger` — (แนะนำ) ตั้ง trigger ทุก 10 นาที ให้ `keepWarm()` อุ่นแคชการอ่าน หน้าเว็บจะเปิดเร็วขึ้น
4. **Deploy → New deployment → Web app**
   - Execute as: **Me**
   - Who has access: **Anyone**
5. คัดลอก URL ที่ลงท้ายด้วย `/exec` ใส่ใน `docs/config.js`:

   ```js
   window.WASTE_CONFIG = { apiUrl: "https://script.google.com/macros/s/XXXX/exec" };
   ```
6. เปิด GitHub Pages ของรีโป `watanathep8-dotcom/balance-waste-nap`: Settings → Pages → Branch `main` / folder `/docs`
   → https://watanathep8-dotcom.github.io/balance-waste-nap/

ฟังก์ชันอื่นที่รันจากตัวแก้ไขได้: `clearReadCache()` (หลังแก้สเปรดชีตด้วยมือ), `removeKeepWarmTrigger()`

## ทดสอบ / ดูตัวอย่างในเครื่อง

```bash
node test/gas-mock.js                    # ทดสอบ backend + frontend ทั้งหมด (ไม่ต่อ Google)
node test/dev-server.js 8777             # เปิด http://127.0.0.1:8777/ ด้วย Apps Script จำลองในหน่วยความจำ
node test/dev-server.js 8777 --demo      # เหมือนข้างบน + ข้อมูลตัวอย่าง (สร้างขึ้นเพื่อดูหน้าจอเท่านั้น ไม่ใช่ข้อมูลจริง)
```

dev-server ใช้รหัสผ่านผู้ดูแลจำลอง `dev-reset` (เฉพาะในเครื่อง)

## โครงสร้าง

```
apps-script/Code.gs      backend (doGet/doPost, LockService, CacheService, setupSystem, keepWarm)
appsscript.json          manifest
docs/index.html          หน้าเว็บ
docs/assets/app.js       frontend
docs/assets/style.css    ธีม (emerald / teal + amber), เอฟเฟกต์หมุนไอคอน, รองรับ prefers-reduced-motion
docs/config.js           apiUrl
test/gas-mock.js         mock ของ Apps Script + ชุดทดสอบ
test/dev-server.js       เซิร์ฟเวอร์ดูตัวอย่างในเครื่อง
```

API: อ่าน `GET ?action=…&_=timestamp` · เขียน `POST` แบบ `text/plain` (JSON) — ไม่เกิด CORS preflight

Programming by Duongruthai_Milk
