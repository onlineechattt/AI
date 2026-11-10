LIL BULLE — ชุดรวมอัปเดตคลัง V11.3 / POS V12.3 / เมนูโปรโมชั่น
สถานะ: จัดเตรียมโค้ดและแพตช์แล้ว ยังไม่ได้ติดตั้งใน Google Apps Script หรือ Netlify จริง
ห้ามเปลี่ยนระบบระหว่างบูทกำลังขาย

ไฟล์ติดตั้ง
1. LilBulle_Code_V11.3_WAREHOUSE_LOT.gs -> ใช้แทน Code.gs ใน Apps Script
2. LilBulle_BoothPOS_V12.3_PROMOTIONS.gs -> ใช้แทน BoothPOS_V11.gs ใน Apps Script
3. patch_repo.py -> ใช้ปรับหน้าเว็บ public/app.js และ netlify/functions/pos.mjs ใน repository onlineechattt/AI
4. LilBulle_V11.3_V12.2_UPDATE_GUIDE.txt -> กติกาและการทดสอบระบบคลัง/บูท

กติกา
- POS ไม่ตัด Stock คลังกลาง เมื่อขายหรือเติมสินค้า
- เบิกจริงจากคลังผ่านระบบคลังเท่านั้น
- คืนสินค้าเข้าคลังต้องแยก SKU + EXP ตามใบเบิกสะสม-ใบคืนสะสม ไม่หัก POS ซ้ำ
- บอสจัดการชื่อสินค้าในชุด จำนวน ราคา เปิด-ปิดโปรโมชั่นผ่านเว็บ
- พนักงานเลือกใช้โปรโมชั่นได้ แต่ไม่แก้ไข
- Sheet ใหม่: LilbullePOS_Promotions สร้างอัตโนมัติเมื่อใช้งาน
- บิลเดิมไม่ถูกเปลี่ยนย้อนหลัง

ติดตั้งแบบทดสอบก่อน production
A. สำรอง Code.gs, BoothPOS_V11.gs และ GitHub main ก่อน
B. ทำโปรเจกต์/Google Sheet สำเนาสำหรับทดสอบ
C. แทน 2 ไฟล์ Apps Script ในโปรเจกต์ทดสอบแล้วบันทึก
D. clone repo onlineechattt/AI มาไว้ในเครื่อง และรันคำสั่ง: python patch_repo.py /path/to/AI
   โดยคัดลอก patch_repo.py จาก ZIP ไปวางตำแหน่งใดก็ได้
E. ตรวจ Node syntax ของ public/app.js / netlify/functions/pos.mjs และ Deploy staging
F. ทดสอบ Login boss / team, สร้างและปิดโปรโมชั่น, สร้างบิล, เปิดปิดบูท และตรวจ Stock/EventIssue ไม่เปลี่ยนจาก POS
G. หลังผ่านครบค่อย Deploy Apps Script เป็น New version และ GitHub/Netlify ในช่วงไม่มีการขาย

ข้อควรรู้
- patch_repo.py อิง source github.com/onlineechattt/AI/main ที่ตรวจวันที่ 10 พ.ย. 2026 หากรูปแบบไฟล์เปลี่ยนจะหยุด ไม่แก้มั่ว
- หน้าสร้างโปรโมชั่นรองรับ 5 SKU ต่อชุด, backend 10 SKU
- การเปลี่ยนราคา/องค์ประกอบโปรโมชันไม่แก้บิลเก่า
- ยังไม่ได้ deploy หรือทดสอบการทำงานจริงบน Netlify/Apps Script
- หากไม่สามารถ run Python บน iPad ต้องใช้ GitHub workflow หรือเครื่องคอมพิวเตอร์เพื่อใช้ patch_repo.py
