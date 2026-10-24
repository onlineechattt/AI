# รายการที่ต้องตั้งค่าก่อนเปิดขายจริง

- [ ] สำรอง Apps Script เดิม
- [ ] แทน Code.gs ด้วยไฟล์ในชุด และเพิ่ม BoothPOS_V11.gs ในโปรเจกต์เดิม
- [ ] ตั้ง POS_API_KEY, POS_TEAM_PIN, POS_GROUP_ID และ POS_UPLOAD_FOLDER_ID
- [ ] Run setupLilbullePOSV11 และ Deploy Web App เวอร์ชันใหม่
- [ ] นำ URL /exec ไปตั้ง APPS_SCRIPT_URL ใน Netlify
- [ ] ตั้ง POS_API_KEY และ POS_SESSION_SECRET บน Netlify แล้ว deploy ทั้งโปรเจกต์
- [ ] เชิญ LINE OA เข้ากลุ่มยอด POS และตรวจ Group ID ให้ตรง
- [ ] ตั้งรูปและราคาสินค้าจริง
- [ ] สร้างใบเบิกทดสอบใหม่ ตรวจรับ ขาย ยกเลิก ปิดวัน คืนของ แล้วเทียบสต็อกและเงิน

ยังไม่ต้องแก้ชีทด้วยมือหรือย้ายข้อมูลบิลเก่า
ไม่ส่ง PIN / channel token / API key ลงแชท
