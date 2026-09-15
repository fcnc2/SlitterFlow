# SlitterFlow — ทดลองใช้งาน

เปิดเว็บที่ https://fcnc2.github.io/SlitterFlow/ บน GitHub Pages โดยไม่ต้องเชื่อม Microsoft 365 หรือ SharePoint

| บทบาท | Username | Password |
| --- | --- | --- |
| Operator | operator | operator123 |
| Viewer | viewer | viewer123 |
| Admin | admin | admin123 |

Operator ทดลองตรวจเครื่องจักร/Dolly, กรอก Condition, เลือกใบมีด และกดยืนยันบันทึกได้ Viewer/Admin กดกรองก่อนอ่านรายการที่ทดลองบันทึกในเครื่องเดียวกันและ Export CSV ได้ Admin เพิ่มหรือยกเลิก Parameter แยก TH3/TH4 ได้

ข้อมูลรายการและ Parameter ใช้ `localStorage` ในเบราว์เซอร์เครื่องที่เปิดเว็บเท่านั้น ไม่ถูกส่งไปฐานข้อมูลและไม่แชร์ข้ามเครื่อง/เบราว์เซอร์ การล้างข้อมูลเว็บไซต์จะลบข้อมูลทดลอง บัญชีทดลองเป็นเพียงเมนูสำหรับลองหน้าจอ ไม่ใช่ระบบยืนยันตัวตนจริง หน้าจัดการผู้ใช้ยังเป็นตัวอย่างในหน่วยความจำของหน้าเว็บ

โค้ดเชื่อม SharePoint เดิมยังอยู่ใน repository สำหรับการพัฒนาครั้งถัดไป แต่หน้าเว็บทดลองไม่เรียกใช้งาน เส้นทาง build ของ GitHub Actions ใช้ `npm ci` และ `npm run build` จาก `src/main.tsx` เพื่อเผยแพร่เวอร์ชันปัจจุบัน
