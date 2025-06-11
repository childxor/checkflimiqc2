# วิธีติดตั้ง ClosedXML สำหรับค้นหาข้อมูลใน Excel

## วิธีที่ 1: ติดตั้งผ่าน NuGet Package Manager
1. เปิดโปรเจ็คใน Visual Studio
2. คลิกขวาที่โปรเจ็ค > Manage NuGet Packages...
3. ไปที่แท็บ "Browse"
4. ค้นหา "ClosedXML"
5. เลือกแพคเกจ "ClosedXML" (ไม่ใช่ ClosedXML.Excel)
6. คลิก "Install"
7. ยอมรับเงื่อนไขและการเปลี่ยนแปลง

## วิธีที่ 2: ติดตั้งผ่าน Package Manager Console
1. เปิด Package Manager Console ใน Visual Studio (Tools > NuGet Package Manager > Package Manager Console)
2. พิมพ์คำสั่ง: `Install-Package ClosedXML`
3. กด Enter

## การใช้งาน ClosedXML
หลังจากติดตั้งแล้ว ระบบจะสามารถใช้ ClosedXML ค้นหาข้อมูลใน Excel ได้โดยอัตโนมัติเมื่อไม่มี Microsoft Office ติดตั้ง

## ข้อดีของ ClosedXML
- ไม่จำเป็นต้องติดตั้ง Microsoft Office บนเครื่อง
- ทำงานได้เร็วกว่า Excel Interop
- ใช้งานง่าย มี API ที่เข้าใจได้ง่าย
- เป็น Open Source และได้รับการพัฒนาอย่างต่อเนื่อง

## หมายเหตุ
ถ้าใช้ฟีเจอร์ซับซ้อนของ Excel เช่น Macro, VBA, Chart ควรใช้ Excel Interop แทน 