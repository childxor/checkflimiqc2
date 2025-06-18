# คู่มือการอัพเกรดระบบจัดการพาธ Network OA/FAB

## ภาพรวมการเปลี่ยนแปลง

โปรเจค CheckFlimIQC2 ได้รับการอัพเกรดเพื่อรองรับการทำงานกับ 2 เครือข่าย:
- **Network OA**: IP `10.24.179.2` → พาธ `\\10.24.179.2\OAFAB\OA2FAB`
- **Network FAB**: IP `172.24.0.3` → พาธ `\\172.24.0.3\OAFAB\OA2FAB`

## ไฟล์ที่ได้รับการปรับปรุง

### 1. ไฟล์ใหม่ที่สร้างขึ้น
- **`databaseManage/NetworkPathManager.vb`** - คลาส core สำหรับจัดการพาธ network
- **`frmMenu/NetworkTestForm.vb`** - ฟอร์มทดสอบการเชื่อมต่อ network

### 2. ไฟล์ที่ปรับปรุง
- **`databaseManage/AccessDatabaseManager.vb`** - ใช้ NetworkPathManager
- **`updateSystems/UpdateManager.vb`** - ใช้ NetworkPathManager
- **`frmHistory/frmHistory.vb`** - ใช้ NetworkPathManager
- **`excelExtention/ExcelUtility.vb`** - ใช้ NetworkPathManager
- **`Settings.config`** - เพิ่มการตั้งค่า network

## คุณสมบัติใหม่

### 1. NetworkPathManager
```vb
' ตรวจสอบการเชื่อมต่อ network
Dim networkResult = NetworkPathManager.CheckNetworkConnection()

' ได้รับพาธต่างๆ
Dim excelPath = NetworkPathManager.GetExcelDatabasePath()
Dim accessPath = NetworkPathManager.GetAccessDatabasePath()
Dim updatePath = NetworkPathManager.GetUpdateSystemPath()
```

### 2. การตรวจสอบ Network อัตโนมัติ
- ระบบจะ ping `10.24.179.2` ก่อน (Network OA)
- หากไม่ได้ผล จะ ping `172.24.0.3` (Network FAB)
- ใช้ timeout 3 วินาที

### 3. การจัดการพาธแบบไดนามิก
- พาธจะถูกสร้างตาม network ที่เชื่อมต่อได้
- ไม่ต้องเปลี่ยนโค้ดเมื่อเปลี่ยน network

## การทดสอบระบบ

เพื่อทดสอบระบบใหม่:

```vb
' เปิดฟอร์มทดสอบ
Dim testForm As New NetworkTestForm()
testForm.ShowDialog()
```

## การตั้งค่า Settings.config

```xml
<!-- Network Settings -->
<Setting key="OA_ServerIP" value="10.24.179.2" />
<Setting key="FAB_ServerIP" value="172.24.0.3" />
<Setting key="BaseSharePath" value="OAFAB\OA2FAB" />
<Setting key="PingTimeout" value="3000" />
```

## ข้อดีของระบบใหม่

### 1. ความยืดหยุ่น
- รองรับทั้ง 2 network อัตโนมัติ
- เปลี่ยน network ได้โดยไม่ต้องแก้โค้ด

### 2. ความเสถียร
- ตรวจสอบการเชื่อมต่อก่อนใช้งาน
- จัดการข้อผิดพลาดได้ดีขึ้น

### 3. ความสะดวก
- โค้ดสั้นลง ไม่ซ้ำซ้อน
- ง่ายต่อการบำรุงรักษา

## การแก้ไขปัญหา

### ปัญหา: ไม่สามารถเชื่อมต่อ network ได้
```
วิธีแก้:
1. ตรวจสอบ IP Address ของเซิร์ฟเวอร์
2. ตรวจสอบการเชื่อมต่อเครือข่าย
3. ตรวจสอบสิทธิ์การเข้าถึง Network Share
4. ใช้ NetworkTestForm เพื่อทดสอบ
```

### ปัญหา: ไม่พบไฟล์ที่ต้องการ
```
วิธีแก้:
1. ตรวจสอบว่าไฟล์มีอยู่จริงบนเซิร์ฟเวอร์
2. ตรวจสอบพาธที่ถูกต้อง
3. ตรวจสอบสิทธิ์การเข้าถึงไฟล์
```

## ตัวอย่างการใช้งาน

### 1. ตรวจสอบสถานะ network
```vb
Dim status = NetworkPathManager.GetNetworkStatus()
MessageBox.Show(status)
```

### 2. ตรวจสอบว่าไฟล์มีอยู่หรือไม่
```vb
Dim excelPath = NetworkPathManager.GetExcelDatabasePath()
If NetworkPathManager.PathExists(excelPath) Then
    ' ไฟล์มีอยู่
Else
    ' ไฟล์ไม่มี
End If
```

### 3. การใช้งานพาธกำหนดเอง
```vb
Dim customPath = NetworkPathManager.GetCustomPath("Film charecter check\Drawing")
If Not String.IsNullOrEmpty(customPath) Then
    ' ใช้งานพาธ
End If
```

## สรุป

การอัพเกรดนี้ทำให้โปรเจค CheckFlimIQC2 สามารถทำงานกับทั้ง 2 เครือข่ายได้อย่างราบรื่น โดยไม่ต้องเปลี่ยนแปลงโค้ดเมื่อเปลี่ยน network ระบบจะตรวจสอบและเลือกใช้พาธที่เหมาะสมโดยอัตโนมัติ

---
*อัพเดทล่าสุด: 2024-XX-XX* 