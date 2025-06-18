Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Xml

''' <summary>
''' คลาสสำหรับจัดการฐานข้อมูล (Wrapper สำหรับ AccessDatabaseManager)
''' </summary>
Public Class DatabaseManager

    ' กำหนดพาธของไฟล์การตั้งค่า
    Private Shared ReadOnly CONFIG_FILE As String = "Settings.config"

    ' ตัวแปรสำหรับการเชื่อมต่อฐานข้อมูล Access
    Private Shared _databasePath As String = "QRCodeScanner.accdb"
    Private Shared _password As String = ""

    ' Connection string สำหรับเชื่อมต่อฐานข้อมูล Access
    Private Shared ReadOnly Property ConnectionString As String
        Get
            Dim connectionStrings As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath};"

            If Not String.IsNullOrEmpty(_password) Then
                connectionStrings += $"Jet OLEDB:Database Password={_password};"
            End If

            Return connectionStrings
        End Get
    End Property

    ''' <summary>
    ''' โหลดการตั้งค่าฐานข้อมูลจากไฟล์ config
    ''' </summary>
    Private Shared Sub LoadDatabaseSettings()
        Try
            If File.Exists(CONFIG_FILE) Then
                Dim doc As New XmlDocument()
                doc.Load(CONFIG_FILE)

                _databasePath = GetSettingValueFromXML(doc, "AccessDatabasePath", "QRCodeScanner.accdb")
                _password = GetSettingValueFromXML(doc, "AccessPassword", "")

                Console.WriteLine($"Database settings loaded - Path: {_databasePath}")
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading database settings: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ดึงค่าจาก XML
    ''' </summary>
    Private Shared Function GetSettingValueFromXML(doc As XmlDocument, key As String, defaultValue As Object) As Object
        Try
            Dim node As XmlNode = doc.SelectSingleNode($"//Setting[@key='{key}']")
            If node IsNot Nothing Then
                Dim value As String = node.Attributes("value").Value

                Select Case defaultValue.GetType()
                    Case GetType(Boolean)
                        Return Boolean.Parse(value)
                    Case GetType(Integer)
                        Return Integer.Parse(value)
                    Case GetType(Decimal)
                        Return Decimal.Parse(value)
                    Case Else
                        If key = "AccessPassword" Then
                            Return DecryptPassword(value)
                        End If
                        Return value
                End Select
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    ''' <summary>
    ''' ถอดรหัสพาสเวิร์ดแบบง่าย
    ''' </summary>
    Private Shared Function DecryptPassword(encryptedPassword As String) As String
        Try
            If String.IsNullOrEmpty(encryptedPassword) Then Return ""

            ' การถอดรหัสแบบง่าย (จาก Base64)
            Dim bytes As Byte() = Convert.FromBase64String(encryptedPassword)
            Return System.Text.Encoding.UTF8.GetString(bytes)
        Catch
            Return encryptedPassword
        End Try
    End Function

    ''' <summary>
    ''' ตรวจสอบการเชื่อมต่อฐานข้อมูล
    ''' </summary>
    ''' <returns>True ถ้าเชื่อมต่อได้, False ถ้าเชื่อมต่อไม่ได้</returns>
    Public Shared Function IsConnected() As Boolean
        Return AccessDatabaseManager.IsConnected()
    End Function

    ''' <summary>
    ''' เริ่มต้นการใช้งานฐานข้อมูล
    ''' </summary>
    ''' <returns>True ถ้าเริ่มต้นสำเร็จ, False ถ้าเริ่มต้นไม่สำเร็จ</returns>
    Public Shared Function Initialize() As Boolean
        Try
            ' โหลดการตั้งค่าฐานข้อมูลก่อน
            LoadDatabaseSettings()

            ' ส่งค่าการตั้งค่าไปยัง AccessDatabaseManager
            ' หมายเหตุ: ตอนนี้ใช้ NetworkPathManager แล้ว ไม่ต้องตั้งค่า path เอง
            AccessDatabaseManager.SetDatabasePassword(_password)

            Return AccessDatabaseManager.Initialize()
        Catch ex As Exception
            Console.WriteLine($"Error in DatabaseManager.Initialize: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' บันทึกข้อมูลการสแกน
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>ID ของรายการที่บันทึก</returns>
    Public Shared Function SaveScanData(record As ScanDataRecord) As Integer
        Return AccessDatabaseManager.SaveScanData(record)
    End Function

    ''' <summary>
    ''' เพิ่มข้อมูลการสแกนใหม่
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>ID ของรายการที่เพิ่ม</returns>
    Public Shared Function AddScanRecord(record As ScanDataRecord) As Integer
        Return AccessDatabaseManager.AddScanRecord(record)
    End Function

    ''' <summary>
    ''' ดึงข้อมูลประวัติการสแกนทั้งหมด
    ''' </summary>
    ''' <param name="limit">จำนวนรายการสูงสุดที่ต้องการดึง</param>
    ''' <returns>รายการข้อมูลการสแกน</returns>
    Public Shared Function GetScanHistory(Optional limit As Integer = 1000) As List(Of ScanDataRecord)
        Return AccessDatabaseManager.GetScanHistory(limit)
    End Function

    ''' <summary>
    ''' ลบข้อมูลการสแกนตาม ID
    ''' </summary>
    ''' <param name="id">ID ของรายการที่ต้องการลบ</param>
    ''' <returns>True ถ้าลบสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Public Shared Function DeleteScanRecord(id As Integer) As Boolean
        Return AccessDatabaseManager.DeleteScanRecord(id)
    End Function

    ''' <summary>
    ''' อัปเดตข้อมูลการสแกน
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกนที่อัปเดต</param>
    ''' <returns>True ถ้าอัปเดตสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Public Shared Function UpdateScanRecord(record As ScanDataRecord) As Boolean
        Return AccessDatabaseManager.UpdateScanRecord(record)
    End Function

    ''' <summary>
    ''' ค้นหาข้อมูลการสแกนตามรหัสผลิตภัณฑ์
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>รายการข้อมูลการสแกนที่ตรงกับรหัสผลิตภัณฑ์</returns>
    Public Shared Function SearchByProductCode(productCode As String) As List(Of ScanDataRecord)
        Return AccessDatabaseManager.SearchByProductCode(productCode)
    End Function

    ''' <summary>
    ''' ได้รับข้อมูลสถิติการใช้งาน
    ''' </summary>
    ''' <returns>ข้อมูลสถิติ</returns>
    Public Shared Function GetStatistics() As DatabaseStatistics
        Return AccessDatabaseManager.GetStatistics()
    End Function

    ''' <summary>
    ''' สำรองข้อมูลฐานข้อมูล
    ''' </summary>
    ''' <param name="backupPath">พาธที่จะสำรองข้อมูล</param>
    ''' <returns>True ถ้าสำรองสำเร็จ</returns>
    Public Shared Function BackupDatabase(backupPath As String) As Boolean
        Return AccessDatabaseManager.BackupDatabase(backupPath)
    End Function

    ''' <summary>
    ''' คืนค่าฐานข้อมูลจากไฟล์สำรอง
    ''' </summary>
    ''' <param name="backupPath">พาธของไฟล์สำรอง</param>
    ''' <returns>True ถ้าคืนค่าสำเร็จ</returns>
    Public Shared Function RestoreDatabase(backupPath As String) As Boolean
        Return AccessDatabaseManager.RestoreDatabase(backupPath)
    End Function

    ''' <summary>
    ''' ทำความสะอาดข้อมูลเก่า
    ''' </summary>
    ''' <param name="daysOld">จำนวนวันที่จะลบข้อมูลเก่า</param>
    ''' <returns>จำนวนรายการที่ลบ</returns>
    Public Shared Function CleanupOldData(daysOld As Integer) As Integer
        Return AccessDatabaseManager.CleanupOldData(daysOld)
    End Function
End Class

