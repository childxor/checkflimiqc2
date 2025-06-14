Imports System

''' <summary>
''' คลาสสำหรับเก็บข้อมูลการสแกน QR Code
''' </summary>
Public Class ScanDataRecord
    Private _id As Integer
    Private _scanDateTime As DateTime
    Private _originalData As String
    Private _productCode As String
    Private _referenceCode As String
    Private _quantity As Integer
    Private _dateCode As String
    Private _isValid As Boolean
    Private _computerName As String
    Private _userName As String
    Private _extractedData As String
    Private _validationMessages As String
    Private _missionStatus As String = "ไม่มี" ' สถานะ Mission: "ไม่มี", "รอดำเนินการ", "สำเร็จ"

    ''' <summary>
    ''' ID ของรายการในฐานข้อมูล
    ''' </summary>
    Public Property Id() As Integer
        Get
            Return _id
        End Get
        Set(value As Integer)
            _id = value
        End Set
    End Property

    ''' <summary>
    ''' วันที่และเวลาที่สแกน
    ''' </summary>
    Public Property ScanDateTime() As DateTime
        Get
            Return _scanDateTime
        End Get
        Set(value As DateTime)
            _scanDateTime = value
        End Set
    End Property

    ''' <summary>
    ''' ข้อมูลดิบที่ได้จากการสแกน
    ''' </summary>
    Public Property OriginalData() As String
        Get
            Return _originalData
        End Get
        Set(value As String)
            _originalData = value
        End Set
    End Property

    ''' <summary>
    ''' ข้อมูลที่ดึงออกมาจากการสแกน
    ''' </summary>
    Public Property ExtractedData() As String
        Get
            Return _extractedData
        End Get
        Set(value As String)
            _extractedData = value
        End Set
    End Property

    ''' <summary>
    ''' ข้อความการตรวจสอบความถูกต้อง
    ''' </summary>
    Public Property ValidationMessages() As String
        Get
            Return _validationMessages
        End Get
        Set(value As String)
            _validationMessages = value
        End Set
    End Property

    ''' <summary>
    ''' รหัสผลิตภัณฑ์
    ''' </summary>
    Public Property ProductCode() As String
        Get
            Return _productCode
        End Get
        Set(value As String)
            _productCode = value
        End Set
    End Property

    ''' <summary>
    ''' รหัสอ้างอิง
    ''' </summary>
    Public Property ReferenceCode() As String
        Get
            Return _referenceCode
        End Get
        Set(value As String)
            _referenceCode = value
        End Set
    End Property

    ''' <summary>
    ''' จำนวน
    ''' </summary>
    Public Property Quantity() As Integer
        Get
            Return _quantity
        End Get
        Set(value As Integer)
            _quantity = value
        End Set
    End Property

    ''' <summary>
    ''' รหัสวันที่ผลิต
    ''' </summary>
    Public Property DateCode() As String
        Get
            Return _dateCode
        End Get
        Set(value As String)
            _dateCode = value
        End Set
    End Property

    ''' <summary>
    ''' สถานะความถูกต้องของข้อมูล
    ''' </summary>
    Public Property IsValid() As Boolean
        Get
            Return _isValid
        End Get
        Set(value As Boolean)
            _isValid = value
        End Set
    End Property

    ''' <summary>
    ''' ชื่อเครื่องคอมพิวเตอร์
    ''' </summary>
    Public Property ComputerName() As String
        Get
            Return _computerName
        End Get
        Set(value As String)
            _computerName = value
        End Set
    End Property

    ''' <summary>
    ''' ชื่อผู้ใช้
    ''' </summary>
    Public Property UserName() As String
        Get
            Return _userName
        End Get
        Set(value As String)
            _userName = value
        End Set
    End Property

    ''' <summary>
    ''' สถานะ Mission
    ''' </summary>
    Public Property MissionStatus() As String
        Get
            Return _missionStatus
        End Get
        Set(value As String)
            _missionStatus = value
        End Set
    End Property

    ''' <summary>
    ''' สร้าง constructor เริ่มต้น
    ''' </summary>
    Public Sub New()
        ' ตั้งค่าเริ่มต้น
        _scanDateTime = DateTime.Now
        _computerName = Environment.MachineName
        _userName = Environment.UserName
        _missionStatus = "ไม่มี"
    End Sub

    ''' <summary>
    ''' สร้าง constructor ที่รับข้อมูลพื้นฐาน
    ''' </summary>
    Public Sub New(productCode As String, originalData As String, isValid As Boolean)
        Me.New()
        _productCode = productCode
        _originalData = originalData
        _isValid = isValid
    End Sub

    ''' <summary>
    ''' สร้าง constructor ที่รับข้อมูลแบบเต็ม
    ''' </summary>
    Public Sub New(productCode As String, referenceCode As String, quantity As Integer, dateCode As String, originalData As String, isValid As Boolean)
        Me.New()
        _productCode = productCode
        _referenceCode = referenceCode
        _quantity = quantity
        _dateCode = dateCode
        _originalData = originalData
        _isValid = isValid
    End Sub

    ''' <summary>
    ''' แปลงเป็นข้อความสำหรับแสดงผล
    ''' </summary>
    Public Overrides Function ToString() As String
        Return $"รหัส: {_productCode}, อ้างอิง: {_referenceCode}, สถานะ: {If(_isValid, "ถูกต้อง", "ไม่ถูกต้อง")}"
    End Function
End Class 