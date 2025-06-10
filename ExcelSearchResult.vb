Imports System.Collections.Generic

''' <summary>
''' คลาสเก็บผลลัพธ์การค้นหาในไฟล์ Excel
''' </summary>
Public Class ExcelSearchResult
    ''' <summary>
    ''' รหัสผลิตภัณฑ์ที่ค้นหา
    ''' </summary>
    Public Property SearchedProductCode As String = ""

    ''' <summary>
    ''' เส้นทางไฟล์ Excel ที่ค้นหา
    ''' </summary>
    Public Property ExcelFilePath As String = ""

    ''' <summary>
    ''' สถานะการค้นหาสำเร็จหรือไม่
    ''' </summary>
    Public Property IsSuccess As Boolean = False

    ''' <summary>
    ''' จำนวนรายการที่พบ
    ''' </summary>
    Public Property MatchCount As Integer = 0

    ''' <summary>
    ''' ข้อความอธิบายผลการค้นหา
    ''' </summary>
    Public Property SummaryMessage As String = ""

    ''' <summary>
    ''' ข้อความแสดงข้อผิดพลาด (ถ้ามี)
    ''' </summary>
    Public Property ErrorMessage As String = ""

    ''' <summary>
    ''' รายการผลลัพธ์ที่พบทั้งหมด
    ''' </summary>
    Public Property Matches As List(Of ExcelMatchResult) = New List(Of ExcelMatchResult)()

    ''' <summary>
    ''' ผลลัพธ์แรกที่พบ
    ''' </summary>
    Public Property FirstMatch As ExcelMatchResult = Nothing

    ''' <summary>
    ''' ตรวจสอบว่ามีข้อผิดพลาดหรือไม่
    ''' </summary>
    Public ReadOnly Property HasError As Boolean
        Get
            Return Not String.IsNullOrEmpty(ErrorMessage)
        End Get
    End Property

    ''' <summary>
    ''' ตรวจสอบว่ามีผลลัพธ์หรือไม่
    ''' </summary>
    Public ReadOnly Property HasMatches As Boolean
        Get
            Return Matches IsNot Nothing AndAlso Matches.Count > 0
        End Get
    End Property
End Class

''' <summary>
''' คลาสเก็บข้อมูลผลลัพธ์ที่ตรงกับการค้นหา
''' </summary>
Public Class ExcelMatchResult
    ''' <summary>
    ''' หมายเลขแถวในไฟล์ Excel
    ''' </summary>
    Public Property RowNumber As Integer = 0

    ''' <summary>
    ''' รหัสผลิตภัณฑ์ที่พบ
    ''' </summary>
    Public Property ProductCode As String = ""

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 1
    ''' </summary>
    Public Property Column1Value As String = ""

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 2
    ''' </summary>
    Public Property Column2Value As String = ""

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 4 (ผลลัพธ์หลัก)
    ''' </summary>
    Public Property Column4Value As String = ""

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 5
    ''' </summary>
    Public Property Column5Value As String = ""

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 6
    ''' </summary>
    Public Property Column6Value As String = ""

    ''' <summary>
    ''' ระบุว่าเป็นการตรงกันแบบ exact match หรือไม่
    ''' </summary>
    Public Property IsExactMatch As Boolean = False
End Class 