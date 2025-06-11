Imports System.Collections.Generic

''' <summary>
''' คลาสเก็บผลลัพธ์การค้นหาในไฟล์ Excel
''' </summary>
Public Class ExcelSearchResult
    Private _searchedProductCode As String = ""
    Private _excelFilePath As String = ""
    Private _isSuccess As Boolean = False
    Private _matchCount As Integer = 0
    Private _summaryMessage As String = ""
    Private _errorMessage As String = ""
    Private _matches As List(Of ExcelMatchResult) = New List(Of ExcelMatchResult)()
    Private _firstMatch As ExcelMatchResult = Nothing

    ''' <summary>
    ''' รหัสผลิตภัณฑ์ที่ค้นหา
    ''' </summary>
    Public Property SearchedProductCode() As String
        Get
            Return _searchedProductCode
        End Get
        Set(value As String)
            _searchedProductCode = value
        End Set
    End Property

    ''' <summary>
    ''' เส้นทางไฟล์ Excel ที่ค้นหา
    ''' </summary>
    Public Property ExcelFilePath() As String
        Get
            Return _excelFilePath
        End Get
        Set(value As String)
            _excelFilePath = value
        End Set
    End Property

    ''' <summary>
    ''' สถานะการค้นหาสำเร็จหรือไม่
    ''' </summary>
    Public Property IsSuccess() As Boolean
        Get
            Return _isSuccess
        End Get
        Set(value As Boolean)
            _isSuccess = value
        End Set
    End Property

    ''' <summary>
    ''' จำนวนรายการที่พบ
    ''' </summary>
    Public Property MatchCount() As Integer
        Get
            Return _matchCount
        End Get
        Set(value As Integer)
            _matchCount = value
        End Set
    End Property

    ''' <summary>
    ''' ข้อความอธิบายผลการค้นหา
    ''' </summary>
    Public Property SummaryMessage() As String
        Get
            Return _summaryMessage
        End Get
        Set(value As String)
            _summaryMessage = value
        End Set
    End Property

    ''' <summary>
    ''' ข้อความแสดงข้อผิดพลาด (ถ้ามี)
    ''' </summary>
    Public Property ErrorMessage() As String
        Get
            Return _errorMessage
        End Get
        Set(value As String)
            _errorMessage = value
        End Set
    End Property

    ''' <summary>
    ''' รายการผลลัพธ์ที่พบทั้งหมด
    ''' </summary>
    Public Property Matches() As List(Of ExcelMatchResult)
        Get
            Return _matches
        End Get
        Set(value As List(Of ExcelMatchResult))
            _matches = value
        End Set
    End Property

    ''' <summary>
    ''' ผลลัพธ์แรกที่พบ
    ''' </summary>
    Public Property FirstMatch() As ExcelMatchResult
        Get
            Return _firstMatch
        End Get
        Set(value As ExcelMatchResult)
            _firstMatch = value
        End Set
    End Property

    ''' <summary>
    ''' ตรวจสอบว่ามีข้อผิดพลาดหรือไม่
    ''' </summary>
    Public ReadOnly Property HasError() As Boolean
        Get
            Return Not String.IsNullOrEmpty(_errorMessage)
        End Get
    End Property

    ''' <summary>
    ''' ตรวจสอบว่ามีผลลัพธ์หรือไม่
    ''' </summary>
    Public ReadOnly Property HasMatches() As Boolean
        Get
            Return _matches IsNot Nothing AndAlso _matches.Count > 0
        End Get
    End Property
End Class

''' <summary>
''' คลาสเก็บข้อมูลผลลัพธ์ที่ตรงกับการค้นหา
''' </summary>
Public Class ExcelMatchResult
    Private _rowNumber As Integer = 0
    Private _productCode As String = ""
    Private _column1Value As String = ""
    Private _column2Value As String = ""
    Private _column4Value As String = ""
    Private _column5Value As String = ""
    Private _column6Value As String = ""
    Private _isExactMatch As Boolean = False

    ''' <summary>
    ''' หมายเลขแถวในไฟล์ Excel
    ''' </summary>
    Public Property RowNumber() As Integer
        Get
            Return _rowNumber
        End Get
        Set(value As Integer)
            _rowNumber = value
        End Set
    End Property

    ''' <summary>
    ''' รหัสผลิตภัณฑ์ที่พบ
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
    ''' ค่าในคอลัมน์ที่ 1
    ''' </summary>
    Public Property Column1Value() As String
        Get
            Return _column1Value
        End Get
        Set(value As String)
            _column1Value = value
        End Set
    End Property

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 2
    ''' </summary>
    Public Property Column2Value() As String
        Get
            Return _column2Value
        End Get
        Set(value As String)
            _column2Value = value
        End Set
    End Property

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 4 (ผลลัพธ์หลัก)
    ''' </summary>
    Public Property Column4Value() As String
        Get
            Return _column4Value
        End Get
        Set(value As String)
            _column4Value = value
        End Set
    End Property

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 5
    ''' </summary>
    Public Property Column5Value() As String
        Get
            Return _column5Value
        End Get
        Set(value As String)
            _column5Value = value
        End Set
    End Property

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 6
    ''' </summary>
    Public Property Column6Value() As String
        Get
            Return _column6Value
        End Get
        Set(value As String)
            _column6Value = value
        End Set
    End Property

    ''' <summary>
    ''' ระบุว่าเป็นการตรงกันแบบ exact match หรือไม่
    ''' </summary>
    Public Property IsExactMatch() As Boolean
        Get
            Return _isExactMatch
        End Get
        Set(value As Boolean)
            _isExactMatch = value
        End Set
    End Property
End Class 