''' <summary>
''' คลาสเก็บข้อมูลแต่ละแถวของ Excel ใน Memory
''' </summary>
Public Class ExcelRowData
    Private _rowNumber As Integer = 0
    Private _productCode As String = ""
    Private _column1Value As String = ""
    Private _column2Value As String = ""
    Private _column4Value As String = ""
    Private _column5Value As String = ""
    Private _column6Value As String = ""
    Private _loadedTime As DateTime = DateTime.Now
    
    ''' <summary>
    ''' หมายเลขแถวในไฟล์ Excel (เริ่มจาก 1)
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
    ''' รหัสผลิตภัณฑ์ (คอลัมน์ที่ 3)
    ''' </summary>
    Public Property ProductCode() As String
        Get
            Return _productCode
        End Get
        Set(value As String)
            _productCode = If(value?.Trim(), "")
        End Set
    End Property
    
    ''' <summary>
    ''' ข้อมูลคอลัมน์ที่ 1
    ''' </summary>
    Public Property Column1Value() As String
        Get
            Return _column1Value
        End Get
        Set(value As String)
            _column1Value = If(value?.Trim(), "")
        End Set
    End Property
    
    ''' <summary>
    ''' ข้อมูลคอลัมน์ที่ 2
    ''' </summary>
    Public Property Column2Value() As String
        Get
            Return _column2Value
        End Get
        Set(value As String)
            _column2Value = If(value?.Trim(), "")
        End Set
    End Property
    
    ''' <summary>
    ''' ข้อมูลคอลัมน์ที่ 4 (ผลลัพธ์หลัก)
    ''' </summary>
    Public Property Column4Value() As String
        Get
            Return _column4Value
        End Get
        Set(value As String)
            _column4Value = If(value?.Trim(), "")
        End Set
    End Property
    
    ''' <summary>
    ''' ข้อมูลคอลัมน์ที่ 5
    ''' </summary>
    Public Property Column5Value() As String
        Get
            Return _column5Value
        End Get
        Set(value As String)
            _column5Value = If(value?.Trim(), "")
        End Set
    End Property
    
    ''' <summary>
    ''' ข้อมูลคอลัมน์ที่ 6
    ''' </summary>
    Public Property Column6Value() As String
        Get
            Return _column6Value
        End Get
        Set(value As String)
            _column6Value = If(value?.Trim(), "")
        End Set
    End Property
    
    ''' <summary>
    ''' เวลาที่โหลดข้อมูลแถวนี้
    ''' </summary>
    Public ReadOnly Property LoadedTime() As DateTime
        Get
            Return _loadedTime
        End Get
    End Property
    
    ''' <summary>
    ''' ตรวจสอบว่าแถวนี้มีข้อมูลหรือไม่
    ''' </summary>
    Public ReadOnly Property HasData() As Boolean
        Get
            Return Not String.IsNullOrWhiteSpace(_productCode) OrElse
                   Not String.IsNullOrWhiteSpace(_column1Value) OrElse
                   Not String.IsNullOrWhiteSpace(_column2Value) OrElse
                   Not String.IsNullOrWhiteSpace(_column4Value) OrElse
                   Not String.IsNullOrWhiteSpace(_column5Value) OrElse
                   Not String.IsNullOrWhiteSpace(_column6Value)
        End Get
    End Property
    
    ''' <summary>
    ''' ตรวจสอบว่าแถวนี้มี Product Code หรือไม่
    ''' </summary>
    Public ReadOnly Property HasProductCode() As Boolean
        Get
            Return Not String.IsNullOrWhiteSpace(_productCode)
        End Get
    End Property
    
    ''' <summary>
    ''' ตรวจสอบว่าแถวนี้มีข้อมูลครบถ้วนหรือไม่
    ''' </summary>
    Public ReadOnly Property IsComplete() As Boolean
        Get
            Return HasProductCode AndAlso Not String.IsNullOrWhiteSpace(_column4Value)
        End Get
    End Property
    
    ''' <summary>
    ''' Constructor เริ่มต้น
    ''' </summary>
    Public Sub New()
        _loadedTime = DateTime.Now
    End Sub
    
    ''' <summary>
    ''' Constructor พร้อมกำหนดหมายเลขแถว
    ''' </summary>
    ''' <param name="rowNumber">หมายเลขแถว</param>
    Public Sub New(rowNumber As Integer)
        Me.New()
        _rowNumber = rowNumber
    End Sub
    
    ''' <summary>
    ''' Constructor พร้อมข้อมูลเบื้องต้น
    ''' </summary>
    ''' <param name="rowNumber">หมายเลขแถว</param>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    Public Sub New(rowNumber As Integer, productCode As String)
        Me.New(rowNumber)
        _productCode = If(productCode?.Trim(), "")
    End Sub
    
    ''' <summary>
    ''' แสดงข้อมูลแถวในรูปแบบข้อความ
    ''' </summary>
    ''' <returns>ข้อความแสดงข้อมูล</returns>
    Public Overrides Function ToString() As String
        If HasProductCode Then
            Return $"Row {_rowNumber}: {_productCode} -> {_column4Value}"
        Else
            Return $"Row {_rowNumber}: (ไม่มี Product Code)"
        End If
    End Function
    
    ''' <summary>
    ''' ได้ข้อมูลในรูปแบบ CSV
    ''' </summary>
    ''' <returns>ข้อมูลรูปแบบ CSV</returns>
    Public Function ToCSV() As String
        Return $"""{_rowNumber}"",""{_productCode}"",""{_column1Value}"",""{_column2Value}"",""{_column4Value}"",""{_column5Value}"",""{_column6Value}"""
    End Function
    
    ''' <summary>
    ''' ได้ข้อมูลในรูปแบบ Dictionary
    ''' </summary>
    ''' <returns>Dictionary ของข้อมูล</returns>
    Public Function ToDictionary() As Dictionary(Of String, String)
        Return New Dictionary(Of String, String) From {
            {"RowNumber", _rowNumber.ToString()},
            {"ProductCode", _productCode},
            {"Column1", _column1Value},
            {"Column2", _column2Value},
            {"Column4", _column4Value},
            {"Column5", _column5Value},
            {"Column6", _column6Value},
            {"LoadedTime", _loadedTime.ToString("yyyy-MM-dd HH:mm:ss")}
        }
    End Function
    
    ''' <summary>
    ''' ได้ข้อมูลในรูปแบบ JSON-like string
    ''' </summary>
    ''' <returns>ข้อมูลรูปแบบ JSON</returns>
    Public Function ToJsonString() As String
        Return "{" & vbNewLine &
               $"  ""rowNumber"": {_rowNumber}," & vbNewLine &
               $"  ""productCode"": ""{_productCode}""," & vbNewLine &
               $"  ""column1"": ""{_column1Value}""," & vbNewLine &
               $"  ""column2"": ""{_column2Value}""," & vbNewLine &
               $"  ""column4"": ""{_column4Value}""," & vbNewLine &
               $"  ""column5"": ""{_column5Value}""," & vbNewLine &
               $"  ""column6"": ""{_column6Value}""," & vbNewLine &
               $"  ""loadedTime"": ""{_loadedTime:yyyy-MM-dd HH:mm:ss}""" & vbNewLine &
               "}"
    End Function
    
    ''' <summary>
    ''' เปรียบเทียบข้อมูลกับแถวอื่น
    ''' </summary>
    ''' <param name="other">แถวอื่นที่จะเปรียบเทียบ</param>
    ''' <returns>True ถ้าข้อมูลเหมือนกัน</returns>
    Public Function IsEqual(other As ExcelRowData) As Boolean
        If other Is Nothing Then
            Return False
        End If
        
        Return _rowNumber = other.RowNumber AndAlso
               _productCode.Equals(other.ProductCode, StringComparison.OrdinalIgnoreCase) AndAlso
               _column1Value.Equals(other.Column1Value, StringComparison.OrdinalIgnoreCase) AndAlso
               _column2Value.Equals(other.Column2Value, StringComparison.OrdinalIgnoreCase) AndAlso
               _column4Value.Equals(other.Column4Value, StringComparison.OrdinalIgnoreCase) AndAlso
               _column5Value.Equals(other.Column5Value, StringComparison.OrdinalIgnoreCase) AndAlso
               _column6Value.Equals(other.Column6Value, StringComparison.OrdinalIgnoreCase)
    End Function
    
    ''' <summary>
    ''' คัดลอกข้อมูลจากแถวอื่น
    ''' </summary>
    ''' <param name="source">แถวต้นฉบับ</param>
    Public Sub CopyFrom(source As ExcelRowData)
        If source Is Nothing Then
            Return
        End If
        
        _rowNumber = source.RowNumber
        _productCode = source.ProductCode
        _column1Value = source.Column1Value
        _column2Value = source.Column2Value
        _column4Value = source.Column4Value
        _column5Value = source.Column5Value
        _column6Value = source.Column6Value
        _loadedTime = DateTime.Now
    End Sub

    ''' <summary>
    ''' สร้างสำเนาของแถวนี้
    ''' </summary>
    ''' <returns>สำเนาของข้อมูล</returns>
    'Public Function Clone() As ExcelRowData
    '    Dim clone As New ExcelRowData(_rowNumber, _productCode) With {
    '        .Column1Value = _column1Value,
    '        .Column2Value = _column2Value,
    '        .Column4Value = _column4Value,
    '        .Column5Value = _column5Value,
    '        .Column6Value = _column6Value
    '    }
    '    Return clone
    'End Function

    ''' <summary>
    ''' ตรวจสอบว่า Product Code ตรงกับที่ค้นหาหรือไม่
    ''' </summary>
    ''' <param name="searchCode">รหัสที่ต้องการค้นหา</param>
    ''' <returns>True ถ้าตรงกัน</returns>
    Public Function MatchesProductCode(searchCode As String) As Boolean
        If String.IsNullOrWhiteSpace(searchCode) OrElse String.IsNullOrWhiteSpace(_productCode) Then
            Return False
        End If
        
        ' ตรวจสอบแบบแม่นยำ
        If _productCode.Equals(searchCode, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If
        
        ' ตรวจสอบแบบไม่สนใจ space และ dash
        Dim cleanProductCode = _productCode.Replace(" ", "").Replace("-", "")
        Dim cleanSearchCode = searchCode.Replace(" ", "").Replace("-", "")
        
        Return cleanProductCode.Equals(cleanSearchCode, StringComparison.OrdinalIgnoreCase)
    End Function
End Class