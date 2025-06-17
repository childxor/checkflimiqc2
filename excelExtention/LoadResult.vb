''' <summary>
''' คลาสสำหรับเก็บผลลัพธ์การโหลดข้อมูล Excel
''' </summary>
Public Class LoadResult
    Private _isSuccess As Boolean = False
    Private _message As String = ""
    Private _errorMessage As String = ""
    Private _data As List(Of ExcelRowData) = New List(Of ExcelRowData)()
    Private _loadTime As TimeSpan = TimeSpan.Zero
    Private _processedRows As Integer = 0
    Private _validRows As Integer = 0
    Private _skippedRows As Integer = 0
    Private _startTime As DateTime = DateTime.Now
    Private _endTime As DateTime = DateTime.MinValue
    
    ''' <summary>
    ''' สถานะความสำเร็จของการโหลด
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
    ''' ข้อความแสดงผลลัพธ์
    ''' </summary>
    Public Property Message() As String
        Get
            Return _message
        End Get
        Set(value As String)
            _message = If(value, "")
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
            _errorMessage = If(value, "")
        End Set
    End Property
    
    ''' <summary>
    ''' ข้อมูลที่โหลดได้
    ''' </summary>
    Public Property Data() As List(Of ExcelRowData)
        Get
            Return _data
        End Get
        Set(value As List(Of ExcelRowData))
            _data = If(value, New List(Of ExcelRowData)())
        End Set
    End Property
    
    ''' <summary>
    ''' เวลาที่ใช้ในการโหลด
    ''' </summary>
    Public Property LoadTime() As TimeSpan
        Get
            Return _loadTime
        End Get
        Set(value As TimeSpan)
            _loadTime = value
        End Set
    End Property
    
    ''' <summary>
    ''' จำนวนแถวที่ประมวลผลทั้งหมด
    ''' </summary>
    Public Property ProcessedRows() As Integer
        Get
            Return _processedRows
        End Get
        Set(value As Integer)
            _processedRows = value
        End Set
    End Property
    
    ''' <summary>
    ''' จำนวนแถวที่มีข้อมูลถูกต้อง
    ''' </summary>
    Public Property ValidRows() As Integer
        Get
            Return _validRows
        End Get
        Set(value As Integer)
            _validRows = value
        End Set
    End Property
    
    ''' <summary>
    ''' จำนวนแถวที่ข้ามไป
    ''' </summary>
    Public Property SkippedRows() As Integer
        Get
            Return _skippedRows
        End Get
        Set(value As Integer)
            _skippedRows = value
        End Set
    End Property
    
    ''' <summary>
    ''' เวลาเริ่มต้นการโหลด
    ''' </summary>
    Public Property StartTime() As DateTime
        Get
            Return _startTime
        End Get
        Set(value As DateTime)
            _startTime = value
        End Set
    End Property
    
    ''' <summary>
    ''' เวลาสิ้นสุดการโหลด
    ''' </summary>
    Public Property EndTime() As DateTime
        Get
            Return _endTime
        End Get
        Set(value As DateTime)
            _endTime = value
            If _startTime <> DateTime.MinValue AndAlso _endTime <> DateTime.MinValue Then
                _loadTime = _endTime - _startTime
            End If
        End Set
    End Property
    
    ''' <summary>
    ''' จำนวนข้อมูลที่โหลดได้
    ''' </summary>
    Public ReadOnly Property DataCount() As Integer
        Get
            Return _data.Count
        End Get
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
    ''' ตรวจสอบว่ามีข้อมูลหรือไม่
    ''' </summary>
    Public ReadOnly Property HasData() As Boolean
        Get
            Return _data IsNot Nothing AndAlso _data.Count > 0
        End Get
    End Property
    
    ''' <summary>
    ''' อัตราความสำเร็จในการโหลด (%)
    ''' </summary>
    Public ReadOnly Property SuccessRate() As Double
        Get
            If _processedRows = 0 Then
                Return 0
            End If
            Return (_validRows / _processedRows) * 100
        End Get
    End Property
    
    ''' <summary>
    ''' Constructor เริ่มต้น
    ''' </summary>
    Public Sub New()
        _startTime = DateTime.Now
        _data = New List(Of ExcelRowData)()
    End Sub
    
    ''' <summary>
    ''' Constructor พร้อมข้อความ
    ''' </summary>
    ''' <param name="isSuccess">สถานะความสำเร็จ</param>
    ''' <param name="message">ข้อความ</param>
    Public Sub New(isSuccess As Boolean, message As String)
        Me.New()
        _isSuccess = isSuccess
        _message = If(message, "")
    End Sub
    
    ''' <summary>
    ''' Constructor สำหรับข้อผิดพลาด
    ''' </summary>
    ''' <param name="errorMessage">ข้อความข้อผิดพลาด</param>
    Public Sub New(errorMessage As String)
        Me.New()
        _isSuccess = False
        _errorMessage = If(errorMessage, "")
        _message = $"เกิดข้อผิดพลาด: {_errorMessage}"
    End Sub
    
    ''' <summary>
    ''' เริ่มนับเวลาการโหลด
    ''' </summary>
    Public Sub StartTiming()
        _startTime = DateTime.Now
    End Sub
    
    ''' <summary>
    ''' หยุดนับเวลาการโหลด
    ''' </summary>
    Public Sub StopTiming()
        _endTime = DateTime.Now
        If _startTime <> DateTime.MinValue Then
            _loadTime = _endTime - _startTime
        End If
    End Sub
    
    ''' <summary>
    ''' เพิ่มข้อมูลแถวใหม่
    ''' </summary>
    ''' <param name="rowData">ข้อมูลแถว</param>
    Public Sub AddRow(rowData As ExcelRowData)
        If rowData IsNot Nothing Then
            _data.Add(rowData)
            _validRows += 1
        End If
        _processedRows += 1
    End Sub
    
    ''' <summary>
    ''' เพิ่มจำนวนแถวที่ข้าม
    ''' </summary>
    Public Sub AddSkippedRow()
        _skippedRows += 1
        _processedRows += 1
    End Sub
    
    ''' <summary>
    ''' ตั้งค่าผลลัพธ์เป็นสำเร็จ
    ''' </summary>
    ''' <param name="message">ข้อความแสดงความสำเร็จ</param>
    Public Sub SetSuccess(message As String)
        _isSuccess = True
        _message = If(message, "")
        _errorMessage = ""
    End Sub
    
    ''' <summary>
    ''' ตั้งค่าผลลัพธ์เป็นข้อผิดพลาด
    ''' </summary>
    ''' <param name="errorMessage">ข้อความข้อผิดพลาด</param>
    Public Sub SetError(errorMessage As String)
        _isSuccess = False
        _errorMessage = If(errorMessage, "")
        _message = $"เกิดข้อผิดพลาด: {_errorMessage}"
    End Sub
    
    ''' <summary>
    ''' แสดงข้อมูลสรุปผลลัพธ์
    ''' </summary>
    ''' <returns>ข้อความสรุป</returns>
    Public Overrides Function ToString() As String
        Dim result As New System.Text.StringBuilder()
        
        result.AppendLine("=== ผลลัพธ์การโหลดข้อมูล Excel ===")
        result.AppendLine($"สถานะ: {If(_isSuccess, "✅ สำเร็จ", "❌ ไม่สำเร็จ")}")
        
        If Not String.IsNullOrEmpty(_message) Then
            result.AppendLine($"ข้อความ: {_message}")
        End If
        
        If HasError Then
            result.AppendLine($"ข้อผิดพลาด: {_errorMessage}")
        End If
        
        If _isSuccess Then
            result.AppendLine($"ข้อมูลที่โหลด: {DataCount} แถว")
            result.AppendLine($"แถวที่ประมวลผล: {_processedRows}")
            result.AppendLine($"แถวที่ถูกต้อง: {_validRows}")
            result.AppendLine($"แถวที่ข้าม: {_skippedRows}")
            result.AppendLine($"อัตราความสำเร็จ: {SuccessRate:F1}%")
            result.AppendLine($"เวลาที่ใช้: {_loadTime.TotalSeconds:F2} วินาที")
        End If
        
        Return result.ToString()
    End Function
    
    ''' <summary>
    ''' ได้ข้อมูลสถิติการโหลดในรูปแบบสั้น
    ''' </summary>
    ''' <returns>ข้อความสถิติแบบสั้น</returns>
    Public Function ToShortString() As String
        If _isSuccess Then
            Return $"โหลด {DataCount} แถว ใน {_loadTime.TotalSeconds:F2} วินาที ({SuccessRate:F1}% สำเร็จ)"
        Else
            Return $"ไม่สำเร็จ: {_errorMessage}"
        End If
    End Function
    
    ''' <summary>
    ''' ได้ข้อมูลสถิติการโหลดในรูปแบบ Dictionary for logging
    ''' </summary>
    ''' <returns>Dictionary ของสถิติ</returns>
    Public Function ToDictionary() As Dictionary(Of String, Object)
        Return New Dictionary(Of String, Object) From {
            {"IsSuccess", _isSuccess},
            {"Message", _message},
            {"ErrorMessage", _errorMessage},
            {"DataCount", DataCount},
            {"ProcessedRows", _processedRows},
            {"ValidRows", _validRows},
            {"SkippedRows", _skippedRows},
            {"SuccessRate", SuccessRate},
            {"LoadTimeSeconds", _loadTime.TotalSeconds},
            {"StartTime", _startTime.ToString("yyyy-MM-dd HH:mm:ss")},
            {"EndTime", _endTime.ToString("yyyy-MM-dd HH:mm:ss")}
        }
    End Function

    ''' <summary>
    ''' สร้าง LoadResult สำหรับความสำเร็จ
    ''' </summary>
    ''' <param name="data">ข้อมูลที่โหลดได้</param>
    ''' <param name="message">ข้อความ</param>
    ''' <returns>LoadResult ที่สำเร็จ</returns>


    ''' <summary>
    ''' สร้าง LoadResult สำหรับข้อผิดพลาด
    ''' </summary>
    ''' <param name="errorMessage">ข้อความข้อผิดพลาด</param>
    ''' <returns>LoadResult ที่ไม่สำเร็จ</returns>
    Public Shared Function Failure(errorMessage As String) As LoadResult
        Dim result As New LoadResult(errorMessage)
        result.StopTiming()
        Return result
    End Function

End Class