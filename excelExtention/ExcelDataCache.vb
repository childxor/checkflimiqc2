Imports System.IO
Imports System.Collections.Generic
Imports System.Threading.Tasks

''' <summary>
''' คลาสสำหรับเก็บข้อมูล Excel ใน Memory (Singleton Pattern)
''' </summary>
Public Class ExcelDataCache
    Private Shared _instance As ExcelDataCache
    Private _excelData As List(Of ExcelRowData)
    Private _isLoaded As Boolean = False
    Private _excelFilePath As String = ""
    Private _loadedTime As DateTime
    Private _isLoading As Boolean = False
    
    ''' <summary>
    ''' Singleton Pattern - ใช้ Instance เดียวทั้งโปรแกรม
    ''' </summary>
    Public Shared ReadOnly Property Instance() As ExcelDataCache
        Get
            If _instance Is Nothing Then
                _instance = New ExcelDataCache()
            End If
            Return _instance
        End Get
    End Property
    
    Private Sub New()
        _excelData = New List(Of ExcelRowData)()
    End Sub
    
    ''' <summary>
    ''' ข้อมูล Excel ทั้งหมดใน Memory
    ''' </summary>
    Public ReadOnly Property ExcelData() As List(Of ExcelRowData)
        Get
            Return _excelData
        End Get
    End Property
    
    ''' <summary>
    ''' สถานะการโหลดข้อมูลสำเร็จแล้วหรือไม่
    ''' </summary>
    Public ReadOnly Property IsLoaded() As Boolean
        Get
            Return _isLoaded
        End Get
    End Property
    
    ''' <summary>
    ''' สถานะกำลังโหลดข้อมูลอยู่หรือไม่
    ''' </summary>
    Public ReadOnly Property IsLoading() As Boolean
        Get
            Return _isLoading
        End Get
    End Property
    
    ''' <summary>
    ''' เส้นทางไฟล์ Excel ที่โหลดอยู่
    ''' </summary>
    Public ReadOnly Property ExcelFilePath() As String
        Get
            Return _excelFilePath
        End Get
    End Property
    
    ''' <summary>
    ''' เวลาที่โหลดข้อมูลล่าสุด
    ''' </summary>
    Public ReadOnly Property LoadedTime() As DateTime
        Get
            Return _loadedTime
        End Get
    End Property
    
    ''' <summary>
    ''' จำนวนแถวข้อมูลที่โหลดแล้ว
    ''' </summary>
    Public ReadOnly Property RowCount() As Integer
        Get
            Return _excelData.Count
        End Get
    End Property
    
    ''' <summary>
    ''' โหลดข้อมูล Excel ทั้งหมดเข้า Memory
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <returns>ผลลัพธ์การโหลด</returns>
    Public Function LoadExcelData(excelPath As String) As LoadResult
        Dim result As New LoadResult()
        
        Try
            ' ตรวจสอบว่ากำลังโหลดอยู่หรือไม่
            If _isLoading Then
                result.IsSuccess = False
                result.Message = "กำลังโหลดข้อมูลอยู่ กรุณารอสักครู่"
                Return result
            End If
            
            Console.WriteLine($"เริ่มโหลดข้อมูล Excel: {Path.GetFileName(excelPath)}")
            Dim startTime = DateTime.Now
            _isLoading = True
            
            ' เคลียร์ข้อมูลเก่า
            _excelData.Clear()
            _isLoaded = False
            
            ' โหลดข้อมูลใหม่
            Dim loadResult = ExcelUtility.LoadDataFromExcel(excelPath)
            
            If loadResult.IsSuccess Then
                _excelData = loadResult.Data
                _excelFilePath = excelPath
                _isLoaded = True
                _loadedTime = DateTime.Now
                
                Dim elapsedTime = DateTime.Now - startTime
                result.IsSuccess = True
                result.Message = $"โหลดข้อมูล {_excelData.Count} แถว สำเร็จ (ใช้เวลา {elapsedTime.TotalSeconds:F2} วินาที)"
                result.Data = _excelData
                
                Console.WriteLine(result.Message)
            Else
                result.IsSuccess = False
                result.Message = loadResult.ErrorMessage
                result.ErrorMessage = loadResult.ErrorMessage
            End If
            
        Catch ex As Exception
            result.IsSuccess = False
            result.Message = $"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}"
            result.ErrorMessage = ex.Message
            Console.WriteLine($"Error in LoadExcelData: {ex.Message}")
        Finally
            _isLoading = False
        End Try
        
        Return result
    End Function
    
    ''' <summary>
    ''' ค้นหาข้อมูลใน Memory (เร็วมาก)
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์ที่ต้องการค้นหา</param>
    ''' <returns>ผลลัพธ์การค้นหา</returns>
    Public Function SearchInMemory(productCode As String) As ExcelUtility.ExcelSearchResult
        Dim result As New ExcelUtility.ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = _excelFilePath
        
        If Not _isLoaded Then
            result.ErrorMessage = "ข้อมูล Excel ยังไม่ได้โหลด กรุณาโหลดข้อมูลก่อน"
            result.IsSuccess = False
            result.SummaryMessage = "❌ ข้อมูล Excel ยังไม่พร้อม"
            Return result
        End If
        
        If String.IsNullOrWhiteSpace(productCode) Then
            result.ErrorMessage = "กรุณาระบุรหัสผลิตภัณฑ์ที่ต้องการค้นหา"
            result.IsSuccess = False
            result.SummaryMessage = "❌ ไม่มีรหัสผลิตภัณฑ์ที่ต้องการค้นหา"
            Return result
        End If
        
        Try
            Dim matches As New List(Of ExcelUtility.ExcelMatchResult)()
            Dim searchStartTime = DateTime.Now
            
            ' ค้นหาใน Memory (เร็วมาก)
            For Each row In _excelData
                If IsProductCodeMatch(row.ProductCode, productCode) Then
                    Dim match As New ExcelUtility.ExcelMatchResult() With {
                        .RowNumber = row.RowNumber,
                        .ProductCode = row.ProductCode,
                        .Column1Value = row.Column1Value,
                        .Column2Value = row.Column2Value,
                        .Column4Value = row.Column4Value,
                        .Column5Value = row.Column5Value,
                        .Column6Value = row.Column6Value,
                        .IsExactMatch = row.ProductCode.Equals(productCode, StringComparison.OrdinalIgnoreCase)
                    }
                    
                    matches.Add(match)
                    
                    ' จำกัดผลลัพธ์ไม่เกิน 10 รายการ
                    If matches.Count >= 10 Then
                        Exit For
                    End If
                End If
            Next
            
            Dim searchTime = DateTime.Now - searchStartTime
            Console.WriteLine($"ค้นหาใน Memory เสร็จสิ้นใน {searchTime.TotalMilliseconds:F2} มิลลิวินาที")
            
            ' กำหนดผลลัพธ์
            result.Matches = matches
            result.MatchCount = matches.Count
            
            If matches.Count > 0 Then
                result.IsSuccess = True
                result.FirstMatch = matches(0)
                
                If matches.Count = 1 Then
                    result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{productCode}' ที่แถว {matches(0).RowNumber}" & vbNewLine &
                                          $"ข้อมูล: {matches(0).Column4Value}"
                Else
                    result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{productCode}' จำนวน {matches.Count} รายการ"
                End If
            Else
                result.IsSuccess = False
                result.MatchCount = 0
                result.SummaryMessage = $"❌ ไม่พบรหัสผลิตภัณฑ์ '{productCode}' ในข้อมูล"
            End If
            
        Catch ex As Exception
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            result.IsSuccess = False
            result.SummaryMessage = "❌ เกิดข้อผิดพลาดในการค้นหา"
            Console.WriteLine($"Error in SearchInMemory: {ex.Message}")
        End Try
        
        Return result
    End Function
    
    ''' <summary>
    ''' ตรวจสอบว่า Product Code ตรงกันหรือไม่
    ''' </summary>
    Private Function IsProductCodeMatch(cellText As String, searchCode As String) As Boolean
        If String.IsNullOrEmpty(cellText) OrElse String.IsNullOrEmpty(searchCode) Then
            Return False
        End If

        ' ตรวจสอบแบบแม่นยำ (case-insensitive)
        If cellText.Equals(searchCode, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        ' ตรวจสอบแบบไม่สนใจ space และ dash
        Dim cleanCellText As String = cellText.Replace(" ", "").Replace("-", "")
        Dim cleanSearchCode As String = searchCode.Replace(" ", "").Replace("-", "")

        Return cleanCellText.Equals(cleanSearchCode, StringComparison.OrdinalIgnoreCase)
    End Function
    
    ''' <summary>
    ''' รีเฟรชข้อมูล (โหลดใหม่จากไฟล์เดิม)
    ''' </summary>
    ''' <returns>ผลลัพธ์การรีเฟรช</returns>
    Public Function RefreshData() As LoadResult
        If String.IsNullOrEmpty(_excelFilePath) Then
            Return New LoadResult() With {
                .IsSuccess = False,
                .Message = "ไม่มีไฟล์ Excel ที่จะรีเฟรช กรุณาโหลดข้อมูลใหม่",
                .ErrorMessage = "No file path available"
            }
        End If
        
        Console.WriteLine("กำลังรีเฟรชข้อมูล Excel...")
        Return LoadExcelData(_excelFilePath)
    End Function
    
    ''' <summary>
    ''' เคลียร์ข้อมูลออกจาก Memory
    ''' </summary>
    Public Sub ClearData()
        Try
            _excelData.Clear()
            _isLoaded = False
            _excelFilePath = ""
            _loadedTime = DateTime.MinValue
            _isLoading = False
            
            ' บังคับ Garbage Collection
            GC.Collect()
            GC.WaitForPendingFinalizers()
            
            Console.WriteLine("เคลียร์ข้อมูล Excel จาก Memory แล้ว")
        Catch ex As Exception
            Console.WriteLine($"Error in ClearData: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' ได้ข้อมูลสถิติการใช้งาน Memory
    ''' </summary>
    ''' <returns>ข้อมูลสถิติ</returns>
    Public Function GetMemoryStats() As String
        Try
            Dim memoryUsed As Long = GC.GetTotalMemory(False)
            Dim memoryMB As Double = memoryUsed / 1024 / 1024
            
            Return $"ข้อมูลใน Memory: {_excelData.Count} แถว" & vbNewLine &
                   $"หน่วยความจำที่ใช้: {memoryMB:F2} MB" & vbNewLine &
                   $"สถานะ: {If(_isLoaded, "โหลดแล้ว", "ยังไม่โหลด")}" & vbNewLine &
                   $"โหลดเมื่อ: {If(_isLoaded, _loadedTime.ToString("dd/MM/yyyy HH:mm:ss"), "ยังไม่ได้โหลด")}"
        Catch ex As Exception
            Return $"ไม่สามารถดูสถิติได้: {ex.Message}"
        End Try
    End Function
    
    ''' <summary>
    ''' ตรวจสอบว่าต้องรีเฟรชข้อมูลหรือไม่ (ตามเวลา)
    ''' </summary>
    ''' <param name="maxAgeMinutes">อายุสูงสุดของข้อมูลในหน่วยนาที</param>
    ''' <returns>True ถ้าควรรีเฟรช</returns>
    Public Function ShouldRefresh(Optional maxAgeMinutes As Integer = 60) As Boolean
        If Not _isLoaded Then
            Return True
        End If
        
        Dim age = DateTime.Now - _loadedTime
        Return age.TotalMinutes > maxAgeMinutes
    End Function
End Class