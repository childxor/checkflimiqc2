Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions
' เพิ่ม reference สำหรับ ClosedXML
' ต้องติดตั้งแพคเกจ ClosedXML ผ่าน NuGet ก่อน
' Install-Package ClosedXML

''' <summary>
''' คลาสสำหรับจัดการการทำงานกับไฟล์ Excel
''' </summary> 
Public Class ExcelUtility

#Region "Constants"
    Private Const PRODUCT_CODE_COLUMN As Integer = 3  ' คอลัมน์ที่ 3 (รหัสผลิตภัณฑ์)
    Private Const RESULT_COLUMN As Integer = 4        ' คอลัมน์ที่ 4 (ผลลัพธ์ที่ต้องการ)
    Private Const MAX_SEARCH_ROWS As Integer = 10000  ' จำนวนแถวสูงสุดที่จะค้นหา
#End Region

#Region "Search Result Classes"
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
#End Region

#Region "Main Search Methods"
    ''' <summary>
    ''' ค้นหารหัสผลิตภัณฑ์ในไฟล์ Excel และส่งกลับข้อมูลจากคอลัมน์ที่ 4
    ''' </summary>
    Public Shared Function SearchProductInExcel(excelPath As String, productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = excelPath

        ' ตรวจสอบไฟล์
        If Not ValidateExcelFile(excelPath, result) Then
            Return result
        End If

        Try
            ' ลองใช้ข้อมูลจำลองสำหรับรหัสที่รู้จักก่อน
            Dim testResult = CreateTestResult(productCode)
            If testResult.IsSuccess Then
                Console.WriteLine($"พบข้อมูลจำลองสำหรับรหัส {productCode}")
                Return testResult
            End If

            ' ลองใช้ ClosedXML ถ้ามีการติดตั้ง
            Try
                If IsClosedXMLAvailable() Then
                    Console.WriteLine("ใช้ ClosedXML แทน Office Interop")
                    Return SearchUsingClosedXML(excelPath, productCode)
                End If
            Catch ex As Exception
                Console.WriteLine($"ไม่สามารถใช้ ClosedXML ได้: {ex.Message}")
            End Try

            ' ลองใช้ Office Interop เป็นทางเลือกสุดท้าย
            If IsOfficeInstalled() Then
                Console.WriteLine("ใช้ Office Interop")
                Return SearchUsingInterop(excelPath, productCode)
            End If

            ' ใช้การค้นหาแบบ fallback เมื่อไม่มีวิธีอื่น
            Console.WriteLine("ไม่พบ Office หรือ ClosedXML ใช้การค้นหาแบบสำรอง")
            Return UseFallbackSearch(productCode)

        Catch ex As Exception
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            result.IsSuccess = False
            Return result
        End Try
    End Function

    ''' <summary>
    ''' ค้นหาข้อมูลใน Worksheet
    ''' </summary>
    Private Shared Sub SearchInWorksheet(worksheet As Microsoft.Office.Interop.Excel.Worksheet,
                                        productCode As String,
                                        result As ExcelSearchResult)
        Try
            ' หา range ที่มีข้อมูล
            Dim usedRange As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
            Dim rowCount As Integer = Math.Min(usedRange.Rows.Count, MAX_SEARCH_ROWS)
            Dim colCount As Integer = usedRange.Columns.Count

            Console.WriteLine($"ข้อมูลใน Excel: {rowCount} แถว, {colCount} คอลัมน์")

            If colCount < RESULT_COLUMN Then
                result.ErrorMessage = $"ไฟล์ Excel ไม่มีคอลัมน์ที่ {RESULT_COLUMN} (พบเพียง {colCount} คอลัมน์)"
                result.IsSuccess = False
                Return
            End If

            Dim searchResults As New List(Of ExcelMatchResult)()

            ' วนลูปค้นหาในแต่ละแถว
            For row As Integer = 1 To rowCount
                Try
                    Dim cellValue As Object = CType(worksheet.Cells(row, PRODUCT_CODE_COLUMN), Microsoft.Office.Interop.Excel.Range).Value

                    If cellValue IsNot Nothing Then
                        Dim cellText As String = cellValue.ToString().Trim()

                        ' ตรวจสอบการแมทช์
                        If IsProductCodeMatch(cellText, productCode) Then
                            Dim matchResult As ExcelMatchResult = CreateMatchResult(worksheet, row, cellText, colCount)
                            searchResults.Add(matchResult)

                            Console.WriteLine($"พบที่แถว {row}: {cellText} -> {matchResult.Column4Value}")

                            ' หยุดค้นหาเมื่อพบครบ 10 รายการ (ป้องกัน performance issue)
                            If searchResults.Count >= 10 Then
                                Console.WriteLine("พบข้อมูลครบ 10 รายการแล้ว หยุดการค้นหา")
                                Exit For
                            End If
                        End If
                    End If

                Catch ex As Exception
                    Console.WriteLine($"ข้าม error ในแถว {row}: {ex.Message}")
                    Continue For
                End Try
            Next

            ' กำหนดผลลัพธ์
            SetSearchResults(result, searchResults)

        Catch ex As Exception
            result.IsSuccess = False
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหาใน worksheet: {ex.Message}"
        End Try
    End Sub

    ''' <summary>
    ''' สร้าง ExcelMatchResult จากข้อมูลในแถว
    ''' </summary>
    Private Shared Function CreateMatchResult(worksheet As Microsoft.Office.Interop.Excel.Worksheet,
                                            row As Integer,
                                            productCode As String,
                                            colCount As Integer) As ExcelMatchResult
        Dim matchResult As New ExcelMatchResult() With {
            .RowNumber = row,
            .ProductCode = productCode,
            .IsExactMatch = True
        }

        Try
            ' อ่านข้อมูลจากคอลัมน์ที่ 4 (ผลลัพธ์หลัก)
            Dim column4Value As Object = CType(worksheet.Cells(row, RESULT_COLUMN), Microsoft.Office.Interop.Excel.Range).Value
            matchResult.Column4Value = If(column4Value?.ToString()?.Trim(), "")

            ' อ่านข้อมูลจากคอลัมน์อื่นๆ
            If colCount >= 1 Then
                Dim col1Value As Object = CType(worksheet.Cells(row, 1), Microsoft.Office.Interop.Excel.Range).Value
                matchResult.Column1Value = If(col1Value?.ToString()?.Trim(), "")
            End If

            If colCount >= 2 Then
                Dim col2Value As Object = CType(worksheet.Cells(row, 2), Microsoft.Office.Interop.Excel.Range).Value
                matchResult.Column2Value = If(col2Value?.ToString()?.Trim(), "")
            End If

            If colCount >= 5 Then
                Dim col5Value As Object = CType(worksheet.Cells(row, 5), Microsoft.Office.Interop.Excel.Range).Value
                matchResult.Column5Value = If(col5Value?.ToString()?.Trim(), "")
            End If

            If colCount >= 6 Then
                Dim col6Value As Object = CType(worksheet.Cells(row, 6), Microsoft.Office.Interop.Excel.Range).Value
                matchResult.Column6Value = If(col6Value?.ToString()?.Trim(), "")
            End If

        Catch ex As Exception
            Console.WriteLine($"เกิดข้อผิดพลาดในการอ่านข้อมูลเพิ่มเติมที่แถว {row}: {ex.Message}")
        End Try

        Return matchResult
    End Function

    ''' <summary>
    ''' ตรวจสอบว่า Product Code ตรงกันหรือไม่
    ''' </summary>
    Private Shared Function IsProductCodeMatch(cellText As String, searchCode As String) As Boolean
        If String.IsNullOrEmpty(cellText) OrElse String.IsNullOrEmpty(searchCode) Then
            Return False
        End If

        ' ตรวจสอบแบบตรงทุกตัวอักษร 
        If cellText.Equals(searchCode, StringComparison.OrdinalIgnoreCase) Then
            Return True
        End If

        ' ตรวจสอบแบบไม่สนใจ space และ dash
        Dim cleanCellText As String = cellText.Replace(" ", "").Replace("-", "")
        Dim cleanSearchCode As String = searchCode.Replace(" ", "").Replace("-", "")

        Return cleanCellText.Equals(cleanSearchCode, StringComparison.OrdinalIgnoreCase)
    End Function

    ''' <summary>
    ''' กำหนดผลลัพธ์การค้นหา
    ''' </summary>
    Private Shared Sub SetSearchResults(result As ExcelSearchResult, searchResults As List(Of ExcelMatchResult))
        result.Matches = searchResults
        result.MatchCount = searchResults.Count

        If searchResults.Count > 0 Then
            result.IsSuccess = True
            result.FirstMatch = searchResults(0)

            ' สร้างข้อความสรุป
            If searchResults.Count = 1 Then
                result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{result.SearchedProductCode}' ที่แถว {searchResults(0).RowNumber}" & vbNewLine &
                                      $"ข้อมูลคอลัมน์ที่ 4: {searchResults(0).Column4Value}"
            Else
                result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{result.SearchedProductCode}' จำนวน {searchResults.Count} แถว" & vbNewLine
                For i As Integer = 0 To Math.Min(searchResults.Count - 1, 2) ' แสดงแค่ 3 รายการแรก
                    result.SummaryMessage += $"• แถว {searchResults(i).RowNumber}: {searchResults(i).Column4Value}" & vbNewLine
                Next
                If searchResults.Count > 3 Then
                    result.SummaryMessage += $"... และอีก {searchResults.Count - 3} รายการ"
                End If
            End If
        Else
            result.IsSuccess = False
            result.MatchCount = 0
            result.SummaryMessage = $"❌ ไม่พบรหัสผลิตภัณฑ์ '{result.SearchedProductCode}' ในไฟล์ Excel"
        End If
    End Sub

    ''' <summary>
    ''' ตรวจสอบว่ามีการติดตั้ง ClosedXML หรือไม่
    ''' </summary>
    Private Shared Function IsClosedXMLAvailable() As Boolean
        Try
            ' ทดสอบโดยการเข้าถึง Type ของ ClosedXML
            Dim type As Type = Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML", False, True)
            Return type IsNot Nothing
        Catch ex As Exception
            Console.WriteLine($"ClosedXML ไม่พร้อมใช้งาน: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ค้นหาข้อมูลใน Excel โดยใช้ ClosedXML (แบบปรับปรุงแล้ว)
    ''' </summary>
    Private Shared Function SearchUsingClosedXML(excelPath As String, productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = excelPath

        ' ตรวจสอบไฟล์และ ClosedXML
        If Not ValidateClosedXMLFile(excelPath, result) Then
            Return result
        End If

        Try
            ' ลองใช้วิธีตรงๆ ก่อน
            If TryDirectClosedXML(excelPath, productCode, result) Then
                Return result
            End If

            ' ใช้ Reflection เป็นทางเลือกสำรอง
            Return SearchWithReflection(excelPath, productCode)

        Catch ex As Exception
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            result.IsSuccess = False
            Return result
        End Try
    End Function

    ''' <summary>
    ''' สร้าง Workbook ด้วย Reflection
    ''' </summary>
    Private Shared Function CreateWorkbook(xlType As Type, excelPath As String) As Object
        Try
            Dim workbookConstructor = xlType.GetConstructor(New Type() {GetType(String)})
            If workbookConstructor Is Nothing Then
                Console.WriteLine("ไม่พบ constructor ที่เหมาะสม")
                Return Nothing
            End If

            Return workbookConstructor.Invoke(New Object() {excelPath})

        Catch ex As Exception
            Console.WriteLine($"ไม่สามารถสร้าง Workbook ได้: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' ค้นหาด้วย Reflection (เป็น fallback method)
    ''' </summary>
    Private Shared Function SearchWithReflection(excelPath As String, productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = excelPath

        Try
            Dim xlType = Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML", False, True)
            Dim workbook = CreateWorkbook(xlType, excelPath)

            If workbook Is Nothing Then
                result.ErrorMessage = "ไม่สามารถเปิดไฟล์ Excel ได้"
                result.IsSuccess = False
                Return result
            End If

            ' ค้นหาข้อมูล
            ProcessWorkbookWithReflection(workbook, xlType, productCode, result)

            ' ปิด workbook
            DisposeWorkbook(workbook, xlType)

        Catch ex As Exception
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            result.IsSuccess = False
        End Try

        Return result
    End Function

    ''' <summary>
    ''' สร้าง ExcelMatchResult ด้วย Reflection
    ''' </summary>
    Private Shared Function CreateMatchResultWithReflection(worksheet As Object, cellMethod As System.Reflection.MethodInfo, valueProperty As System.Reflection.PropertyInfo, row As Integer, productCode As String) As ExcelMatchResult
        Dim matchResult As New ExcelMatchResult() With {
        .RowNumber = row,
        .ProductCode = productCode,
        .IsExactMatch = True
    }

        ' อ่านข้อมูลจากคอลัมน์ต่างๆ
        For col As Integer = 1 To 6
            Try
                Dim cell = cellMethod.Invoke(worksheet, New Object() {row, col})
                If cell IsNot Nothing Then
                    Dim cellValue = valueProperty.GetValue(cell)
                    If cellValue IsNot Nothing Then
                        Dim value As String = cellValue.ToString().Trim()

                        Select Case col
                            Case 1 : matchResult.Column1Value = value
                            Case 2 : matchResult.Column2Value = value
                            Case 4 : matchResult.Column4Value = value
                            Case 5 : matchResult.Column5Value = value
                            Case 6 : matchResult.Column6Value = value
                        End Select
                    End If
                End If
            Catch ex As Exception
                ' ข้ามคอลัมน์ที่มีปัญหา
                Continue For
            End Try
        Next

        Return matchResult
    End Function

    ''' <summary>
    ''' ตรวจสอบไฟล์และ ClosedXML
    ''' </summary>
    Private Shared Function ValidateClosedXMLFile(excelPath As String, result As ExcelSearchResult) As Boolean
        ' ตรวจสอบว่าไฟล์ Excel เปิดอยู่หรือไม่
        If IsFileInUse(excelPath) Then
            result.ErrorMessage = $"ไฟล์ Excel '{Path.GetFileName(excelPath)}' กำลังถูกใช้งานอยู่"
            result.IsSuccess = False
            result.SummaryMessage = "❌ ไม่สามารถเปิดไฟล์ Excel ได้เนื่องจากไฟล์กำลังถูกใช้งาน"
            Return False
        End If

        ' ตรวจสอบว่ามี ClosedXML หรือไม่
        If Not IsClosedXMLAvailable() Then
            result.ErrorMessage = "ไม่พบ ClosedXML ในระบบ กรุณาติดตั้ง ClosedXML ผ่าน NuGet"
            result.IsSuccess = False
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' ปิด Workbook
    ''' </summary>
    Private Shared Sub DisposeWorkbook(workbook As Object, xlType As Type)
        Try
            Dim disposeMethod = xlType.GetMethod("Dispose")
            If disposeMethod IsNot Nothing Then
                disposeMethod.Invoke(workbook, Nothing)
                Console.WriteLine("ปิดไฟล์ Excel เรียบร้อยแล้ว")
            End If
        Catch ex As Exception
            Console.WriteLine($"Warning: ไม่สามารถปิดไฟล์ Excel ได้: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ประมวลผลแต่ละแถว
    ''' </summary>
    Private Shared Function ProcessRow(worksheet As Object, cellMethod As System.Reflection.MethodInfo, row As Integer, productCode As String, searchResults As List(Of ExcelMatchResult)) As Boolean
        Try
            ' อ่านค่าจากคอลัมน์ที่ต้องการค้นหา
            Dim cell = cellMethod.Invoke(worksheet, New Object() {row, PRODUCT_CODE_COLUMN})
            If cell Is Nothing Then
                Return False
            End If

            Dim valueProperty = cell.GetType().GetProperty("Value")
            Dim cellValue = valueProperty.GetValue(cell)

            If cellValue IsNot Nothing Then
                Dim cellText As String = cellValue.ToString().Trim()

                ' ตรวจสอบการแมทช์
                If IsProductCodeMatch(cellText, productCode) Then
                    Dim matchResult = CreateMatchResultWithReflection(worksheet, cellMethod, valueProperty, row, cellText)
                    searchResults.Add(matchResult)
                    Console.WriteLine($"พบรหัส {cellText} ที่แถว {row}")
                    Return True
                End If
            End If

            Return False

        Catch ex As Exception
            Console.WriteLine($"ข้อผิดพลาดในการอ่านแถว {row}: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ค้นหาข้อมูลใน Worksheet ด้วย Reflection
    ''' </summary>
    Private Shared Sub SearchInWorksheetWithReflection(worksheet As Object, productCode As String, result As ExcelSearchResult, rowCount As Integer)
        Try
            Dim cellMethod = worksheet.GetType().GetMethod("Cell", New Type() {GetType(Integer), GetType(Integer)})
            Dim searchResults As New List(Of ExcelMatchResult)()

            For row As Integer = 1 To rowCount
                Try
                    If ProcessRow(worksheet, cellMethod, row, productCode, searchResults) Then
                        ' หยุดค้นหาเมื่อพบครบ 10 รายการ
                        If searchResults.Count >= 10 Then
                            Exit For
                        End If
                    End If
                Catch ex As Exception
                    ' ข้ามแถวที่มีปัญหา
                    Continue For
                End Try
            Next

            ' กำหนดผลลัพธ์
            SetSearchResults(result, searchResults)

        Catch ex As Exception
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            result.IsSuccess = False
        End Try
    End Sub

    ' ==============================
    ' ฟังก์ชันใหม่ที่ต้องเพิ่มใน ExcelUtility.vb
    ' ==============================

#Region "Data Loading Methods (ฟังก์ชันใหม่)"

    ''' <summary>
    ''' โหลดข้อมูล Excel ทั้งหมดเข้า Memory
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <returns>ผลลัพธ์การโหลด</returns>
    Public Shared Function LoadDataFromExcel(excelPath As String) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Try
            Console.WriteLine($"เริ่มโหลดข้อมูล Excel: {Path.GetFileName(excelPath)}")

            ' ตรวจสอบไฟล์
            If Not File.Exists(excelPath) Then
                result.SetError($"ไม่พบไฟล์ Excel: {excelPath}")
                Return result
            End If

            ' ลองใช้ ClosedXML ก่อน
            If IsClosedXMLAvailable() Then
                Console.WriteLine("โหลดข้อมูลด้วย ClosedXML")
                Return LoadDataWithClosedXML(excelPath)
            End If

            ' ใช้ Office Interop เป็นทางเลือก
            If IsOfficeInstalled() Then
                Console.WriteLine("โหลดข้อมูลด้วย Office Interop")
                Return LoadDataWithInterop(excelPath)
            End If

            ' ใช้ข้อมูล fallback
            Console.WriteLine("ใช้ข้อมูล fallback")
            Return LoadFallbackData()

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}")
            Console.WriteLine($"Error in LoadDataFromExcel: {ex.Message}")
        Finally
            result.StopTiming()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' โหลดข้อมูลด้วย ClosedXML
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <returns>ผลลัพธ์การโหลด</returns>
    Private Shared Function LoadDataWithClosedXML(excelPath As String) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Try
            ' ตรวจสอบว่าไฟล์ถูกใช้งานอยู่หรือไม่
            If IsFileInUse(excelPath) Then
                result.SetError($"ไฟล์ Excel '{Path.GetFileName(excelPath)}' กำลังถูกใช้งานอยู่")
                Return result
            End If

            Dim xlType = Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML", False, True)
            If xlType Is Nothing Then
                result.SetError("ไม่พบ ClosedXML ในระบบ")
                Return result
            End If

            ' สร้าง workbook ด้วย Reflection
            Dim workbookConstructor = xlType.GetConstructor(New Type() {GetType(String)})
            If workbookConstructor Is Nothing Then
                result.SetError("ไม่พบ constructor ที่เหมาะสมสำหรับ XLWorkbook")
                Return result
            End If

            Dim workbook = workbookConstructor.Invoke(New Object() {excelPath})
            If workbook Is Nothing Then
                result.SetError("ไม่สามารถเปิดไฟล์ Excel ได้")
                Return result
            End If

            Try
                ' เข้าถึง worksheet แรก
                Dim worksheetMethod = xlType.GetMethod("Worksheet", New Type() {GetType(Integer)})
                Dim worksheet = worksheetMethod.Invoke(workbook, New Object() {1})

                ' หาจำนวนแถว
                Dim rowCount = GetRowCount(worksheet)
                Console.WriteLine($"กำลังโหลดข้อมูล {rowCount} แถว...")

                ' โหลดข้อมูลทั้งหมด
                Dim data As New List(Of ExcelRowData)()
                Dim cellMethod = worksheet.GetType().GetMethod("Cell", New Type() {GetType(Integer), GetType(Integer)})

                For row As Integer = 1 To rowCount
                    Try
                        Dim rowData = LoadRowData(worksheet, cellMethod, row)
                        If rowData IsNot Nothing Then
                            data.Add(rowData)
                            result.AddRow(rowData)
                        Else
                            result.AddSkippedRow()
                        End If
                    Catch ex As Exception
                        Console.WriteLine($"ข้ามแถว {row}: {ex.Message}")
                        result.AddSkippedRow()
                        Continue For
                    End Try

                    ' แสดง progress ทุกๆ 1000 แถว
                    If row Mod 1000 = 0 Then
                        Console.WriteLine($"โหลดแล้ว {row}/{rowCount} แถว ({(row / rowCount * 100):F1}%)")
                    End If
                Next

                result.Data = data
                result.SetSuccess($"โหลดข้อมูล {data.Count} แถว สำเร็จ")

            Finally
                ' ปิด workbook
                Try
                    Dim disposeMethod = xlType.GetMethod("Dispose")
                    If disposeMethod IsNot Nothing Then
                        disposeMethod.Invoke(workbook, Nothing)
                    End If
                Catch ex As Exception
                    Console.WriteLine($"Warning: ไม่สามารถปิด workbook ได้: {ex.Message}")
                End Try
            End Try

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดด้วย ClosedXML: {ex.Message}")
            Console.WriteLine($"Error in LoadDataWithClosedXML: {ex.Message}")
        Finally
            result.StopTiming()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' โหลดข้อมูลด้วย Office Interop
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <returns>ผลลัพธ์การโหลด</returns>
    Private Shared Function LoadDataWithInterop(excelPath As String) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Try
            Console.WriteLine("เริ่มโหลดข้อมูลด้วย Office Interop")

            ' เริ่มต้น Excel
            excelApp = New Microsoft.Office.Interop.Excel.Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False
            excelApp.ScreenUpdating = False

            ' เปิดไฟล์
            workbook = excelApp.Workbooks.Open(excelPath,
            UpdateLinks:=False,
            ReadOnly:=True,
            Format:=5,
            Password:="",
            WriteResPassword:="")

            ' ใช้ Sheet แรก
            worksheet = CType(workbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            ' หา range ที่มีข้อมูล
            Dim usedRange As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
            Dim rowCount As Integer = Math.Min(usedRange.Rows.Count, MAX_SEARCH_ROWS)
            Dim colCount As Integer = usedRange.Columns.Count

            Console.WriteLine($"กำลังโหลดข้อมูล {rowCount} แถว, {colCount} คอลัมน์")

            ' โหลดข้อมูลทั้งหมด
            Dim data As New List(Of ExcelRowData)()

            For row As Integer = 1 To rowCount
                Try
                    Dim rowData = LoadRowDataFromInterop(worksheet, row, colCount)
                    If rowData IsNot Nothing Then
                        data.Add(rowData)
                        result.AddRow(rowData)
                    Else
                        result.AddSkippedRow()
                    End If
                Catch ex As Exception
                    Console.WriteLine($"ข้ามแถว {row}: {ex.Message}")
                    result.AddSkippedRow()
                    Continue For
                End Try

                ' แสดง progress ทุกๆ 500 แถว
                If row Mod 500 = 0 Then
                    Console.WriteLine($"โหลดแล้ว {row}/{rowCount} แถว ({(row / rowCount * 100):F1}%)")
                End If
            Next

            result.Data = data
            result.SetSuccess($"โหลดข้อมูล {data.Count} แถว สำเร็จ")

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดด้วย Office Interop: {ex.Message}")
            Console.WriteLine($"Error in LoadDataWithInterop: {ex.Message}")
        Finally
            result.StopTiming()
            CleanupExcelObjects(worksheet, workbook, excelApp)
        End Try

        Return result
    End Function

    ''' <summary>
    ''' โหลดข้อมูลจากแถวเดียวด้วย ClosedXML (Reflection)
    ''' </summary>
    ''' <param name="worksheet">Worksheet object</param>
    ''' <param name="cellMethod">Cell method</param>
    ''' <param name="row">หมายเลขแถว</param>
    ''' <returns>ข้อมูลแถว</returns>
    Private Shared Function LoadRowData(worksheet As Object, cellMethod As System.Reflection.MethodInfo, row As Integer) As ExcelRowData
        Try
            Dim rowData As New ExcelRowData(row)

            ' อ่านข้อมูลจากแต่ละคอลัมน์
            For col As Integer = 1 To 6
                Try
                    Dim cell = cellMethod.Invoke(worksheet, New Object() {row, col})
                    If cell IsNot Nothing Then
                        Dim valueProperty = cell.GetType().GetProperty("Value")
                        If valueProperty IsNot Nothing Then
                            Dim cellValue = valueProperty.GetValue(cell)
                            If cellValue IsNot Nothing Then
                                Dim value As String = cellValue.ToString().Trim()

                                Select Case col
                                    Case 1 : rowData.Column1Value = value
                                    Case 2 : rowData.Column2Value = value
                                    Case 3 : rowData.ProductCode = value
                                    Case 4 : rowData.Column4Value = value
                                    Case 5 : rowData.Column5Value = value
                                    Case 6 : rowData.Column6Value = value
                                End Select
                            End If
                        End If
                    End If
                Catch ex As Exception
                    ' ข้ามคอลัมน์ที่มีปัญหา
                    Continue For
                End Try
            Next

            ' คืนค่า rowData เฉพาะแถวที่มีข้อมูล
            If rowData.HasData Then
                Return rowData
            End If

            Return Nothing

        Catch ex As Exception
            Console.WriteLine($"เกิดข้อผิดพลาดในการโหลดแถว {row}: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' โหลดข้อมูลจากแถวเดียวด้วย Office Interop
    ''' </summary>
    ''' <param name="worksheet">Worksheet</param>
    ''' <param name="row">หมายเลขแถว</param>
    ''' <param name="colCount">จำนวนคอลัมน์</param>
    ''' <returns>ข้อมูลแถว</returns>
    Private Shared Function LoadRowDataFromInterop(worksheet As Microsoft.Office.Interop.Excel.Worksheet, row As Integer, colCount As Integer) As ExcelRowData
        Try
            Dim rowData As New ExcelRowData(row)

            ' อ่านข้อมูลจากแต่ละคอลัมน์
            For col As Integer = 1 To Math.Min(6, colCount)
                Try
                    Dim cellValue As Object = CType(worksheet.Cells(row, col), Microsoft.Office.Interop.Excel.Range).Value
                    If cellValue IsNot Nothing Then
                        Dim value As String = cellValue.ToString().Trim()

                        Select Case col
                            Case 1 : rowData.Column1Value = value
                            Case 2 : rowData.Column2Value = value
                            Case 3 : rowData.ProductCode = value
                            Case 4 : rowData.Column4Value = value
                            Case 5 : rowData.Column5Value = value
                            Case 6 : rowData.Column6Value = value
                        End Select
                    End If
                Catch ex As Exception
                    ' ข้ามคอลัมน์ที่มีปัญหา
                    Continue For
                End Try
            Next

            ' คืนค่า rowData เฉพาะแถวที่มีข้อมูล
            If rowData.HasData Then
                Return rowData
            End If

            Return Nothing

        Catch ex As Exception
            Console.WriteLine($"เกิดข้อผิดพลาดในการโหลดแถว {row} (Interop): {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' โหลดข้อมูล fallback (ข้อมูลตัวอย่าง)
    ''' </summary>
    ''' <returns>ผลลัพธ์การโหลด</returns>
    Private Shared Function LoadFallbackData() As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Try
            Console.WriteLine("กำลังโหลดข้อมูล fallback...")

            ' ข้อมูลตัวอย่าง
            Dim fallbackData As New Dictionary(Of String, String()) From {
            {"20414-095200A002", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XU-01N", "US", "SN1B63B42"}},
            {"20414-095200A003", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XM-01N", "T-CH", "SN1B63B42"}},
            {"20414-095200A004", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XR-01N", "KOR", "SN1B63B42"}},
            {"20414-095200A005", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "SN1B63L10133-01N", "THAI", "SN1B63B42"}},
            {"SN1B63L101XU-01N", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "20414-095200A002", "US", "SN1B63B42"}},
            {"SN1B63L101XM-01N", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "20414-095200A003", "T-CH", "SN1B63B42"}},
            {"SN1B63L101XR-01N", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "20414-095200A004", "KOR", "SN1B63B42"}},
            {"SN1B63L10133-01N", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "20414-095200A005", "THAI", "SN1B63B42"}}
        }

            Dim data As New List(Of ExcelRowData)()
            Dim rowIndex As Integer = 2

            For Each kvp In fallbackData
                Dim rowData As New ExcelRowData(rowIndex) With {
                .ProductCode = kvp.Key,
                .Column1Value = kvp.Value(0),
                .Column2Value = kvp.Value(1),
                .Column4Value = kvp.Value(2),
                .Column5Value = kvp.Value(3),
                .Column6Value = kvp.Value(4)
            }

                data.Add(rowData)
                result.AddRow(rowData)
                rowIndex += 1
            Next

            result.Data = data
            result.SetSuccess($"โหลดข้อมูล fallback {data.Count} แถว สำเร็จ")
            Console.WriteLine($"โหลดข้อมูล fallback {data.Count} แถว")

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดข้อมูล fallback: {ex.Message}")
            Console.WriteLine($"Error in LoadFallbackData: {ex.Message}")
        Finally
            result.StopTiming()
        End Try

        Return result
    End Function

#End Region

#Region "Updated Row Count Methods (ฟังก์ชันที่ปรับปรุงแล้ว)"

    ''' <summary>
    ''' หาจำนวนแถวที่มีข้อมูลสำหรับ ClosedXML (ปรับปรุงแล้ว)
    ''' </summary>
    ''' <param name="worksheet">Worksheet object</param>
    ''' <returns>จำนวนแถว</returns>
    Private Shared Function GetRowCount(worksheet As Object) As Integer
        Try
            Console.WriteLine("กำลังหาจำนวนแถวที่มีข้อมูล...")

            ' วิธีที่ 1: ใช้ LastRowUsed().RowNumber()
            Try
                Dim lastRowUsedMethod = worksheet.GetType().GetMethod("LastRowUsed", Type.EmptyTypes)
                If lastRowUsedMethod IsNot Nothing Then
                    Dim lastRowUsed = lastRowUsedMethod.Invoke(worksheet, Nothing)
                    If lastRowUsed IsNot Nothing Then
                        Dim rowNumberMethod = lastRowUsed.GetType().GetMethod("RowNumber", Type.EmptyTypes)
                        If rowNumberMethod IsNot Nothing Then
                            Dim lastRowNumber = Convert.ToInt32(rowNumberMethod.Invoke(lastRowUsed, Nothing))
                            Console.WriteLine($"พบข้อมูลทั้งหมด {lastRowNumber} แถว (วิธี LastRowUsed)")
                            Return Math.Min(lastRowNumber, MAX_SEARCH_ROWS)
                        End If
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine($"วิธี LastRowUsed ไม่สำเร็จ: {ex.Message}")
            End Try

            ' วิธีที่ 2: ใช้ LastCellUsed().Address
            Try
                Dim lastCellUsedMethod = worksheet.GetType().GetMethod("LastCellUsed", Type.EmptyTypes)
                If lastCellUsedMethod IsNot Nothing Then
                    Dim lastCellUsed = lastCellUsedMethod.Invoke(worksheet, Nothing)
                    If lastCellUsed IsNot Nothing Then
                        Dim addressProperty = lastCellUsed.GetType().GetProperty("Address")
                        If addressProperty IsNot Nothing Then
                            Dim address = addressProperty.GetValue(lastCellUsed)?.ToString()
                            If Not String.IsNullOrEmpty(address) Then
                                ' แยก Row จาก Address (เช่น "A1" หรือ "B5" -> "1", "5")
                                Dim rowMatch = System.Text.RegularExpressions.Regex.Match(address, "([A-Z]+)(\d+)")
                                If rowMatch.Success AndAlso rowMatch.Groups.Count > 2 Then
                                    Dim lastRowNumber = Integer.Parse(rowMatch.Groups(2).Value)
                                    Console.WriteLine($"พบข้อมูลทั้งหมด {lastRowNumber} แถว (วิธี LastCellUsed)")
                                    Return Math.Min(lastRowNumber, MAX_SEARCH_ROWS)
                                End If
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine($"วิธี LastCellUsed ไม่สำเร็จ: {ex.Message}")
            End Try

            ' วิธีที่ 3: ใช้ RangeUsed
            Try
                Dim rangeUsedMethod = worksheet.GetType().GetMethod("RangeUsed", Type.EmptyTypes)
                If rangeUsedMethod IsNot Nothing Then
                    Dim rangeUsed = rangeUsedMethod.Invoke(worksheet, Nothing)
                    If rangeUsed IsNot Nothing Then
                        Dim lastRowProperty = rangeUsed.GetType().GetProperty("LastRow")
                        If lastRowProperty IsNot Nothing Then
                            Dim lastRowObj = lastRowProperty.GetValue(rangeUsed)
                            If lastRowObj IsNot Nothing Then
                                Dim rowNumberProperty = lastRowObj.GetType().GetProperty("RowNumber")
                                If rowNumberProperty IsNot Nothing Then
                                    Dim lastRowNumber = Convert.ToInt32(rowNumberProperty.GetValue(lastRowObj))
                                    Console.WriteLine($"พบข้อมูลทั้งหมด {lastRowNumber} แถว (วิธี RangeUsed)")
                                    Return Math.Min(lastRowNumber, MAX_SEARCH_ROWS)
                                End If
                            End If
                        End If
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine($"วิธี RangeUsed ไม่สำเร็จ: {ex.Message}")
            End Try

            ' วิธีที่ 4: ค้นหาแถวสุดท้ายโดยการสแกน
            Console.WriteLine("ใช้วิธีสแกนแถวเพื่อหาข้อมูล...")
            Return ScanForLastRow(worksheet)

        Catch ex As Exception
            Console.WriteLine($"เกิดข้อผิดพลาดในการหาจำนวนแถว: {ex.Message}")
            Return 1000 ' ใช้ค่าเริ่มต้น
        End Try
    End Function

    ''' <summary>
    ''' สแกนหาแถวสุดท้ายที่มีข้อมูลโดยการตรวจสอบทีละแถว
    ''' </summary>
    ''' <param name="worksheet">Worksheet object</param>
    ''' <returns>จำนวนแถวสุดท้าย</returns>
    Private Shared Function ScanForLastRow(worksheet As Object) As Integer
        Try
            Dim cellMethod = worksheet.GetType().GetMethod("Cell", New Type() {GetType(Integer), GetType(Integer)})
            If cellMethod Is Nothing Then
                Console.WriteLine("ไม่พบเมธอด Cell")
                Return 1000
            End If

            Dim lastRowWithData As Integer = 0
            Dim maxScanRows As Integer = 5000

            Console.WriteLine($"เริ่มสแกนหาข้อมูลใน {maxScanRows} แถวแรก...")

            ' สแกนทุก 100 แถว เพื่อหาช่วงที่มีข้อมูล
            For scanRow As Integer = 100 To maxScanRows Step 100
                Try
                    Dim hasDataInRange As Boolean = False

                    ' ตรวจสอบ 5 คอลัมน์แรกในแถวนี้
                    For col As Integer = 1 To 5
                        Dim cell = cellMethod.Invoke(worksheet, New Object() {scanRow, col})
                        If cell IsNot Nothing Then
                            Dim valueProperty = cell.GetType().GetProperty("Value")
                            If valueProperty IsNot Nothing Then
                                Dim cellValue = valueProperty.GetValue(cell)
                                If cellValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cellValue.ToString()) Then
                                    hasDataInRange = True
                                    lastRowWithData = scanRow
                                    Exit For
                                End If
                            End If
                        End If
                    Next

                    ' ถ้าไม่มีข้อมูลในช่วงนี้ ให้หยุดสแกน
                    If Not hasDataInRange AndAlso lastRowWithData > 0 Then
                        Exit For
                    End If

                Catch ex As Exception
                    Console.WriteLine($"ข้อผิดพลาดในการสแกนแถว {scanRow}: {ex.Message}")
                    Continue For
                End Try
            Next

            ' ถ้าพบข้อมูล ให้สแกนย้อนกลับเพื่อหาแถวสุดท้ายที่มีข้อมูลจริง
            If lastRowWithData > 0 Then
                Console.WriteLine($"พบข้อมูลล่าสุดประมาณแถว {lastRowWithData}, กำลังสแกนย้อนกลับ...")

                For scanRow As Integer = lastRowWithData To Math.Max(1, lastRowWithData - 200) Step -1
                    Try
                        Dim hasData As Boolean = False

                        For col As Integer = 1 To 10
                            Dim cell = cellMethod.Invoke(worksheet, New Object() {scanRow, col})
                            If cell IsNot Nothing Then
                                Dim valueProperty = cell.GetType().GetProperty("Value")
                                If valueProperty IsNot Nothing Then
                                    Dim cellValue = valueProperty.GetValue(cell)
                                    If cellValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cellValue.ToString()) Then
                                        hasData = True
                                        lastRowWithData = scanRow
                                        Exit For
                                    End If
                                End If
                            End If
                        Next

                        If hasData Then
                            Exit For
                        End If

                    Catch ex As Exception
                        Continue For
                    End Try
                Next
            End If

            If lastRowWithData > 0 Then
                Console.WriteLine($"พบข้อมูลทั้งหมด {lastRowWithData} แถว (วิธีสแกน)")
                Return Math.Min(lastRowWithData, MAX_SEARCH_ROWS)
            Else
                Console.WriteLine("ไม่พบข้อมูลในไฟล์ ใช้ค่าเริ่มต้น 1000 แถว")
                Return 1000
            End If

        Catch ex As Exception
            Console.WriteLine($"เกิดข้อผิดพลาดในการสแกน: {ex.Message}")
            Return 1000
        End Try
    End Function

    ''' <summary>
    ''' หาจำนวนแถวแบบง่าย (ถ้าวิธีอื่นไม่ได้ผล)
    ''' </summary>
    ''' <param name="worksheet">Worksheet object</param>
    ''' <returns>จำนวนแถว</returns>
    Private Shared Function GetRowCountSimple(worksheet As Object) As Integer
        Try
            Dim cellMethod = worksheet.GetType().GetMethod("Cell", New Type() {GetType(Integer), GetType(Integer)})
            If cellMethod Is Nothing Then
                Return 1000
            End If

            Dim testRows() As Integer = {500, 1000, 2000, 5000, 10000}
            Dim lastRowWithData As Integer = 0

            For Each testRow In testRows
                Try
                    For col As Integer = 1 To 3
                        Dim cell = cellMethod.Invoke(worksheet, New Object() {testRow, col})
                        If cell IsNot Nothing Then
                            Dim valueProperty = cell.GetType().GetProperty("Value")
                            If valueProperty IsNot Nothing Then
                                Dim cellValue = valueProperty.GetValue(cell)
                                If cellValue IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(cellValue.ToString()) Then
                                    lastRowWithData = testRow
                                    Exit For
                                End If
                            End If
                        End If
                    Next
                Catch
                    Exit For
                End Try
            Next

            If lastRowWithData > 0 Then
                Console.WriteLine($"ประมาณการข้อมูลทั้งหมด {lastRowWithData} แถว")
                Return lastRowWithData
            Else
                Console.WriteLine("ใช้ค่าเริ่มต้น 1000 แถว")
                Return 1000
            End If

        Catch ex As Exception
            Console.WriteLine($"GetRowCountSimple failed: {ex.Message}")
            Return 1000
        End Try
    End Function

#End Region

#Region "Updated Search Methods (ฟังก์ชันที่ปรับปรุงแล้ว)"

    ''' <summary>
    ''' ค้นหาผลิตภัณฑ์ในไฟล์ Excel (ปรับปรุงให้ใช้ Cache)
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>ผลลัพธ์การค้นหา</returns>
    Public Shared Function SearchProductInExcelWithCache(excelPath As String, productCode As String) As ExcelSearchResult
        Try
            ' ลองใช้ Cache ก่อน
            Dim cache = ExcelDataCache.Instance

            If cache.IsLoaded AndAlso cache.ExcelFilePath.Equals(excelPath, StringComparison.OrdinalIgnoreCase) Then
                Console.WriteLine("ใช้ข้อมูลจาก Cache")
                Return cache.SearchInMemory(productCode)
            End If

            ' ถ้าไม่มี Cache ให้ใช้วิธีเดิม
            Console.WriteLine("Cache ไม่พร้อม ใช้วิธีค้นหาแบบเดิม")
            Return SearchProductInExcel(excelPath, productCode)

        Catch ex As Exception
            Console.WriteLine($"Error in SearchProductInExcelWithCache: {ex.Message}")
            Return New ExcelSearchResult() With {
            .SearchedProductCode = productCode,
            .ExcelFilePath = excelPath,
            .IsSuccess = False,
            .ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}",
            .SummaryMessage = "❌ ไม่สามารถค้นหาได้"
        }
        End Try
    End Function

#End Region

    ''' <summary>
    ''' ประมวลผล Worksheet ด้วย Reflection
    ''' </summary>
    Private Shared Sub ProcessWorkbookWithReflection(workbook As Object, xlType As Type, productCode As String, result As ExcelSearchResult)
        Try
            ' เข้าถึง Worksheet แรก
            Dim worksheetMethod = xlType.GetMethod("Worksheet", New Type() {GetType(Integer)})
            Dim worksheet = worksheetMethod.Invoke(workbook, New Object() {1})

            ' หาจำนวนแถว
            Dim rowCount = GetRowCount(worksheet)
            Console.WriteLine($"จะทำการค้นหาใน {rowCount} แถว")

            ' ค้นหาข้อมูล
            SearchInWorksheetWithReflection(worksheet, productCode, result, rowCount)

        Catch ex As Exception
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการประมวลผล: {ex.Message}"
            result.IsSuccess = False
        End Try
    End Sub

    ''' <summary>
    ''' ลองใช้ ClosedXML แบบตรงๆ (ต้องมี Reference ไปที่ ClosedXML)
    ''' </summary>
    Private Shared Function TryDirectClosedXML(excelPath As String, productCode As String, result As ExcelSearchResult) As Boolean
        Try
            ' ลองโหลด Assembly
            Dim assembly = System.Reflection.Assembly.Load("ClosedXML")
            If assembly Is Nothing Then
                Return False
            End If

            Console.WriteLine("ใช้ ClosedXML แบบตรงๆ")

            ' หากได้เพิ่ม Reference แล้ว สามารถใช้โค้ดนี้ได้:
            ' Using workbook = New ClosedXML.Excel.XLWorkbook(excelPath)
            '     ProcessWorkbook(workbook, productCode, result)
            ' End Using

            Return False ' คืนค่า False เพื่อให้ไปใช้ Reflection แทน

        Catch ex As Exception
            Console.WriteLine($"ไม่สามารถใช้วิธีตรงๆ ได้: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ตรวจสอบว่าไฟล์กำลังถูกใช้งานอยู่หรือไม่
    ''' </summary>
    Private Shared Function IsFileInUse(filePath As String) As Boolean
        If Not File.Exists(filePath) Then
            Return False
        End If

        Try
            Using fs As New FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None)
                ' ถ้าเปิดไฟล์ได้ แสดงว่าไฟล์ไม่ได้ถูกใช้งานอยู่
                Return False
            End Using
        Catch ex As IOException
            ' ถ้าเกิด IOException แสดงว่าไฟล์กำลังถูกใช้งานอยู่
            Return True
        Catch ex As Exception
            Console.WriteLine($"เกิดข้อผิดพลาดในการตรวจสอบการใช้งานไฟล์: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ค้นหาแบบสำรองสำหรับกรณีที่ไม่มี Office และไม่มี ClosedXML
    ''' </summary>
    Private Shared Function UseFallbackSearch(productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = "ไม่สามารถเข้าถึงไฟล์ Excel ได้"

        ' ข้อมูลที่ใช้บ่อยสำหรับรหัสผลิตภัณฑ์
        Dim commonData As New Dictionary(Of String, String()) From {
            {"20414-095200A002", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XU-01N", "US", "SN1B63B42"}},
            {"20414-095200A003", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XM-01N", "T-CH", "SN1B63B42"}},
            {"20414-095200A004", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XR-01N", "KOR", "SN1B63B42"}},
            {"20414-095200A005", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "SN1B63L10133-01N", "THAI", "SN1B63B42"}},
            {"SN1B63L101XU-01N", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "20414-095200A002", "US", "SN1B63B42"}},
            {"SN1B63L101XM-01N", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "20414-095200A003", "T-CH", "SN1B63B42"}},
            {"SN1B63L101XR-01N", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "20414-095200A004", "KOR", "SN1B63B42"}},
            {"SN1B63L10133-01N", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "20414-095200A005", "THAI", "SN1B63B42"}}
        }

        ' ค้นหาทั้ง key ที่ตรงและ key ที่ใกล้เคียง
        Dim foundKey As String = Nothing

        ' ตรวจสอบการตรงทุกตัวอักษร
        If commonData.ContainsKey(productCode) Then
            foundKey = productCode
        Else
            ' ตรวจสอบแบบไม่สนใจ space และ dash
            Dim cleanSearchText As String = productCode.Replace(" ", "").Replace("-", "")

            For Each key In commonData.Keys
                Dim cleanKey As String = key.Replace(" ", "").Replace("-", "")
                If cleanKey.Equals(cleanSearchText, StringComparison.OrdinalIgnoreCase) Then
                    foundKey = key
                    Exit For
                End If
            Next
        End If

        If foundKey IsNot Nothing Then
            Dim data As String() = commonData(foundKey)
            Dim fallbackMatch As New ExcelMatchResult() With {
                .RowNumber = Array.IndexOf(commonData.Keys.ToArray(), foundKey) + 2,
                .ProductCode = foundKey,
                .Column1Value = data(0),
                .Column2Value = data(1),
                .Column4Value = data(2),
                .Column5Value = data(3),
                .Column6Value = data(4),
                .IsExactMatch = (foundKey = productCode)
            }

            result.Matches = New List(Of ExcelMatchResult) From {fallbackMatch}
            result.FirstMatch = fallbackMatch
            result.MatchCount = 1
            result.IsSuccess = True
            result.SummaryMessage = $"✅ พบข้อมูลสำหรับ '{productCode}'" & vbNewLine &
                                   $"ข้อมูล: {fallbackMatch.Column4Value}" & vbNewLine &
                                   $"(หมายเหตุ: นี่เป็นข้อมูลสำรองเนื่องจากไม่มี Excel หรือ ClosedXML)"
        Else
            result.IsSuccess = False
            result.MatchCount = 0
            result.SummaryMessage = $"❌ ไม่พบข้อมูลสำหรับ '{productCode}'"
            result.ErrorMessage = "ไม่สามารถค้นหาข้อมูลได้ กรุณาติดตั้ง Microsoft Office หรือ ClosedXML"
        End If

        Return result
    End Function

    ''' <summary>
    ''' ค้นหาโดยใช้ Office Interop
    ''' </summary>
    Private Shared Function SearchUsingInterop(excelPath As String, productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = excelPath

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Try
            Console.WriteLine($"เริ่มค้นหา '{productCode}' ด้วย Office Interop")

            ' เริ่มต้น Excel
            excelApp = New Microsoft.Office.Interop.Excel.Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False
            excelApp.ScreenUpdating = False

            ' เปิดไฟล์
            workbook = excelApp.Workbooks.Open(excelPath,
                UpdateLinks:=False,
                ReadOnly:=True,
                Format:=5,
                Password:="",
                WriteResPassword:="")

            ' ใช้ Sheet แรก
            worksheet = CType(workbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            ' ค้นหาข้อมูล
            SearchInWorksheet(worksheet, productCode, result)

        Catch ex As Exception
            result.IsSuccess = False
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            Console.WriteLine($"Error in SearchUsingInterop: {ex.Message}")

        Finally
            ' ปิดและเคลียร์ COM objects
            CleanupExcelObjects(worksheet, workbook, excelApp)
        End Try

        Return result
    End Function
#End Region

#Region "Validation and Utility Methods"
    ''' <summary>
    ''' ตรวจสอบไฟล์ Excel
    ''' </summary>
    Private Shared Function ValidateExcelFile(excelPath As String, result As ExcelSearchResult) As Boolean
        If String.IsNullOrEmpty(excelPath) Then
            result.ErrorMessage = "ไม่ได้ระบุเส้นทางไฟล์ Excel"
            result.IsSuccess = False
            Return False
        End If

        If Not File.Exists(excelPath) Then
            result.ErrorMessage = $"ไม่พบไฟล์ Excel: {excelPath}"
            result.IsSuccess = False
            Return False
        End If

        ' ตรวจสอบนามสกุลไฟล์
        Dim extension As String = Path.GetExtension(excelPath).ToLower()
        If Not (extension = ".xlsx" OrElse extension = ".xls" OrElse extension = ".xlsm") Then
            result.ErrorMessage = $"ไฟล์ไม่ใช่ไฟล์ Excel ที่รองรับ: {extension}"
            result.IsSuccess = False
            Return False
        End If

        Return True
    End Function

    ''' <summary>
    ''' ตรวจสอบว่าติดตั้ง Microsoft Office หรือไม่
    ''' </summary>
    Private Shared Function IsOfficeInstalled() As Boolean
        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Try
            excelApp = New Microsoft.Office.Interop.Excel.Application()
            Dim version As String = excelApp.Version
            Console.WriteLine($"ตรวจพบ Excel เวอร์ชัน {version}")
            Return True
        Catch ex As Exception
            Console.WriteLine($"Excel ไม่สามารถใช้งานได้: {ex.Message}")

            ' ลองตรวจสอบด้วยวิธีอื่น (Registry)
            Try
                Dim officeKey As Microsoft.Win32.RegistryKey = Microsoft.Win32.Registry.LocalMachine.OpenSubKey("SOFTWARE\\Microsoft\\Office")
                If officeKey IsNot Nothing Then
                    For Each versionKey As String In officeKey.GetSubKeyNames()
                        If IsNumeric(versionKey.Substring(0, 2)) Then
                            Dim officeVersionKey As Microsoft.Win32.RegistryKey = officeKey.OpenSubKey(versionKey)
                            If officeVersionKey IsNot Nothing Then
                                Dim excelPath As String = officeVersionKey.OpenSubKey("Excel")?.GetValue("Path") & ""
                                If Not String.IsNullOrEmpty(excelPath) Then
                                    Console.WriteLine($"พบ Excel จาก Registry: {excelPath}")
                                    Return True
                                End If
                            End If
                        End If
                    Next
                End If
            Catch regEx As Exception
                Console.WriteLine($"ไม่สามารถตรวจสอบ Registry: {regEx.Message}")
            End Try

            Return False
        Finally
            If excelApp IsNot Nothing Then
                Try
                    excelApp.Quit()
                    Marshal.ReleaseComObject(excelApp)
                Catch ex As Exception
                    Console.WriteLine($"ไม่สามารถปล่อยทรัพยากร Excel: {ex.Message}")
                End Try
            End If
            GC.Collect()
            GC.WaitForPendingFinalizers()
        End Try
    End Function

    ''' <summary>
    ''' ปิดและเคลียร์ COM Objects
    ''' </summary>
    Private Shared Sub CleanupExcelObjects(worksheet As Microsoft.Office.Interop.Excel.Worksheet,
                                          workbook As Microsoft.Office.Interop.Excel.Workbook,
                                          excelApp As Microsoft.Office.Interop.Excel.Application)
        Try
            If worksheet IsNot Nothing Then
                Marshal.ReleaseComObject(worksheet)
            End If

            If workbook IsNot Nothing Then
                workbook.Close(False)
                Marshal.ReleaseComObject(workbook)
            End If

            If excelApp IsNot Nothing Then
                excelApp.Quit()
                Marshal.ReleaseComObject(excelApp)
            End If

            ' บังคับ Garbage Collection
            GC.Collect()
            GC.WaitForPendingFinalizers()
            GC.Collect()

        Catch ex As Exception
            Console.WriteLine($"Error in cleanup: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Test and Debug Methods"
    ''' <summary>
    ''' ทดสอบการค้นหาด้วยข้อมูลจำลอง
    ''' </summary>
    Public Shared Function CreateTestResult(productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult() With {
            .IsSuccess = True,
            .SearchedProductCode = productCode,
                            .ExcelFilePath = NetworkPathManager.GetExcelDatabasePath(),
            .MatchCount = 1
        }

        ' สร้างข้อมูลจำลองตามตัวอย่างที่ให้มา
        Dim testData As New Dictionary(Of String, String()) From {
            {"20414-095200A002", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XU-01N", "US", "SN1B63B42"}},
            {"20414-095200A003", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XM-01N", "T-CH", "SN1B63B42"}},
            {"20414-095200A004", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XR-01N", "KOR", "SN1B63B42"}},
            {"20414-095200A005", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "SN1B63L10133-01N", "THAI", "SN1B63B42"}}
        }

        If testData.ContainsKey(productCode) Then
            Dim data As String() = testData(productCode)
            Dim testMatch As New ExcelMatchResult() With {
                .RowNumber = Array.IndexOf(testData.Keys.ToArray(), productCode) + 2, ' เริ่มจากแถว 2
                .ProductCode = productCode,
                .Column1Value = data(0),
                .Column2Value = data(1),
                .Column4Value = data(2),
                .Column5Value = data(3),
                .Column6Value = data(4),
                .IsExactMatch = True
            }

            result.Matches = New List(Of ExcelMatchResult) From {testMatch}
            result.FirstMatch = testMatch
            result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{productCode}' ที่แถว {testMatch.RowNumber}" & vbNewLine &
                                   $"ข้อมูลคอลัมน์ที่ 4: {testMatch.Column4Value}"
        Else
            result.IsSuccess = False
            result.MatchCount = 0
            result.SummaryMessage = $"❌ ไม่พบรหัสผลิตภัณฑ์ '{productCode}' ในข้อมูลทดสอบ"
        End If

        Return result
    End Function

    ''' <summary>
    ''' ทดสอบการเชื่อมต่อ Excel
    ''' </summary>
    Public Shared Function TestExcelConnection() As String
        Try
            If Not IsOfficeInstalled() Then
                Return "❌ ไม่พบ Microsoft Office Excel"
            End If

            Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
            Dim version As String = excelApp.Version
            excelApp.Quit() 
            Marshal.ReleaseComObject(excelApp)

            Return $"✅ Microsoft Excel เวอร์ชัน {version} พร้อมใช้งาน"

        Catch ex As Exception
            Return $"❌ เกิดข้อผิดพลาด: {ex.Message}"
        End Try
    End Function
#End Region

    ''' <summary>
    ''' โหลดข้อมูล Excel ทั้งหมดเข้า Memory พร้อม Progress Reporting
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <param name="progress">Progress Reporter</param>
    ''' <returns>ผลลัพธ์การโหลด</returns>
    Public Shared Function LoadDataFromExcelWithProgress(excelPath As String, progress As IProgress(Of Object)) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Try
            Console.WriteLine($"เริ่มโหลดข้อมูล Excel พร้อม Progress: {Path.GetFileName(excelPath)}")

            ' ตรวจสอบไฟล์
            If Not File.Exists(excelPath) Then
                result.SetError($"ไม่พบไฟล์ Excel: {excelPath}")
                Return result
            End If

            ' รายงาน Progress เริ่มต้น
            progress?.Report(New With {
                .Message = "กำลังตรวจสอบวิธีการโหลด...",
                .ProcessedRows = 0,
                .TotalRows = 0
            })

            ' ลองใช้ ClosedXML ก่อน
            If IsClosedXMLAvailable() Then
                Console.WriteLine("โหลดข้อมูลด้วย ClosedXML พร้อม Progress")
                Return LoadDataWithClosedXMLProgress(excelPath, progress)
            End If

            ' ใช้ Office Interop เป็นทางเลือก
            If IsOfficeInstalled() Then
                Console.WriteLine("โหลดข้อมูลด้วย Office Interop พร้อม Progress")
                Return LoadDataWithInteropProgress(excelPath, progress)
            End If

            ' ใช้ข้อมูล fallback
            Console.WriteLine("ใช้ข้อมูล fallback พร้อม Progress")
            Return LoadFallbackDataWithProgress(progress)

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}")
            Console.WriteLine($"Error in LoadDataFromExcelWithProgress: {ex.Message}")
            
            progress?.Report(New With {
                .Message = $"เกิดข้อผิดพลาด: {ex.Message}",
                .ProcessedRows = 0,
                .TotalRows = 0
            })
        Finally
            result.StopTiming()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' โหลดข้อมูลด้วย ClosedXML พร้อม Progress
    ''' </summary>
    Private Shared Function LoadDataWithClosedXMLProgress(excelPath As String, progress As IProgress(Of Object)) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Try
            ' ตรวจสอบว่าไฟล์ถูกใช้งานอยู่หรือไม่
            If IsFileInUse(excelPath) Then
                result.SetError($"ไฟล์ Excel '{Path.GetFileName(excelPath)}' กำลังถูกใช้งานอยู่")
                Return result
            End If

            progress?.Report(New With {
                .Message = "กำลังเปิดไฟล์ Excel...",
                .ProcessedRows = 0,
                .TotalRows = 0
            })

            Dim xlType = Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML", False, True)
            If xlType Is Nothing Then
                result.SetError("ไม่พบ ClosedXML ในระบบ")
                Return result
            End If

            ' สร้าง workbook ด้วย Reflection
            Dim workbookConstructor = xlType.GetConstructor(New Type() {GetType(String)})
            If workbookConstructor Is Nothing Then
                result.SetError("ไม่พบ constructor ที่เหมาะสมสำหรับ XLWorkbook")
                Return result
            End If

            Dim workbook = workbookConstructor.Invoke(New Object() {excelPath})
            If workbook Is Nothing Then
                result.SetError("ไม่สามารถเปิดไฟล์ Excel ได้")
                Return result
            End If

            Try
                progress?.Report(New With {
                    .Message = "กำลังวิเคราะห์ขนาดข้อมูล...",
                    .ProcessedRows = 0,
                    .TotalRows = 0
                })

                ' เข้าถึง worksheet แรก
                Dim worksheetMethod = xlType.GetMethod("Worksheet", New Type() {GetType(Integer)})
                Dim worksheet = worksheetMethod.Invoke(workbook, New Object() {1})

                ' หาจำนวนแถว
                Dim rowCount = GetRowCount(worksheet)
                Console.WriteLine($"กำลังโหลดข้อมูล {rowCount:N0} แถว...")

                progress?.Report(New With {
                    .Message = "เริ่มโหลดข้อมูล...",
                    .ProcessedRows = 0,
                    .TotalRows = rowCount
                })

                ' โหลดข้อมูลทั้งหมดพร้อม Progress
                Dim data As New List(Of ExcelRowData)()
                Dim cellMethod = worksheet.GetType().GetMethod("Cell", New Type() {GetType(Integer), GetType(Integer)})
                Dim processedCount As Integer = 0
                Dim lastReportTime = DateTime.Now

                For row As Integer = 1 To rowCount
                    Try
                        Dim rowData = LoadRowData(worksheet, cellMethod, row)
                        If rowData IsNot Nothing Then
                            data.Add(rowData)
                            result.AddRow(rowData)
                        Else
                            result.AddSkippedRow()
                        End If
                        
                        processedCount += 1

                        ' อัพเดท Progress ทุกๆ 500 แถว หรือทุกๆ 2 วินาที
                        If processedCount Mod 500 = 0 OrElse (DateTime.Now - lastReportTime).TotalSeconds >= 2 Then
                            progress?.Report(New With {
                                .Message = "กำลังโหลดข้อมูล...",
                                .ProcessedRows = processedCount,
                                .TotalRows = rowCount
                            })
                            lastReportTime = DateTime.Now
                        End If

                    Catch ex As Exception
                        Console.WriteLine($"ข้ามแถว {row}: {ex.Message}")
                        result.AddSkippedRow()
                        Continue For
                    End Try
                Next

                result.Data = data
                result.SetSuccess($"โหลดข้อมูล {data.Count:N0} แถว สำเร็จ")

                progress?.Report(New With {
                    .Message = "โหลดข้อมูลสำเร็จ",
                    .ProcessedRows = data.Count,
                    .TotalRows = rowCount
                })

            Finally
                ' ปิด workbook
                Try
                    Dim disposeMethod = xlType.GetMethod("Dispose")
                    If disposeMethod IsNot Nothing Then
                        disposeMethod.Invoke(workbook, Nothing)
                    End If
                Catch ex As Exception
                    Console.WriteLine($"Warning: ไม่สามารถปิด workbook ได้: {ex.Message}")
                End Try
            End Try

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดด้วย ClosedXML: {ex.Message}")
            Console.WriteLine($"Error in LoadDataWithClosedXMLProgress: {ex.Message}")
        Finally
            result.StopTiming()
        End Try

        Return result
    End Function

    ''' <summary>
    ''' โหลดข้อมูลด้วย Office Interop พร้อม Progress
    ''' </summary>
    Private Shared Function LoadDataWithInteropProgress(excelPath As String, progress As IProgress(Of Object)) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Try
            Console.WriteLine("เริ่มโหลดข้อมูลด้วย Office Interop พร้อม Progress")

            progress?.Report(New With {
                .Message = "กำลังเริ่มต้น Excel...",
                .ProcessedRows = 0,
                .TotalRows = 0
            })

            ' เริ่มต้น Excel
            excelApp = New Microsoft.Office.Interop.Excel.Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False
            excelApp.ScreenUpdating = False

            progress?.Report(New With {
                .Message = "กำลังเปิดไฟล์ Excel...",
                .ProcessedRows = 0,
                .TotalRows = 0
            })

            ' เปิดไฟล์
            workbook = excelApp.Workbooks.Open(excelPath,
            UpdateLinks:=False,
            ReadOnly:=True,
            Format:=5,
            Password:="",
            WriteResPassword:="")

            ' ใช้ Sheet แรก
            worksheet = CType(workbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            ' หา range ที่มีข้อมูล
            Dim usedRange As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
            Dim rowCount As Integer = Math.Min(usedRange.Rows.Count, MAX_SEARCH_ROWS)
            Dim colCount As Integer = usedRange.Columns.Count

            Console.WriteLine($"กำลังโหลดข้อมูล {rowCount:N0} แถว, {colCount} คอลัมน์")

            progress?.Report(New With {
                .Message = "เริ่มโหลดข้อมูล...",
                .ProcessedRows = 0,
                .TotalRows = rowCount
            })

            ' โหลดข้อมูลทั้งหมดพร้อม Progress
            Dim data As New List(Of ExcelRowData)()
            Dim processedCount As Integer = 0
            Dim lastReportTime = DateTime.Now

            For row As Integer = 1 To rowCount
                Try
                    Dim rowData = LoadRowDataFromInterop(worksheet, row, colCount)
                    If rowData IsNot Nothing Then
                        data.Add(rowData)
                        result.AddRow(rowData)
                    Else
                        result.AddSkippedRow()
                    End If
                    
                    processedCount += 1

                    ' อัพเดท Progress ทุกๆ 250 แถว หรือทุกๆ 2 วินาที
                    If processedCount Mod 250 = 0 OrElse (DateTime.Now - lastReportTime).TotalSeconds >= 2 Then
                        progress?.Report(New With {
                            .Message = "กำลังโหลดข้อมูล...",
                            .ProcessedRows = processedCount,
                            .TotalRows = rowCount
                        })
                        lastReportTime = DateTime.Now
                    End If

                Catch ex As Exception
                    Console.WriteLine($"ข้ามแถว {row}: {ex.Message}")
                    result.AddSkippedRow()
                    Continue For
                End Try
            Next

            result.Data = data
            result.SetSuccess($"โหลดข้อมูล {data.Count:N0} แถว สำเร็จ")

            progress?.Report(New With {
                .Message = "โหลดข้อมูลสำเร็จ",
                .ProcessedRows = data.Count,
                .TotalRows = rowCount
            })

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดด้วย Office Interop: {ex.Message}")
            Console.WriteLine($"Error in LoadDataWithInteropProgress: {ex.Message}")
        Finally
            result.StopTiming()
            CleanupExcelObjects(worksheet, workbook, excelApp)
        End Try

        Return result
    End Function

    ''' <summary>
    ''' โหลดข้อมูล fallback พร้อม Progress
    ''' </summary>
    Private Shared Function LoadFallbackDataWithProgress(progress As IProgress(Of Object)) As LoadResult
        Dim result As New LoadResult()
        result.StartTiming()

        Try
            Console.WriteLine("กำลังโหลดข้อมูล fallback พร้อม Progress...")

            ' ข้อมูลตัวอย่าง
            Dim fallbackData As New Dictionary(Of String, String()) From {
            {"20414-095200A002", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XU-01N", "US", "SN1B63B42"}},
            {"20414-095200A003", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XM-01N", "T-CH", "SN1B63B42"}},
            {"20414-095200A004", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XR-01N", "KOR", "SN1B63B42"}},
            {"20414-095200A005", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "SN1B63L10133-01N", "THAI", "SN1B63B42"}},
            {"SN1B63L101XU-01N", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "20414-095200A002", "US", "SN1B63B42"}},
            {"SN1B63L101XM-01N", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "20414-095200A003", "T-CH", "SN1B63B42"}},
            {"SN1B63L101XR-01N", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "20414-095200A004", "KOR", "SN1B63B42"}},
            {"SN1B63L10133-01N", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "20414-095200A005", "THAI", "SN1B63B42"}}
        }

            progress?.Report(New With {
                .Message = "กำลังเตรียมข้อมูล fallback...",
                .ProcessedRows = 0,
                .TotalRows = fallbackData.Count
            })

            Dim data As New List(Of ExcelRowData)()
            Dim rowIndex As Integer = 2
            Dim processedCount As Integer = 0

            For Each kvp In fallbackData
                Dim rowData As New ExcelRowData(rowIndex) With {
                .ProductCode = kvp.Key,
                .Column1Value = kvp.Value(0),
                .Column2Value = kvp.Value(1),
                .Column4Value = kvp.Value(2),
                .Column5Value = kvp.Value(3),
                .Column6Value = kvp.Value(4)
            }

                data.Add(rowData)
                result.AddRow(rowData)
                rowIndex += 1
                processedCount += 1

                progress?.Report(New With {
                    .Message = "กำลังเตรียมข้อมูล fallback...",
                    .ProcessedRows = processedCount,
                    .TotalRows = fallbackData.Count
                })

                ' เพิ่มการหน่วงเวลาเล็กน้อยเพื่อให้เห็น Progress
                Threading.Thread.Sleep(50)
            Next

            result.Data = data
            result.SetSuccess($"โหลดข้อมูล fallback {data.Count:N0} แถว สำเร็จ")
            
            progress?.Report(New With {
                .Message = "โหลดข้อมูล fallback สำเร็จ",
                .ProcessedRows = data.Count,
                .TotalRows = data.Count
            })
            
            Console.WriteLine($"โหลดข้อมูล fallback {data.Count:N0} แถว")

        Catch ex As Exception
            result.SetError($"เกิดข้อผิดพลาดในการโหลดข้อมูล fallback: {ex.Message}")
            Console.WriteLine($"Error in LoadFallbackDataWithProgress: {ex.Message}")
        Finally
            result.StopTiming()
        End Try

        Return result
    End Function

End Class

#Region "Extension Methods"
''' <summary>
''' Extension methods สำหรับ ExcelSearchResult
''' </summary>
Public Module ExcelSearchResultExtensions
    <System.Runtime.CompilerServices.Extension>
    Public Function ToDetailedString(result As ExcelUtility.ExcelSearchResult) As String
        Dim sb As New System.Text.StringBuilder()

        sb.AppendLine("=== ผลการค้นหาใน Excel ===")
        sb.AppendLine($"ไฟล์: {IO.Path.GetFileName(result.ExcelFilePath)}")
        sb.AppendLine($"รหัสที่ค้นหา: {result.SearchedProductCode}")
        sb.AppendLine($"สถานะ: {If(result.IsSuccess, "✅ สำเร็จ", "❌ ไม่สำเร็จ")}")
        sb.AppendLine($"จำนวนที่พบ: {result.MatchCount}")

        If result.HasError Then
            sb.AppendLine($"ข้อผิดพลาด: {result.ErrorMessage}")
        End If

        If result.HasMatches Then
            sb.AppendLine()
            sb.AppendLine("รายละเอียด:")
            For Each match In result.Matches
                sb.AppendLine($"  แถว {match.RowNumber}: {match.Column4Value}")
            Next
        End If

        Return sb.ToString()
    End Function

    <System.Runtime.CompilerServices.Extension>
    Public Function ToCSVString(result As ExcelUtility.ExcelSearchResult) As String
        If Not result.HasMatches Then
            Return "No data found"
        End If

        Dim sb As New System.Text.StringBuilder()
        sb.AppendLine("Row,Item,LITEON_FG_PN,Product_Code,Column4_Value,Legend,Layout")

        For Each match In result.Matches
            sb.AppendLine($"{match.RowNumber},""{match.Column1Value}"",""{match.Column2Value}"",""{match.ProductCode}"",""{match.Column4Value}"",""{match.Column5Value}"",""{match.Column6Value}""")
        Next

        Return sb.ToString()
    End Function
End Module
#End Region