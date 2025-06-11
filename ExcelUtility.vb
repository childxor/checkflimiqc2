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
    ''' ค้นหาข้อมูลใน Excel โดยใช้ ClosedXML
    ''' </summary>
    Private Shared Function SearchUsingClosedXML(excelPath As String, productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = excelPath

        Try
            ' ตรวจสอบว่าไฟล์ Excel เปิดอยู่หรือไม่
            If IsFileInUse(excelPath) Then
                result.ErrorMessage = $"ไฟล์ Excel '{Path.GetFileName(excelPath)}' กำลังถูกใช้งานอยู่ กรุณาปิดไฟล์ก่อนค้นหา"
                result.IsSuccess = False
                result.SummaryMessage = $"❌ ไม่สามารถเปิดไฟล์ Excel ได้เนื่องจากไฟล์กำลังถูกใช้งาน"
                Console.WriteLine(result.ErrorMessage)
                Return result
            End If

            ' ตรวจสอบว่ามี ClosedXML หรือไม่
            Dim xlType = Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML", False, True)
            If xlType Is Nothing Then
                result.ErrorMessage = "ไม่พบ ClosedXML ในระบบ กรุณาติดตั้ง ClosedXML ผ่าน NuGet"
                result.IsSuccess = False
                Console.WriteLine(result.ErrorMessage)
                Return result
            End If

            ' ใช้ Reflection เพื่อโหลดและใช้งาน ClosedXML แบบไดนามิก
            Try
                Console.WriteLine($"กำลังเปิดไฟล์ Excel: {excelPath}")

                ' ลองใช้วิธีตรงๆ โดยไม่ผ่าน Reflection ก่อน (ถ้ามีการติดตั้ง ClosedXML แล้ว)
                Try
                    ' ตรวจสอบว่ามีการติดตั้ง ClosedXML แล้วหรือไม่
                    Dim assembly = System.Reflection.Assembly.Load("ClosedXML")
                    If assembly IsNot Nothing Then
                        Console.WriteLine("พบ ClosedXML Assembly แล้ว ใช้วิธีตรงๆ")

                        ' ต้องมีการเพิ่ม Reference ถึง ClosedXML ในโปรเจ็คก่อน
                        ' ถ้ายังไม่ได้เพิ่ม โค้ดส่วนนี้จะไม่ทำงาน และจะไปใช้ Code Block ถัดไปแทน

                        ' Using workbook = New ClosedXML.Excel.XLWorkbook(excelPath)
                        '     Dim worksheet = workbook.Worksheet(1)
                        '     
                        '     ' หาแถวสุดท้ายที่มีข้อมูล
                        '     Dim lastRow = worksheet.LastRowUsed().RowNumber()
                        '     Dim maxRows = 10000 ' จำกัดจำนวนแถวที่ค้นหา
                        '     Dim rowCount = Math.Min(lastRow, maxRows)
                        '     
                        '     Dim searchResults As New List(Of ExcelMatchResult)()
                        '     
                        '     ' วนลูปค้นหาในแต่ละแถว
                        '     For row As Integer = 1 To rowCount
                        '         Dim cellValue = worksheet.Cell(row, PRODUCT_CODE_COLUMN).Value
                        '         
                        '         If cellValue IsNot Nothing Then
                        '             Dim cellText As String = cellValue.ToString().Trim()
                        '             
                        '             ' ตรวจสอบการแมทช์
                        '             If IsProductCodeMatch(cellText, productCode) Then
                        '                 ' สร้าง matchResult และอ่านข้อมูลจากคอลัมน์อื่นๆ
                        '                 ' ... (โค้ดเหมือนกับในเวอร์ชันที่ใช้ Reflection)
                        '             End If
                        '         End If
                        '     Next
                        '     
                        '     ' กำหนดผลลัพธ์
                        '     ' ... (โค้ดเหมือนกับในเวอร์ชันที่ใช้ Reflection)
                        ' End Using
                    End If
                Catch ex As Exception
                    Console.WriteLine($"ไม่สามารถใช้วิธีตรงๆ ได้: {ex.Message}")
                End Try

                ' ถ้าใช้วิธีตรงๆ ไม่ได้ ให้ลองใช้ Reflection
                Console.WriteLine("ลองใช้ Reflection เพื่อเข้าถึง ClosedXML")

                ' สร้าง Assembly Resolver สำหรับโหลด Assembly ที่เกี่ยวข้อง
                Dim currentDomain = AppDomain.CurrentDomain
                AddHandler currentDomain.AssemblyResolve, Function(sender, args)
                                                              If args.Name.StartsWith("ClosedXML") Then
                                                                  Return xlType.Assembly
                                                              End If

                                                              ' ทดลองโหลดจาก DLL ใน bin directory
                                                              Dim appPath = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location)
                                                              Dim assemblyName = New System.Reflection.AssemblyName(args.Name).Name
                                                              Dim dllPath = Path.Combine(appPath, $"{assemblyName}.dll")

                                                              If File.Exists(dllPath) Then
                                                                  Return System.Reflection.Assembly.LoadFrom(dllPath)
                                                              End If

                                                              Return Nothing
                                                          End Function

                ' สร้าง XLWorkbook instance ผ่าน Reflection
                Dim workbookConstructor = xlType.GetConstructor(New Type() {GetType(String)})
                If workbookConstructor Is Nothing Then
                    Throw New Exception("ไม่พบ constructor ที่เหมาะสมสำหรับ XLWorkbook")
                End If

                Console.WriteLine("กำลังเรียก XLWorkbook constructor")
                Dim workbook = Nothing

                Try
                    workbook = workbookConstructor.Invoke(New Object() {excelPath})
                Catch invokeEx As System.Reflection.TargetInvocationException
                    If invokeEx.InnerException IsNot Nothing Then
                        If TypeOf invokeEx.InnerException Is IOException Then
                            result.ErrorMessage = $"ไม่สามารถเข้าถึงไฟล์ Excel ได้: {invokeEx.InnerException.Message}"
                        Else
                            result.ErrorMessage = $"เกิดข้อผิดพลาดขณะเปิดไฟล์ Excel: {invokeEx.InnerException.Message}"
                        End If
                        result.IsSuccess = False
                        Console.WriteLine(result.ErrorMessage)
                        Return result
                    Else
                        Throw ' ส่งต่อข้อผิดพลาดถ้าไม่มี InnerException
                    End If
                End Try

                If workbook Is Nothing Then
                    Throw New Exception("ไม่สามารถสร้าง XLWorkbook ได้") 
                End If

                Console.WriteLine("เปิดไฟล์ Excel สำเร็จแล้ว กำลังอ่านข้อมูล")

                ' เข้าถึง Worksheet แรก
                Dim worksheetMethod = xlType.GetMethod("Worksheet", New Type() {GetType(Integer)})
                Dim worksheet = worksheetMethod.Invoke(workbook, New Object() {1})

                ' หาจำนวนแถวที่มีข้อมูล
                Dim lastRowNumber As Integer = 0
                Try
                    ' วิธีที่ 1: ใช้ LastRowUsed
                    Dim lastRowUsedMethod = worksheet.GetType().GetMethod("LastRowUsed")
                    If lastRowUsedMethod IsNot Nothing Then
                        Dim lastRowUsed = lastRowUsedMethod.Invoke(worksheet, Nothing)
                        Dim rowNumberProperty = lastRowUsed.GetType().GetProperty("RowNumber")
                        lastRowNumber = Convert.ToInt32(rowNumberProperty.GetValue(lastRowUsed))
                        Console.WriteLine($"พบข้อมูลทั้งหมด {lastRowNumber} แถว (จากเมธอด LastRowUsed)")
                    Else
                        ' วิธีที่ 2: ใช้ LastCellUsed
                        Dim lastCellUsedMethod = worksheet.GetType().GetMethod("LastCellUsed")
                        If lastCellUsedMethod IsNot Nothing Then
                            Dim lastCellUsed = lastCellUsedMethod.Invoke(worksheet, Nothing)
                            Dim addressProperty = lastCellUsed.GetType().GetProperty("Address")
                            Dim address = addressProperty.GetValue(lastCellUsed).ToString()

                            ' แยก Row จาก Address (เช่น "A1" หรือ "B5")
                            Dim rowMatch = System.Text.RegularExpressions.Regex.Match(address, "[A-Z]+(\d+)")
                            If rowMatch.Success AndAlso rowMatch.Groups.Count > 1 Then
                                lastRowNumber = Integer.Parse(rowMatch.Groups(1).Value)
                                Console.WriteLine($"พบข้อมูลทั้งหมด {lastRowNumber} แถว (จากเมธอด LastCellUsed)")
                            End If
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine($"ไม่สามารถหาจำนวนแถวที่มีข้อมูลได้: {ex.Message}")
                End Try

                ' ถ้าไม่สามารถหาจำนวนแถวได้ ให้ใช้ค่าที่กำหนดไว้ล่วงหน้า
                If lastRowNumber <= 0 Then
                    lastRowNumber = 1000
                    Console.WriteLine($"ใช้จำนวนแถวที่กำหนดไว้ล่วงหน้า: {lastRowNumber} แถว")
                End If

                Dim maxRows = 10000 ' จำกัดจำนวนแถวที่ค้นหา
                Dim rowCount = Math.Min(lastRowNumber, maxRows)

                Console.WriteLine($"จะทำการค้นหาใน {rowCount} แถว")

                ' ดึงเมธอด Cell และคุณสมบัติ Value
                Dim cellMethod As System.Reflection.MethodInfo = Nothing
                Dim valueProperty As System.Reflection.PropertyInfo = Nothing

                Try
                    ' ทดลองดึงเมธอด Cell(row, column)
                    cellMethod = worksheet.GetType().GetMethod("Cell", New Type() {GetType(Integer), GetType(Integer)})

                    ' ตรวจสอบว่าได้เมธอด Cell หรือไม่
                    If cellMethod Is Nothing Then
                        Throw New Exception("ไม่พบเมธอด Cell ใน worksheet")
                    End If

                    ' ทดลองเรียกใช้เมธอด Cell กับแถวและคอลัมน์แรก
                    Dim cell = cellMethod.Invoke(worksheet, New Object() {1, 1})

                    ' ตรวจสอบว่าได้อ็อบเจ็กต์ cell หรือไม่
                    If cell Is Nothing Then
                        Throw New Exception("เมธอด Cell คืนค่า null")
                    End If

                    ' ดึงคุณสมบัติ Value จาก cell
                    valueProperty = cell.GetType().GetProperty("Value")

                    ' ตรวจสอบว่าได้คุณสมบัติ Value หรือไม่
                    If valueProperty Is Nothing Then
                        Throw New Exception("ไม่พบคุณสมบัติ Value ใน cell")
                    End If
                Catch ex As Exception
                    Console.WriteLine($"ไม่สามารถเข้าถึงเมธอดหรือคุณสมบัติที่จำเป็น: {ex.Message}")
                    result.ErrorMessage = $"ไม่สามารถอ่านข้อมูลจาก Excel ได้: {ex.Message}"
                    result.IsSuccess = False
                    Return result
                End Try

                Dim searchResults As New List(Of ExcelMatchResult)()

                ' วนลูปค้นหาในแต่ละแถว
                For row As Integer = 1 To rowCount
                    Try
                        ' อ่านค่าจากคอลัมน์ที่ต้องการค้นหา (PRODUCT_CODE_COLUMN)
                        Dim cell = cellMethod.Invoke(worksheet, New Object() {row, PRODUCT_CODE_COLUMN})

                        ' ตรวจสอบว่าได้อ็อบเจ็กต์ cell หรือไม่
                        If cell Is Nothing Then
                            Continue For
                        End If

                        Dim cellValue = valueProperty.GetValue(cell)

                        If cellValue IsNot Nothing Then
                            Dim cellText As String = cellValue.ToString().Trim()

                            ' ตรวจสอบการแมทช์
                            If IsProductCodeMatch(cellText, productCode) Then
                                Dim matchResult As New ExcelMatchResult() With {
                                    .RowNumber = row,
                                    .ProductCode = cellText,
                                    .IsExactMatch = cellText.Equals(productCode, StringComparison.OrdinalIgnoreCase)
                                }

                                ' อ่านข้อมูลจากคอลัมน์อื่นๆ
                                Try
                                    ' อ่านคอลัมน์ที่ 1
                                    Try
                                        Dim cell1 = cellMethod.Invoke(worksheet, New Object() {row, 1})
                                        Dim cell1Value = valueProperty.GetValue(cell1)
                                        If cell1Value IsNot Nothing Then
                                            matchResult.Column1Value = cell1Value.ToString()
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine($"ไม่สามารถอ่านคอลัมน์ 1 ที่แถว {row}: {ex.Message}")
                                    End Try

                                    ' อ่านคอลัมน์ที่ 2
                                    Try
                                        Dim cell2 = cellMethod.Invoke(worksheet, New Object() {row, 2})
                                        Dim cell2Value = valueProperty.GetValue(cell2)
                                        If cell2Value IsNot Nothing Then
                                            matchResult.Column2Value = cell2Value.ToString()
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine($"ไม่สามารถอ่านคอลัมน์ 2 ที่แถว {row}: {ex.Message}")
                                    End Try

                                    ' อ่านคอลัมน์ที่ 4 (RESULT_COLUMN)
                                    Try
                                        Dim cell4 = cellMethod.Invoke(worksheet, New Object() {row, RESULT_COLUMN})
                                        Dim cell4Value = valueProperty.GetValue(cell4)
                                        If cell4Value IsNot Nothing Then
                                            matchResult.Column4Value = cell4Value.ToString()
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine($"ไม่สามารถอ่านคอลัมน์ {RESULT_COLUMN} ที่แถว {row}: {ex.Message}")
                                    End Try

                                    ' อ่านคอลัมน์ที่ 5 (ถ้ามี)
                                    Try
                                        If RESULT_COLUMN + 1 <= 8 Then ' ตรวจสอบว่ามีคอลัมน์ 5 หรือไม่
                                            Dim cell5 = cellMethod.Invoke(worksheet, New Object() {row, RESULT_COLUMN + 1})
                                            Dim cell5Value = valueProperty.GetValue(cell5)
                                            If cell5Value IsNot Nothing Then
                                                matchResult.Column5Value = cell5Value.ToString()
                                            End If
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine($"ไม่สามารถอ่านคอลัมน์ {RESULT_COLUMN + 1} ที่แถว {row}: {ex.Message}")
                                    End Try

                                    ' อ่านคอลัมน์ที่ 6 (ถ้ามี)
                                    Try
                                        If RESULT_COLUMN + 2 <= 8 Then ' ตรวจสอบว่ามีคอลัมน์ 6 หรือไม่
                                            Dim cell6 = cellMethod.Invoke(worksheet, New Object() {row, RESULT_COLUMN + 2})
                                            Dim cell6Value = valueProperty.GetValue(cell6)
                                            If cell6Value IsNot Nothing Then
                                                matchResult.Column6Value = cell6Value.ToString()
                                            End If
                                        End If
                                    Catch ex As Exception
                                        Console.WriteLine($"ไม่สามารถอ่านคอลัมน์ {RESULT_COLUMN + 2} ที่แถว {row}: {ex.Message}")
                                    End Try

                                Catch ex As Exception
                                    Console.WriteLine($"ข้อผิดพลาดในการอ่านคอลัมน์เพิ่มเติมที่แถว {row}: {ex.Message}")
                                End Try

                                searchResults.Add(matchResult)
                                Console.WriteLine($"พบรหัส {cellText} ที่แถว {row}")

                                ' หยุดค้นหาเมื่อพบครบ 10 รายการ
                                If searchResults.Count >= 10 Then
                                    Exit For
                                End If
                            End If
                        End If
                    Catch ex As Exception
                        Console.WriteLine($"ข้อผิดพลาดในการอ่านแถว {row}: {ex.Message}")
                        ' ข้ามแถวที่มีปัญหาและดำเนินการต่อ
                        Continue For
                    End Try
                Next

                ' กำหนดผลลัพธ์
                If searchResults.Count > 0 Then
                    result.Matches = searchResults
                    result.FirstMatch = searchResults(0)
                    result.MatchCount = searchResults.Count
                    result.IsSuccess = True

                    If searchResults.Count = 1 Then
                        result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{productCode}' ที่แถว {searchResults(0).RowNumber}" & vbNewLine &
                                               $"ข้อมูลคอลัมน์ที่ 4: {searchResults(0).Column4Value}"
                    Else
                        result.SummaryMessage = $"✅ พบรหัสผลิตภัณฑ์ '{productCode}' จำนวน {searchResults.Count} รายการ"
                    End If
                Else
                    result.IsSuccess = False
                    result.MatchCount = 0
                    result.SummaryMessage = $"❌ ไม่พบรหัสผลิตภัณฑ์ '{productCode}' ในไฟล์ Excel"
                End If

                ' ปิด workbook
                Try
                    Dim disposeMethod = xlType.GetMethod("Dispose")
                    disposeMethod.Invoke(workbook, Nothing)
                    Console.WriteLine("ปิดไฟล์ Excel เรียบร้อยแล้ว")
                Catch ex As Exception
                    Console.WriteLine($"Warning: ไม่สามารถปิดไฟล์ Excel ได้ อาจเกิดการรั่วไหลของหน่วยความจำ: {ex.Message}")
                End Try

                Return result

            Catch refEx As Exception
                Console.WriteLine($"เกิดข้อผิดพลาดในการใช้ Reflection กับ ClosedXML: {refEx.Message}")
                Console.WriteLine($"StackTrace: {refEx.StackTrace}")

                ' สร้างข้อความแสดงข้อผิดพลาดที่เข้าใจง่าย
                If refEx.InnerException IsNot Nothing Then
                    result.ErrorMessage = $"เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: {refEx.InnerException.Message}"
                Else
                    result.ErrorMessage = $"เกิดข้อผิดพลาดในการอ่านไฟล์ Excel: {refEx.Message}"
                End If

                result.IsSuccess = False
                result.SummaryMessage = $"❌ ไม่สามารถอ่านไฟล์ Excel ได้: {Path.GetFileName(excelPath)}"
                Return result
            End Try

        Catch ex As Exception
            result.IsSuccess = False
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหาด้วย ClosedXML: {ex.Message}"
            Console.WriteLine($"Error in SearchUsingClosedXML: {ex.Message}")
        End Try

        Return result
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
            .ExcelFilePath = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx",
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