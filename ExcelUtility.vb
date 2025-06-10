Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel
Imports System.IO
Imports System.Text
Imports System.Text.RegularExpressions

''' <summary>
''' คลาสสำหรับจัดการการทำงานกับไฟล์ Excel
''' </summary>
Public Class ExcelUtility

#Region "Constants"
    Private Const PRODUCT_CODE_COLUMN As Integer = 3  ' คอลัมน์ที่ 3 (รหัสผลิตภัณฑ์)
    Private Const RESULT_COLUMN As Integer = 4        ' คอลัมน์ที่ 4 (ผลลัพธ์ที่ต้องการ)
    Private Const MAX_SEARCH_ROWS As Integer = 10000  ' จำนวนแถวสูงสุดที่จะค้นหา
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

        ' เลือกวิธีการค้นหาตามสถานการณ์
        If IsOfficeInstalled() Then
            Return SearchUsingInterop(excelPath, productCode)
        Else
            ' ใช้วิธีอื่นเมื่อไม่มี Office
            result.ErrorMessage = "ไม่พบ Microsoft Office Excel ในเครื่อง กรุณาติดตั้ง Microsoft Office"
            result.IsSuccess = False
            Return result
        End If
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
        Try
            Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
            excelApp.Quit()
            Marshal.ReleaseComObject(excelApp)
            Return True
        Catch
            Return False
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
                .productCode = productCode,
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
    Public Function ToDetailedString(result As ExcelSearchResult) As String
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
    Public Function ToCSVString(result As ExcelSearchResult) As String
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