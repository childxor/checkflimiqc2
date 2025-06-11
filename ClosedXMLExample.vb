Imports System.IO
Imports System.Collections.Generic

''' <summary>
''' ตัวอย่างการใช้งาน ClosedXML อย่างง่าย
''' </summary>
Public Class ClosedXMLExample

    ''' <summary>
    ''' ค้นหาข้อมูลในไฟล์ Excel ด้วย ClosedXML
    ''' </summary>
    ''' <param name="excelPath">เส้นทางไฟล์ Excel</param>
    ''' <param name="searchText">ข้อความที่ต้องการค้นหา</param>
    ''' <param name="searchColumn">คอลัมน์ที่ต้องการค้นหา (เริ่มจาก 1)</param>
    ''' <returns>ผลลัพธ์การค้นหา</returns>
    Public Shared Function SearchInExcel(excelPath As String, searchText As String, searchColumn As Integer) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = searchText
        result.ExcelFilePath = excelPath
        
        ' ตรวจสอบไฟล์
        If Not File.Exists(excelPath) Then
            result.ErrorMessage = $"ไม่พบไฟล์ Excel: {excelPath}"
            result.IsSuccess = False
            Return result
        End If
        
        Try
            ' ต้องติดตั้ง ClosedXML ก่อนโดยใช้คำสั่ง:
            ' Install-Package ClosedXML
            
            ' ตรวจสอบว่ามี ClosedXML หรือไม่
            Dim xlType = Type.GetType("ClosedXML.Excel.XLWorkbook, ClosedXML", False, True)
            If xlType Is Nothing Then
                result.ErrorMessage = "กรุณาติดตั้ง ClosedXML ก่อนใช้งาน (ดูวิธีใน README.md)"
                result.IsSuccess = False
                Return result
            End If
            
            ' ใช้ตัวแปลโค้ดแบบ Dynamic
            ' ถ้าติดตั้ง ClosedXML แล้ว ให้ลบคอมเมนต์ส่วนนี้และคอมเมนต์ส่วนด้านล่างแทน
            
            ' Using workbook = New ClosedXML.Excel.XLWorkbook(excelPath)
            '     ' ใช้ Sheet แรก
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
            '         Dim cellValue = worksheet.Cell(row, searchColumn).Value
            '         
            '         If cellValue IsNot Nothing Then
            '             Dim cellText As String = cellValue.ToString().Trim()
            '             
            '             ' ตรวจสอบการแมทช์
            '             If cellText.Equals(searchText, StringComparison.OrdinalIgnoreCase) OrElse 
            '                cellText.Replace(" ", "").Replace("-", "").Equals(searchText.Replace(" ", "").Replace("-", ""), StringComparison.OrdinalIgnoreCase) Then
            '                 
            '                 Dim matchResult = New ExcelMatchResult() With {
            '                     .RowNumber = row,
            '                     .ProductCode = cellText,
            '                     .IsExactMatch = cellText.Equals(searchText, StringComparison.OrdinalIgnoreCase)
            '                 }
            '                 
            '                 ' อ่านข้อมูลจากคอลัมน์อื่นๆ (ตัวอย่างอ่านคอลัมน์ที่ 1, 2, 4, 5, 6)
            '                 If worksheet.Cell(row, 1).Value IsNot Nothing Then
            '                     matchResult.Column1Value = worksheet.Cell(row, 1).Value.ToString()
            '                 End If
            '                 
            '                 If worksheet.Cell(row, 2).Value IsNot Nothing Then
            '                     matchResult.Column2Value = worksheet.Cell(row, 2).Value.ToString()
            '                 End If
            '                 
            '                 If worksheet.Cell(row, 4).Value IsNot Nothing Then
            '                     matchResult.Column4Value = worksheet.Cell(row, 4).Value.ToString()
            '                 End If
            '                 
            '                 If worksheet.Cell(row, 5).Value IsNot Nothing Then
            '                     matchResult.Column5Value = worksheet.Cell(row, 5).Value.ToString()
            '                 End If
            '                 
            '                 If worksheet.Cell(row, 6).Value IsNot Nothing Then
            '                     matchResult.Column6Value = worksheet.Cell(row, 6).Value.ToString()
            '                 End If
            '                 
            '                 searchResults.Add(matchResult)
            '                 
            '                 ' หยุดค้นหาเมื่อพบครบ 10 รายการ
            '                 If searchResults.Count >= 10 Then
            '                     Exit For
            '                 End If
            '             End If
            '         End If
            '     Next
            '     
            '     ' กำหนดผลลัพธ์
            '     If searchResults.Count > 0 Then
            '         result.Matches = searchResults
            '         result.FirstMatch = searchResults(0)
            '         result.MatchCount = searchResults.Count
            '         result.IsSuccess = True
            '         
            '         If searchResults.Count = 1 Then
            '             result.SummaryMessage = $"✅ พบ '{searchText}' ที่แถว {searchResults(0).RowNumber}" & vbNewLine &
            '                                   $"ข้อมูล: {searchResults(0).Column4Value}"
            '         Else
            '             result.SummaryMessage = $"✅ พบ '{searchText}' จำนวน {searchResults.Count} รายการ"
            '         End If
            '     Else
            '         result.IsSuccess = False
            '         result.MatchCount = 0
            '         result.SummaryMessage = $"❌ ไม่พบ '{searchText}' ในไฟล์ Excel"
            '     End If
            ' End Using
            
            ' ตัวอย่างการจำลองข้อมูล (ใช้ในกรณีที่ยังไม่ได้ติดตั้ง ClosedXML)
            ' เมื่อติดตั้ง ClosedXML แล้ว ให้ลบส่วนนี้และใช้โค้ดด้านบนแทน
            Dim demoData As New Dictionary(Of String, String()) From {
                {"20414-095200A002", {"SG-C1010-XUA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XU-01N", "US", "SN1B63B42"}},
                {"20414-095200A003", {"SG-C1010-XMA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XM-01N", "T-CH", "SN1B63B42"}},
                {"20414-095200A004", {"SG-C1010-XRA", "LITEON FG PN painting keycaps part no.", "SN1B63L101XR-01N", "KOR", "SN1B63B42"}},
                {"20414-095200A005", {"SG-C1010-33A", "LITEON FG PN painting keycaps part no.", "SN1B63L10133-01N", "THAI", "SN1B63B42"}}
            }
            
            ' ค้นหาทั้ง key ที่ตรงและ key ที่ใกล้เคียง
            Dim foundKey As String = Nothing
            
            ' ตรวจสอบการตรงทุกตัวอักษร
            If demoData.ContainsKey(searchText) Then
                foundKey = searchText
            Else
                ' ตรวจสอบแบบไม่สนใจ space และ dash
                Dim cleanSearchText As String = searchText.Replace(" ", "").Replace("-", "")
                
                For Each key In demoData.Keys
                    Dim cleanKey As String = key.Replace(" ", "").Replace("-", "")
                    If cleanKey.Equals(cleanSearchText, StringComparison.OrdinalIgnoreCase) Then
                        foundKey = key
                        Exit For
                    End If
                Next
            End If
            
            If foundKey IsNot Nothing Then
                Dim data As String() = demoData(foundKey)
                Dim demoMatch As New ExcelMatchResult() With {
                    .RowNumber = Array.IndexOf(demoData.Keys.ToArray(), foundKey) + 2,
                    .ProductCode = foundKey,
                    .Column1Value = data(0),
                    .Column2Value = data(1),
                    .Column4Value = data(2),
                    .Column5Value = data(3),
                    .Column6Value = data(4),
                    .IsExactMatch = (foundKey = searchText)
                }
                
                result.Matches = New List(Of ExcelMatchResult) From {demoMatch}
                result.FirstMatch = demoMatch
                result.MatchCount = 1
                result.IsSuccess = True
                result.SummaryMessage = $"✅ พบ '{searchText}' ที่แถว {demoMatch.RowNumber}" & vbNewLine &
                                       $"ข้อมูล: {demoMatch.Column4Value}" & vbNewLine &
                                       $"(หมายเหตุ: นี่เป็นข้อมูลจำลอง)"
            Else
                result.IsSuccess = False
                result.MatchCount = 0
                result.SummaryMessage = $"❌ ไม่พบ '{searchText}' ในข้อมูลจำลอง"
            End If
            
        Catch ex As Exception
            result.IsSuccess = False
            result.ErrorMessage = $"เกิดข้อผิดพลาด: {ex.Message}"
        End Try
        
        Return result
    End Function
    
    ''' <summary>
    ''' ตัวอย่างการใช้งาน
    ''' </summary>
    Public Shared Sub ExampleUsage()
        ' ตัวอย่างการใช้งาน
        Dim excelPath As String = "C:\path\to\your\file.xlsx"
        Dim searchText As String = "20414-095200A002"
        Dim searchColumn As Integer = 3 ' คอลัมน์ C
        
        Dim result = SearchInExcel(excelPath, searchText, searchColumn)
        
        If result.IsSuccess Then
            Console.WriteLine(result.SummaryMessage)
            
            ' แสดงข้อมูลเพิ่มเติม
            If result.FirstMatch IsNot Nothing Then
                Console.WriteLine($"รายละเอียด:")
                Console.WriteLine($"- Item: {result.FirstMatch.Column1Value}")
                Console.WriteLine($"- Description: {result.FirstMatch.Column2Value}")
                Console.WriteLine($"- Result: {result.FirstMatch.Column4Value}")
                Console.WriteLine($"- Legend: {result.FirstMatch.Column5Value}")
                Console.WriteLine($"- Layout: {result.FirstMatch.Column6Value}")
            End If
        Else
            Console.WriteLine(result.SummaryMessage)
            If Not String.IsNullOrEmpty(result.ErrorMessage) Then
                Console.WriteLine($"ข้อผิดพลาด: {result.ErrorMessage}")
            End If
        End If
    End Sub

End Class 