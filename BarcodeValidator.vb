Imports System.Text.RegularExpressions
Imports System.IO

Public Class BarcodeValidator

#Region "Barcode Validation Constants"
    ' กำหนดความยาวขั้นต่ำของ barcode
    Private Const MIN_BARCODE_LENGTH As Integer = 50
    
    ' กำหนด pattern หลักสำหรับการตรวจสอบ
    Private Const BASE_PATTERN As String = "R\d+C-\d+.*\+Q\d+.*\+P.*\+D\d+.*"
    
    ' กำหนด patterns สำหรับการ extract ข้อมูลแต่ละส่วน
    Private Const PART_R_PATTERN As String = "R(\d+C-\d+)"
    Private Const PART_Q_PATTERN As String = "\+Q(\d+)"
    Private Const PART_P_PATTERN As String = "\+P([^+]+)"
    Private Const PART_D_PATTERN As String = "\+D(\d+)"
    Private Const PART_L_PATTERN As String = "\+L([^+]*)"
    Private Const PART_V_PATTERN As String = "\+V([^+]*)"
    Private Const PART_U_PATTERN As String = "\+U(\d+)"
#End Region

#Region "Barcode Validation Methods"
    ''' <summary>
    ''' ตรวจสอบความถูกต้องของ barcode data
    ''' </summary>
    Public Shared Function ValidateBarcode(barcodeData As String) As BarcodeValidationResult
        Dim result As New BarcodeValidationResult()
        result.OriginalData = barcodeData
        result.IsValid = True
        result.ValidationMessages = New List(Of String)()

        ' ตรวจสอบความยาวขั้นต่ำ
        If String.IsNullOrEmpty(barcodeData) Then
            result.IsValid = False
            result.ValidationMessages.Add("ข้อมูล barcode ว่างเปล่า")
            Return result
        End If

        If barcodeData.Length < MIN_BARCODE_LENGTH Then
            result.IsValid = False
            result.ValidationMessages.Add($"ข้อมูล barcode สั้นเกินไป (ความยาว: {barcodeData.Length}, ต้องการอย่างน้อย: {MIN_BARCODE_LENGTH})")
        End If

        ' ตรวจสอบ pattern หลัก
        If Not Regex.IsMatch(barcodeData, BASE_PATTERN) Then
            result.IsValid = False
            result.ValidationMessages.Add("รูปแบบ barcode ไม่ถูกต้อง (ไม่พบ pattern หลัก)")
        End If

        ' ตรวจสอบแต่ละส่วนย่อย
        ValidateRequiredParts(barcodeData, result)

        Return result
    End Function

    ''' <summary>
    ''' ตรวจสอบส่วนประกอบที่จำเป็นของ barcode
    ''' </summary>
    Private Shared Sub ValidateRequiredParts(barcodeData As String, result As BarcodeValidationResult)
        ' ตรวจสอบส่วน R (Required)
        If Not Regex.IsMatch(barcodeData, PART_R_PATTERN) Then
            result.IsValid = False
            result.ValidationMessages.Add("ไม่พบข้อมูลส่วน R (รหัสอ้างอิง)")
        End If

        ' ตรวจสอบส่วน Q (Required)
        If Not Regex.IsMatch(barcodeData, PART_Q_PATTERN) Then
            result.IsValid = False
            result.ValidationMessages.Add("ไม่พบข้อมูลส่วน Q (จำนวน)")
        End If

        ' ตรวจสอบส่วน P (Required)
        If Not Regex.IsMatch(barcodeData, PART_P_PATTERN) Then
            result.IsValid = False
            result.ValidationMessages.Add("ไม่พบข้อมูลส่วน P (รหัสผลิตภัณฑ์)")
        End If

        ' ตรวจสอบส่วน D (Required)
        If Not Regex.IsMatch(barcodeData, PART_D_PATTERN) Then
            result.IsValid = False
            result.ValidationMessages.Add("ไม่พบข้อมูลส่วน D (วันที่)")
        End If

        ' ส่วน L, V, U เป็น optional
        If Not Regex.IsMatch(barcodeData, PART_L_PATTERN) Then
            result.ValidationMessages.Add("ไม่พบข้อมูลส่วน L (ข้อมูลเพิ่มเติม) - อาจไม่จำเป็น")
        End If
    End Sub
#End Region

#Region "Data Extraction Methods"
    ''' <summary>
    ''' ดึงข้อมูลจาก barcode แบบยืดหยุ่น
    ''' </summary>
    Public Shared Function ExtractBarcodeData(barcodeData As String, extractionMode As ExtractionMode) As BarcodeExtractedData
        Dim extractedData As New BarcodeExtractedData()
        extractedData.OriginalData = barcodeData

        Try
            Select Case extractionMode
                Case ExtractionMode.ProductCodeOnly
                    extractedData.ExtractedValue = ExtractProductCode(barcodeData)
                Case ExtractionMode.AllParts
                    extractedData = ExtractAllParts(barcodeData)
                Case ExtractionMode.CustomPattern
                    extractedData.ExtractedValue = ExtractWithCustomPattern(barcodeData)
                Case ExtractionMode.Intelligent
                    extractedData = IntelligentExtraction(barcodeData)
            End Select

        Catch ex As Exception
            extractedData.ExtractedValue = barcodeData
            extractedData.ErrorMessage = $"เกิดข้อผิดพลาดในการดึงข้อมูล: {ex.Message}"
        End Try

        Return extractedData
    End Function

    ''' <summary>
    ''' ดึงรหัสผลิตภัณฑ์เท่านั้น
    ''' </summary>
    Private Shared Function ExtractProductCode(barcodeData As String) As String
        Dim match As Match = Regex.Match(barcodeData, PART_P_PATTERN)
        If match.Success AndAlso match.Groups.Count > 1 Then
            Return match.Groups(1).Value
        End If
        Return barcodeData
    End Function

    ''' <summary>
    ''' ดึงข้อมูลทุกส่วน
    ''' </summary>
    Private Shared Function ExtractAllParts(barcodeData As String) As BarcodeExtractedData
        Dim result As New BarcodeExtractedData()
        result.OriginalData = barcodeData

        ' ดึงข้อมูลแต่ละส่วน
        result.ReferenceCode = ExtractPart(barcodeData, PART_R_PATTERN)
        result.Quantity = ExtractPart(barcodeData, PART_Q_PATTERN)
        result.ProductCode = ExtractPart(barcodeData, PART_P_PATTERN)
        result.DateCode = ExtractPart(barcodeData, PART_D_PATTERN)
        result.LocationCode = ExtractPart(barcodeData, PART_L_PATTERN)
        result.VersionCode = ExtractPart(barcodeData, PART_V_PATTERN) 'test git
        result.UserCode = ExtractPart(barcodeData, PART_U_PATTERN)

        ' สร้างข้อมูลที่ดึงออกมาในรูปแบบที่อ่านง่าย
        result.ExtractedValue = $"รหัสผลิตภัณฑ์: {result.ProductCode} " & vbNewLine &
                               $"รหัสอ้างอิง: {result.ReferenceCode} " & vbNewLine &
                               $"จำนวน: {result.Quantity} " & vbNewLine &
                               $"วันที่: {result.DateCode} "

        Return result
    End Function

    ''' <summary>
    ''' ดึงข้อมูลด้วย custom pattern
    ''' </summary>
    Private Shared Function ExtractWithCustomPattern(barcodeData As String) As String
        ' ใช้ pattern จากการตั้งค่า (ถ้ามี)
        Dim customPattern As String = GetCustomPattern()
        If Not String.IsNullOrEmpty(customPattern) Then
            Dim match As Match = Regex.Match(barcodeData, customPattern)
            If match.Success AndAlso match.Groups.Count > 1 Then
                Return match.Groups(1).Value
            End If
        End If
        Return ExtractProductCode(barcodeData)
    End Function

    ''' <summary>
    ''' การดึงข้อมูลแบบอัจฉริยะ - เลือกวิธีที่เหมาะสมอัตโนมัติ
    ''' </summary>
    Private Shared Function IntelligentExtraction(barcodeData As String) As BarcodeExtractedData
        ' ตรวจสอบความถูกต้องก่อน
        Dim validation As BarcodeValidationResult = ValidateBarcode(barcodeData)
        
        If validation.IsValid Then
            ' ถ้าข้อมูลถูกต้อง ดึงทุกส่วน
            Return ExtractAllParts(barcodeData)
        Else
            ' ถ้าข้อมูลไม่สมบูรณ์ พยายามดึงส่วนที่มี
            Dim result As New BarcodeExtractedData()
            result.OriginalData = barcodeData
            result.ProductCode = ExtractProductCode(barcodeData)
            result.ExtractedValue = result.ProductCode
            result.ErrorMessage = String.Join("; ", validation.ValidationMessages)
            Return result
        End If
    End Function

    ''' <summary>
    ''' ดึงข้อมูลส่วนใดส่วนหนึ่ง
    ''' </summary>
    Private Shared Function ExtractPart(barcodeData As String, pattern As String) As String
        Try
            Dim match As Match = Regex.Match(barcodeData, pattern)
            If match.Success AndAlso match.Groups.Count > 1 Then
                Return match.Groups(1).Value
            End If
        Catch
            ' ไม่ต้องทำอะไร
        End Try
        Return ""
    End Function

    ''' <summary>
    ''' ดึง custom pattern จากการตั้งค่า
    ''' </summary>
    Private Shared Function GetCustomPattern() As String
        Try
            ' อ่านจากไฟล์ config หรือ registry
            If File.Exists("Settings.config") Then
                ' โค้ดสำหรับอ่าน pattern จาก config
                Return "\+P([^+]+)\+D" ' ค่าเริ่มต้น
            End If
        Catch
            ' ไม่ต้องทำอะไร
        End Try
        Return "\+P([^+]+)\+D"
    End Function
#End Region

End Class

#Region "Supporting Classes and Enums"
''' <summary>
''' โหมดการดึงข้อมูล
''' </summary>
Public Enum ExtractionMode
    ProductCodeOnly     ' ดึงรหัสผลิตภัณฑ์เท่านั้น
    AllParts           ' ดึงข้อมูลทุกส่วน
    CustomPattern      ' ดึงตาม pattern ที่กำหนด
    Intelligent        ' เลือกวิธีที่เหมาะสมอัตโนมัติ
End Enum

''' <summary>
''' ผลการตรวจสอบ barcode
''' </summary>
Public Class BarcodeValidationResult
    Public Property IsValid As Boolean
    Public Property OriginalData As String
    Public Property ValidationMessages As List(Of String)
    
    Public ReadOnly Property IsPartiallyValid As Boolean
        Get
            Return ValidationMessages IsNot Nothing AndAlso 
                   ValidationMessages.Count > 0 AndAlso 
                   ValidationMessages.Count <= 2
        End Get
    End Property
End Class

''' <summary>
''' ข้อมูลที่ดึงออกมาจาก barcode
''' </summary>
Public Class BarcodeExtractedData
    Public Property OriginalData As String
    Public Property ExtractedValue As String
    Public Property ReferenceCode As String
    Public Property Quantity As String
    Public Property ProductCode As String
    Public Property DateCode As String
    Public Property LocationCode As String
    Public Property VersionCode As String
    Public Property UserCode As String
    Public Property ErrorMessage As String
    
    Public ReadOnly Property HasError As Boolean
        Get
            Return Not String.IsNullOrEmpty(ErrorMessage)
        End Get
    End Property
End Class
#End Region