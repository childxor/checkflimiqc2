''' <summary>
''' คลาสสำหรับรายละเอียดการค้นหาใน Excel
''' </summary>
Public Class ExcelMatchDetail
    ''' <summary>
    ''' หมายเลขแถวในไฟล์ Excel
    ''' </summary>
    Public Property RowNumber As Integer = 0

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 1
    ''' </summary>
    Public Property Column1Value As String = ""

    ''' <summary>
    ''' ค่าในคอลัมน์ที่ 4
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
End Class 