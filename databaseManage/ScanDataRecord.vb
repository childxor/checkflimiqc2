Imports System
Imports System.IO

''' <summary>
''' คลาสสำหรับเก็บข้อมูลการสแกน QR Code
''' </summary>
Public Class ScanDataRecord
    Public Property Id As Integer = 0
    Public Property ScanDateTime As DateTime = DateTime.Now
    Public Property ProductCode As String = ""
    Public Property ReferenceCode As String = ""
    Public Property Quantity As Integer = 0
    Public Property DateCode As String = ""
    Public Property IsValid As Boolean = False
    Public Property OriginalData As String = ""
    Public Property ExtractedData As String = ""
    Public Property ValidationMessages As String = ""
    Public Property ComputerName As String = ""
    Public Property UserName As String = ""
    Public Property MissionStatus As String = "ไม่มี"
    
    ''' <summary>
    ''' เส้นทางไฟล์ที่เกี่ยวข้องกับข้อมูลนี้
    ''' </summary>
    Public Property RelatedFilePath As String = ""

    Public Sub New()
        ScanDateTime = DateTime.Now
        ComputerName = Environment.MachineName
        UserName = Environment.UserName
    End Sub

    ''' <summary>
    ''' ตรวจสอบว่ามีไฟล์ที่เกี่ยวข้องหรือไม่
    ''' </summary>
    Public ReadOnly Property HasRelatedFile As Boolean
        Get
            Return Not String.IsNullOrEmpty(RelatedFilePath) AndAlso File.Exists(RelatedFilePath)
        End Get
    End Property

    ''' <summary>
    ''' ชื่อไฟล์ที่เกี่ยวข้อง (ไม่รวม path)
    ''' </summary>
    Public ReadOnly Property RelatedFileName As String
        Get
            If String.IsNullOrEmpty(RelatedFilePath) Then
                Return ""
            End If
            Return Path.GetFileName(RelatedFilePath)
        End Get
    End Property

    Public Overrides Function ToString() As String
        Return $"ID: {Id}, ProductCode: {ProductCode}, DateTime: {ScanDateTime:yyyy-MM-dd HH:mm:ss}, Status: {MissionStatus}"
    End Function
End Class 