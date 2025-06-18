Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' ฟอร์มสำหรับทดสอบการเชื่อมต่อ Network
''' </summary>
Public Class NetworkTestForm
    Inherits Form
    
    Private lblStatus As Label
    Private btnTest As Button
    Private txtResults As TextBox
    Private btnClose As Button
    
    Public Sub New()
        InitializeComponent()
    End Sub
    
    Private Sub InitializeComponent()
        Me.Text = "ทดสอบการเชื่อมต่อ Network"
        Me.Size = New Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterScreen
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        
        ' สร้าง Controls
        lblStatus = New Label()
        lblStatus.Text = "กด 'ทดสอบ' เพื่อตรวจสอบการเชื่อมต่อเครือข่าย"
        lblStatus.Location = New Point(20, 20)
        lblStatus.Size = New Size(550, 30)
        lblStatus.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        
        btnTest = New Button()
        btnTest.Text = "ทดสอบการเชื่อมต่อ"
        btnTest.Location = New Point(20, 60)
        btnTest.Size = New Size(150, 35)
        btnTest.BackColor = Color.FromArgb(0, 123, 255)
        btnTest.ForeColor = Color.White
        btnTest.FlatStyle = FlatStyle.Flat
        AddHandler btnTest.Click, AddressOf BtnTest_Click
        
        txtResults = New TextBox()
        txtResults.Location = New Point(20, 110)
        txtResults.Size = New Size(540, 300)
        txtResults.Multiline = True
        txtResults.ScrollBars = ScrollBars.Vertical
        txtResults.ReadOnly = True
        txtResults.Font = New Font("Consolas", 9)
        txtResults.BackColor = Color.FromArgb(248, 249, 250)
        
        btnClose = New Button()
        btnClose.Text = "ปิด"
        btnClose.Location = New Point(485, 420)
        btnClose.Size = New Size(75, 30)
        btnClose.BackColor = Color.FromArgb(108, 117, 125)
        btnClose.ForeColor = Color.White
        btnClose.FlatStyle = FlatStyle.Flat
        AddHandler btnClose.Click, AddressOf BtnClose_Click
        
        ' เพิ่ม Controls เข้าฟอร์ม
        Me.Controls.AddRange({lblStatus, btnTest, txtResults, btnClose})
    End Sub
    
    Private Sub BtnTest_Click(sender As Object, e As EventArgs)
        btnTest.Enabled = False
        btnTest.Text = "กำลังทดสอบ..."
        txtResults.Clear()
        
        Try
            txtResults.AppendText("=== ทดสอบการเชื่อมต่อ Network ===" & vbNewLine)
            txtResults.AppendText($"เวลา: {DateTime.Now:yyyy-MM-dd HH:mm:ss}" & vbNewLine)
            txtResults.AppendText("" & vbNewLine)
            
            txtResults.AppendText("Logic การตรวจสอบ:" & vbNewLine)
            txtResults.AppendText("• ถ้าปิง 172.24.0.3 ไม่ได้ = เครือข่าย OA" & vbNewLine)
            txtResults.AppendText("• ถ้าปิงได้ทั้ง 10.24.179.2 และ 172.24.0.3 = เครือข่าย FAB" & vbNewLine)
            txtResults.AppendText("• ถ้าปิงได้แค่ 172.24.0.3 = เครือข่าย FAB" & vbNewLine)
            txtResults.AppendText("" & vbNewLine)
            
            ' ทดสอบการเชื่อมต่อ
            txtResults.AppendText("กำลังทดสอบการเชื่อมต่อ..." & vbNewLine)
            Application.DoEvents()
            
            Dim result = NetworkPathManager.CheckNetworkConnection()
            
            txtResults.AppendText("" & vbNewLine)
            txtResults.AppendText("=== ผลการทดสอบ ===" & vbNewLine)
            
            If result.IsConnected Then
                txtResults.AppendText($"✅ เชื่อมต่อสำเร็จ!" & vbNewLine)
                txtResults.AppendText($"ประเภทเครือข่าย: {result.NetworkType}" & vbNewLine)
                txtResults.AppendText($"เซิร์ฟเวอร์: {result.ServerIP}" & vbNewLine)
                txtResults.AppendText($"Base Path: {result.BasePath}" & vbNewLine)
                
                lblStatus.Text = $"เชื่อมต่อกับเครือข่าย {result.NetworkType} สำเร็จ! 🎉"
                lblStatus.ForeColor = Color.Green
            Else
                txtResults.AppendText($"❌ ไม่สามารถเชื่อมต่อได้" & vbNewLine)
                txtResults.AppendText($"ข้อผิดพลาด: {result.ErrorMessage}" & vbNewLine)
                
                lblStatus.Text = "ไม่สามารถเชื่อมต่อเครือข่ายได้ ❌"
                lblStatus.ForeColor = Color.Red
            End If
            
            txtResults.AppendText("" & vbNewLine)
            txtResults.AppendText("=== ทดสอบ Path ต่างๆ ===" & vbNewLine)
            
            ' ทดสอบ path ต่างๆ
            Dim paths As New Dictionary(Of String, String) From {
                {"Excel Database", NetworkPathManager.GetExcelDatabasePath()},
                {"Access Database", NetworkPathManager.GetAccessDatabasePath()},
                {"Update System", NetworkPathManager.GetUpdateSystemPath()},
                {"Film Character Check", NetworkPathManager.GetFilmCharacterCheckPath()},
                {"Drawing Folder", NetworkPathManager.GetDrawingFolderPath()}
            }
            
            For Each kvp In paths
                If Not String.IsNullOrEmpty(kvp.Value) Then
                    Dim exists = NetworkPathManager.PathExists(kvp.Value)
                    Dim statusIcon = If(exists, "✅", "⚠️")
                    txtResults.AppendText($"{statusIcon} {kvp.Key}: {kvp.Value}" & vbNewLine)
                Else
                    txtResults.AppendText($"❌ {kvp.Key}: ไม่มี path" & vbNewLine)
                End If
            Next
            
            txtResults.AppendText("" & vbNewLine)
            txtResults.AppendText("=== Network Status ===" & vbNewLine)
            txtResults.AppendText(NetworkPathManager.GetNetworkStatus() & vbNewLine)
            
        Catch ex As Exception
            txtResults.AppendText($"เกิดข้อผิดพลาด: {ex.Message}" & vbNewLine)
            lblStatus.Text = "เกิดข้อผิดพลาดในการทดสอบ"
            lblStatus.ForeColor = Color.Red
        Finally
            btnTest.Enabled = True
            btnTest.Text = "ทดสอบการเชื่อมต่อ"
        End Try
    End Sub
    
    Private Sub BtnClose_Click(sender As Object, e As EventArgs)
        Me.Close()
    End Sub
    
End Class 