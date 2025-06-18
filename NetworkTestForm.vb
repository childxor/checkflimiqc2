Imports System.Windows.Forms
Imports System.Drawing

''' <summary>
''' ฟอร์มสำหรับทดสอบการเชื่อมต่อ network และแสดงพาธต่างๆ
''' </summary>
Public Class NetworkTestForm
    Inherits Form
    
    Private WithEvents btnTestConnection As Button
    Private WithEvents btnRefresh As Button
    Private WithEvents txtStatus As TextBox
    Private WithEvents lblTitle As Label
    
    Public Sub New()
        InitializeComponent()
    End Sub
    
    Private Sub InitializeComponent()
        ' ตั้งค่าฟอร์ม
        Me.Text = "ทดสอบการเชื่อมต่อ Network"
        Me.Size = New Size(600, 500)
        Me.StartPosition = FormStartPosition.CenterParent
        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.BackColor = Color.White
        
        ' หัวข้อ
        lblTitle = New Label()
        lblTitle.Text = "🌐 ทดสอบการเชื่อมต่อ Network OA/FAB"
        lblTitle.Font = New Font("Segoe UI", 14, FontStyle.Bold)
        lblTitle.Location = New Point(20, 20)
        lblTitle.Size = New Size(550, 30)
        lblTitle.ForeColor = Color.FromArgb(52, 73, 94)
        lblTitle.TextAlign = ContentAlignment.MiddleCenter
        
        ' ปุ่มทดสอบการเชื่อมต่อ
        btnTestConnection = New Button()
        btnTestConnection.Text = "🔍 ทดสอบการเชื่อมต่อ"
        btnTestConnection.Size = New Size(150, 40)
        btnTestConnection.Location = New Point(150, 70)
        btnTestConnection.BackColor = Color.FromArgb(52, 152, 219)
        btnTestConnection.ForeColor = Color.White
        btnTestConnection.FlatStyle = FlatStyle.Flat
        btnTestConnection.FlatAppearance.BorderSize = 0
        btnTestConnection.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        
        ' ปุ่มรีเฟรช
        btnRefresh = New Button()
        btnRefresh.Text = "🔄 รีเฟรช"
        btnRefresh.Size = New Size(100, 40)
        btnRefresh.Location = New Point(320, 70)
        btnRefresh.BackColor = Color.FromArgb(39, 174, 96)
        btnRefresh.ForeColor = Color.White
        btnRefresh.FlatStyle = FlatStyle.Flat
        btnRefresh.FlatAppearance.BorderSize = 0
        btnRefresh.Font = New Font("Segoe UI", 10, FontStyle.Bold)
        
        ' กล่องข้อความแสดงสถานะ
        txtStatus = New TextBox()
        txtStatus.Multiline = True
        txtStatus.ReadOnly = True
        txtStatus.ScrollBars = ScrollBars.Vertical
        txtStatus.Location = New Point(20, 130)
        txtStatus.Size = New Size(540, 300)
        txtStatus.Font = New Font("Consolas", 10)
        txtStatus.BackColor = Color.FromArgb(248, 249, 250)
        txtStatus.BorderStyle = BorderStyle.FixedSingle
        
        ' เพิ่ม controls เข้าฟอร์ม
        Me.Controls.AddRange({lblTitle, btnTestConnection, btnRefresh, txtStatus})
        
        ' ทดสอบครั้งแรกเมื่อเปิดฟอร์ม
        TestNetworkConnection()
    End Sub
    
    Private Sub btnTestConnection_Click(sender As Object, e As EventArgs) Handles btnTestConnection.Click
        TestNetworkConnection()
    End Sub
    
    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        TestNetworkConnection()
    End Sub
    
    Private Sub TestNetworkConnection()
        Try
            txtStatus.Clear()
            txtStatus.AppendText("🔍 กำลังทดสอบการเชื่อมต่อ network..." & vbNewLine & vbNewLine)
            
            Application.DoEvents()
            
            ' ทดสอบการเชื่อมต่อ
            Dim networkResult = NetworkPathManager.CheckNetworkConnection()
            
            If networkResult.IsConnected Then
                txtStatus.AppendText($"✅ เชื่อมต่อสำเร็จ!" & vbNewLine)
                txtStatus.AppendText($"🌐 Network Type: {networkResult.NetworkType}" & vbNewLine)
                txtStatus.AppendText($"🖥️ Server IP: {networkResult.ServerIP}" & vbNewLine)
                txtStatus.AppendText($"📁 Base Path: {networkResult.BasePath}" & vbNewLine)
                txtStatus.AppendText(vbNewLine & "📂 รายการพาธที่พร้อมใช้งาน:" & vbNewLine)
                txtStatus.AppendText("=" & New String("="c, 50) & vbNewLine)
                
                ' ทดสอบพาธต่างๆ
                TestPath("Excel Database", NetworkPathManager.GetExcelDatabasePath())
                TestPath("Access Database", NetworkPathManager.GetAccessDatabasePath())
                TestPath("Update System", NetworkPathManager.GetUpdateSystemPath())
                TestPath("Film Character Check", NetworkPathManager.GetFilmCharacterCheckPath())
                TestPath("Drawing Folder", NetworkPathManager.GetDrawingFolderPath())
                
                ' ทดสอบพาธกำหนดเอง
                txtStatus.AppendText(vbNewLine & "🔧 ทดสอบพาธกำหนดเอง:" & vbNewLine)
                TestPath("Drawing Folder (Custom)", NetworkPathManager.GetCustomPath("Film charecter check\Drawing"))
                TestPath("Debug Systems (Custom)", NetworkPathManager.GetCustomPath("Film charecter check\DebugSystems"))
                
            Else
                txtStatus.AppendText($"❌ ไม่สามารถเชื่อมต่อได้" & vbNewLine)
                txtStatus.AppendText($"🔴 ข้อผิดพลาด: {networkResult.ErrorMessage}" & vbNewLine)
                txtStatus.AppendText(vbNewLine & "💡 คำแนะนำ:" & vbNewLine)
                txtStatus.AppendText("• ตรวจสอบการเชื่อมต่อเครือข่าย" & vbNewLine)
                txtStatus.AppendText("• ตรวจสอบ IP Address เซิร์ฟเวอร์" & vbNewLine)
                txtStatus.AppendText("• ตรวจสอบสิทธิ์การเข้าถึง Network Share" & vbNewLine)
            End If
            
            txtStatus.AppendText(vbNewLine & "⏰ ทดสอบเสร็จสิ้น: " & DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"))
            
        Catch ex As Exception
            txtStatus.AppendText($"💥 เกิดข้อผิดพลาดในการทดสอบ: {ex.Message}" & vbNewLine)
        End Try
    End Sub
    
    Private Sub TestPath(description As String, path As String)
        Try
            If String.IsNullOrEmpty(path) Then
                txtStatus.AppendText($"❌ {description}: ไม่ได้รับพาธ" & vbNewLine)
                Return
            End If
            
            Dim exists = NetworkPathManager.PathExists(path)
            Dim status = If(exists, "✅ พบ", "⚠️ ไม่พบ")
            
            txtStatus.AppendText($"{status} {description}:" & vbNewLine)
            txtStatus.AppendText($"   📍 {path}" & vbNewLine)
            
        Catch ex As Exception
            txtStatus.AppendText($"❌ {description}: ข้อผิดพลาด - {ex.Message}" & vbNewLine)
        End Try
    End Sub
    
End Class 