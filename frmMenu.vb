Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.IO

Public Class frmMenu
    Private Const WH_KEYBOARD_LL As Integer = 13
    Private Delegate Function LowLevelKeyboardProc(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr

    Private _proc As LowLevelKeyboardProc = AddressOf HookCallback
    Private _hookID As IntPtr = IntPtr.Zero
    Private barcodeBuffer As String = ""
    Private lastKeyTime As DateTime = DateTime.Now

    ' Settings variables
    Private scanTimeout As Integer = 100
    Private showFullData As Boolean = False
    Private autoExtract As Boolean = True
    Private soundEnabled As Boolean = True
    Private extractPattern As String = "\+P([^+]+)\+D"

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function SetWindowsHookEx(idHook As Integer, lpfn As LowLevelKeyboardProc, hMod As IntPtr, dwThreadId As UInteger) As IntPtr
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function UnhookWindowsHookEx(hhk As IntPtr) As <MarshalAs(UnmanagedType.Bool)> Boolean
    End Function

    <DllImport("user32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function CallNextHookEx(hhk As IntPtr, nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
    End Function

    <DllImport("kernel32.dll", CharSet:=CharSet.Auto, SetLastError:=True)>
    Private Shared Function GetModuleHandle(lpModuleName As String) As IntPtr
    End Function

    ' เล่นเสียงแจ้งเตือน
    <DllImport("winmm.dll", SetLastError:=True)>
    Private Shared Function PlaySound(pszSound As String, hmod As IntPtr, fdwSound As UInteger) As Boolean
    End Function

    Private Const SND_FILENAME As UInteger = &H20000
    Private Const SND_ASYNC As UInteger = &H1

    Private Sub frmMenu_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        _hookID = SetHook(_proc)
        LoadSettingsFromConfig()
        InitializeUI()
        ApplyThemeSettings()

        ' แสดงสถานะเริ่มต้น
        UpdateStatusBar("พร้อมรับการสแกน QR Code")
    End Sub

    Private Sub InitializeUI()
        Try
            ' สร้าง MenuStrip หากยังไม่มี
            If Me.MainMenuStrip Is Nothing Then
                Dim menuStrip As New MenuStrip()
                Me.MainMenuStrip = menuStrip
                Me.Controls.Add(menuStrip)

                ' สร้างเมนู
                Dim fileMenu As New ToolStripMenuItem("ไฟล์")
                Dim settingsMenu As New ToolStripMenuItem("การตั้งค่า", Nothing, AddressOf OpenSettings)
                Dim exitMenu As New ToolStripMenuItem("ออก", Nothing, AddressOf ExitApplication)

                fileMenu.DropDownItems.AddRange({settingsMenu, New ToolStripSeparator(), exitMenu})
                menuStrip.Items.Add(fileMenu)

                Dim helpMenu As New ToolStripMenuItem("ช่วยเหลือ")
                Dim aboutMenu As New ToolStripMenuItem("เกี่ยวกับ", Nothing, AddressOf ShowAbout)
                helpMenu.DropDownItems.Add(aboutMenu)
                menuStrip.Items.Add(helpMenu)
            End If

            ' สร้าง StatusStrip หากยังไม่มี
            If Me.Controls.OfType(Of StatusStrip)().Count = 0 Then
                Dim statusStrip As New StatusStrip()
                Dim statusLabel As New ToolStripStatusLabel("พร้อมใช้งาน")
                statusLabel.Name = "statusLabel"
                statusStrip.Items.Add(statusLabel)
                Me.Controls.Add(statusStrip)
            End If

            ' สร้าง TextBox สำหรับแสดงบาร์โค้ด หากยังไม่มี
            If Me.Controls.Find("txtBarcode", True).Length = 0 Then
                Dim txtBarcode As New TextBox()
                txtBarcode.Name = "txtBarcode"
                txtBarcode.Location = New Point(20, 50)
                txtBarcode.Size = New Size(400, 30)
                txtBarcode.Font = New Font("Segoe UI", 12)
                txtBarcode.ReadOnly = True

                Dim lblBarcode As New Label()
                lblBarcode.Text = "ข้อมูลที่สแกนได้:"
                lblBarcode.Location = New Point(20, 25)
                lblBarcode.Size = New Size(200, 20)

                Me.Controls.Add(lblBarcode)
                Me.Controls.Add(txtBarcode)
            End If

            ' สร้างปุ่มต่างๆ
            CreateButtons()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการสร้างหน้าจอ: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CreateButtons()
        Try
            ' ปุ่มการตั้งค่า
            If Me.Controls.Find("btnSettings", True).Length = 0 Then
                Dim btnSettings As New Button()
                btnSettings.Name = "btnSettings"
                btnSettings.Text = "การตั้งค่า"
                btnSettings.Location = New Point(450, 50)
                btnSettings.Size = New Size(100, 30)
                AddHandler btnSettings.Click, AddressOf OpenSettings
                Me.Controls.Add(btnSettings)
            End If

            ' ปุ่มทดสอบ
            If Me.Controls.Find("btnTest", True).Length = 0 Then
                Dim btnTest As New Button()
                btnTest.Name = "btnTest"
                btnTest.Text = "ทดสอบ"
                btnTest.Location = New Point(560, 50)
                btnTest.Size = New Size(100, 30)
                AddHandler btnTest.Click, AddressOf TestScan
                Me.Controls.Add(btnTest)
            End If

            ' ปุ่มล้างข้อมูล
            If Me.Controls.Find("btnClear", True).Length = 0 Then
                Dim btnClear As New Button()
                btnClear.Name = "btnClear"
                btnClear.Text = "ล้างข้อมูล"
                btnClear.Location = New Point(670, 50)
                btnClear.Size = New Size(100, 30)
                AddHandler btnClear.Click, AddressOf ClearData
                Me.Controls.Add(btnClear)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการสร้างปุ่ม: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadSettingsFromConfig()
        Try
            If File.Exists("Settings.config") Then
                ' โหลดการตั้งค่าจากไฟล์
                Dim tempSettings As New frmSettings()
                scanTimeout = CInt(tempSettings.GetSetting("scantimeout"))
                showFullData = CBool(tempSettings.GetSetting("showfulldata"))
                autoExtract = CBool(tempSettings.GetSetting("autoextract"))
                soundEnabled = CBool(tempSettings.GetSetting("soundenabled"))
                extractPattern = CStr(tempSettings.GetSetting("extractpattern"))
                tempSettings.Dispose()
            End If
        Catch ex As Exception
            ' ใช้ค่าเริ่มต้นหากโหลดไม่สำเร็จ
            scanTimeout = 100
            showFullData = False
            autoExtract = True
            soundEnabled = True
            extractPattern = "\+P([^+]+)\+D"
        End Try
    End Sub

    Private Sub ApplyThemeSettings()
        Try
            If File.Exists("Settings.config") Then
                Dim tempSettings As New frmSettings()
                Dim theme As String = CStr(tempSettings.GetSetting("theme"))
                Dim showStatusBar As Boolean = CBool(tempSettings.GetSetting("showstatusbar"))
                Dim showToolbar As Boolean = CBool(tempSettings.GetSetting("showtoolbar"))

                ' ปรับธีม
                Select Case theme.ToLower()
                    Case "dark"
                        ApplyDarkTheme()
                    Case "light"
                        ApplyLightTheme()
                End Select

                ' แสดง/ซ่อน UI elements
                If Me.MainMenuStrip IsNot Nothing Then
                    Me.MainMenuStrip.Visible = showToolbar
                End If

                Dim statusStrip = Me.Controls.OfType(Of StatusStrip)().FirstOrDefault()
                If statusStrip IsNot Nothing Then
                    statusStrip.Visible = showStatusBar
                End If

                tempSettings.Dispose()
            End If
        Catch
            ' ใช้ธีม default หากมีปัญหา
            ApplyLightTheme()
        End Try
    End Sub

    Private Sub ApplyDarkTheme()
        Me.BackColor = Color.FromArgb(45, 45, 48)
        Me.ForeColor = Color.White

        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.BackColor = Color.FromArgb(37, 37, 38)
                ctrl.ForeColor = Color.White
            ElseIf TypeOf ctrl Is Button Then
                ctrl.BackColor = Color.FromArgb(62, 62, 64)
                ctrl.ForeColor = Color.White
            End If
        Next
    End Sub

    Private Sub ApplyLightTheme()
        Me.BackColor = SystemColors.Control
        Me.ForeColor = SystemColors.ControlText

        For Each ctrl As Control In Me.Controls
            If TypeOf ctrl Is TextBox Then
                ctrl.BackColor = SystemColors.Window
                ctrl.ForeColor = SystemColors.WindowText
            ElseIf TypeOf ctrl Is Button Then
                ctrl.BackColor = SystemColors.Control
                ctrl.ForeColor = SystemColors.ControlText
            End If
        Next
    End Sub

    Private Function SetHook(proc As LowLevelKeyboardProc) As IntPtr
        Using curProcess As Process = Process.GetCurrentProcess()
            Using curModule As ProcessModule = curProcess.MainModule
                Return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0)
            End Using
        End Using
    End Function

    Private Function HookCallback(nCode As Integer, wParam As IntPtr, lParam As IntPtr) As IntPtr
        If nCode >= 0 AndAlso wParam = CType(256, IntPtr) Then ' WM_KEYDOWN
            Dim vkCode As Integer = Marshal.ReadInt32(lParam)
            Dim currentTime As DateTime = DateTime.Now

            ' ใช้ค่า timeout จากการตั้งค่า
            If (currentTime - lastKeyTime).TotalMilliseconds > scanTimeout Then
                barcodeBuffer = ""
            End If

            lastKeyTime = currentTime

            ' แปลง virtual key code เป็นตัวอักษร
            If vkCode >= 48 AndAlso vkCode <= 57 Then ' 0-9
                barcodeBuffer += Chr(vkCode)
            ElseIf vkCode >= 65 AndAlso vkCode <= 90 Then ' A-Z
                barcodeBuffer += Chr(vkCode)
            ElseIf vkCode = 189 Then ' Minus key (-)
                barcodeBuffer += "-"
            ElseIf vkCode = 187 AndAlso Control.ModifierKeys = Keys.Shift Then ' Plus key (+)
                barcodeBuffer += "+"
            ElseIf vkCode = 13 Then ' Enter key (จบการสแกน)
                If barcodeBuffer.Length > 0 Then
                    ProcessBarcode(barcodeBuffer)
                    barcodeBuffer = ""
                End If
            End If
        End If

        Return CallNextHookEx(_hookID, nCode, wParam, lParam)
    End Function

    Private Sub ProcessBarcode(barcode As String)
        ' ดึงข้อมูลส่วนที่ต้องการจาก QR code
        Dim extractedData As String = barcode

        If autoExtract Then
            extractedData = ExtractProductCode(barcode)
        End If

        ' แสดงผลในกล่องข้อความ
        Me.Invoke(Sub()
                      Dim txtBarcode = CType(Me.Controls.Find("txtBarcode", True).FirstOrDefault(), TextBox)
                      If txtBarcode IsNot Nothing Then
                          txtBarcode.Text = extractedData
                      End If

                      ' เล่นเสียงหากเปิดใช้งาน
                      If soundEnabled Then
                          PlaySystemSound()
                      End If

                      ' แสดงข้อความแจ้งเตือน
                      If showFullData Then
                          MessageBox.Show($"QR Code ทั้งหมด: {barcode}{vbNewLine}{vbNewLine}ข้อมูลที่ดึงออกมา: {extractedData}",
                                        "ผลการแสกน QR Code",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information)
                      Else
                          UpdateStatusBar($"สแกนสำเร็จ: {extractedData}")
                      End If

                      ' บันทึก log หากเปิดใช้งาน
                      WriteLog($"Scanned: {barcode} -> Extracted: {extractedData}")
                  End Sub)
    End Sub

    Private Function ExtractProductCode(qrData As String) As String
        Try
            If String.IsNullOrEmpty(extractPattern) Then
                Return qrData
            End If

            ' ใช้ pattern จากการตั้งค่า
            Dim regex As New Regex(extractPattern)
            Dim match As Match = regex.Match(qrData)

            If match.Success AndAlso match.Groups.Count > 1 Then
                Return match.Groups(1).Value
            End If

            ' ถ้าไม่พบข้อมูลตามรูปแบบที่กำหนด ให้ return ข้อมูลเดิม
            Return qrData

        Catch ex As Exception
            ' ในกรณีที่เกิดข้อผิดพลาด ให้ return ข้อมูลเดิม
            WriteLog($"Error in ExtractProductCode: {ex.Message}")
            Return qrData
        End Try
    End Function

    Private Sub PlaySystemSound()
        Try
            ' เล่นเสียงระบบ
            System.Media.SystemSounds.Beep.Play()
        Catch
            ' ไม่ต้องทำอะไรถ้าเล่นเสียงไม่ได้
        End Try
    End Sub

    Private Sub UpdateStatusBar(message As String)
        Try
            Dim statusStrip = Me.Controls.OfType(Of StatusStrip)().FirstOrDefault()
            If statusStrip IsNot Nothing Then
                Dim statusLabel = CType(statusStrip.Items.Find("statusLabel", False).FirstOrDefault(), ToolStripStatusLabel)
                If statusLabel IsNot Nothing Then
                    statusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}"
                End If
            End If
        Catch
            ' ไม่ต้องทำอะไรถ้าอัปเดต status bar ไม่ได้
        End Try
    End Sub

    Private Sub WriteLog(message As String)
        Try
            If File.Exists("Settings.config") Then
                Dim tempSettings As New frmSettings()
                Dim enableLogging As Boolean = CBool(tempSettings.GetSetting("enablelogging"))
                Dim logPath As String = CStr(tempSettings.GetSetting("logpath"))

                If enableLogging AndAlso Not String.IsNullOrEmpty(logPath) Then
                    If Not Directory.Exists(logPath) Then
                        Directory.CreateDirectory(logPath)
                    End If

                    Dim logFile As String = Path.Combine(logPath, $"ScanLog_{DateTime.Now:yyyyMMdd}.log")
                    Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}"

                    File.AppendAllText(logFile, logEntry)
                End If

                tempSettings.Dispose()
            End If
        Catch
            ' ไม่ต้องทำอะไรถ้าเขียน log ไม่ได้
        End Try
    End Sub

    ' Event Handlers สำหรับเมนูและปุ่ม
    Private Sub OpenSettings(sender As Object, e As EventArgs)
        Try
            Dim settingsForm As New frmSettings()
            If settingsForm.ShowDialog() = DialogResult.OK Then
                ' โหลดการตั้งค่าใหม่
                LoadSettingsFromConfig()
                ApplyThemeSettings()
                UpdateStatusBar("อัปเดตการตั้งค่าเรียบร้อยแล้ว")
            End If
            settingsForm.Dispose()
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเปิดหน้าการตั้งค่า: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExitApplication(sender As Object, e As EventArgs)
        Application.Exit()
    End Sub

    Private Sub ShowAbout(sender As Object, e As EventArgs)
        MessageBox.Show("QR Code Scanner v1.0" & vbNewLine &
                       "ระบบสแกน QR Code พร้อมการดึงข้อมูลอัตโนมัติ" & vbNewLine & vbNewLine &
                       "พัฒนาโดย: คุณ" & vbNewLine &
                       "วันที่: " & DateTime.Now.ToString("yyyy-MM-dd"),
                       "เกี่ยวกับโปรแกรม", MessageBoxButtons.OK, MessageBoxIcon.Information)
    End Sub

    Private Sub TestScan(sender As Object, e As EventArgs)
        Try
            Dim testData As String = "R00C-191604255012766+Q000060+P20414-007700A000+D20250527+LPT0000000+V00C-191604+U0000000"
            ProcessBarcode(testData)
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการทดสอบ: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ClearData(sender As Object, e As EventArgs)
        Try
            Dim txtBarcode = CType(Me.Controls.Find("txtBarcode", True).FirstOrDefault(), TextBox)
            If txtBarcode IsNot Nothing Then
                txtBarcode.Clear()
            End If
            UpdateStatusBar("ล้างข้อมูลแล้ว")
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการล้างข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmMenu_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        Try
            UnhookWindowsHookEx(_hookID)
        Catch
            ' ไม่ต้องทำอะไรถ้า unhook ไม่สำเร็จ
        End Try
    End Sub

    ' เมธอดสำหรับการอัปเดตการตั้งค่าจากภายนอก
    Public Sub RefreshSettings()
        LoadSettingsFromConfig()
        ApplyThemeSettings()
    End Sub

    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click
        OpenSettings(sender, e)
    End Sub

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        ClearData(sender, e)
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ExitApplication(sender, e)
    End Sub
End Class