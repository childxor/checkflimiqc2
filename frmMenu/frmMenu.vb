Imports System.Runtime.InteropServices
Imports System.Windows.Forms
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Xml
Imports System.Data.SqlClient
Imports System.Configuration

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
        ApplyThemeSettings()
        InitializeDatabase()
        UpdateStatusBar("พร้อมรับการสแกน QR Code")

        ' อัปเดตชื่อโปรแกรมด้วยเวอร์ชันจาก Assembly
        UpdateFormTitleWithVersion()
    End Sub

    ' Private Sub InitializeUI()
    '     Try
    '         ' ใช้ controls ที่มีอยู่แล้วจาก Designer
    '         LoadSettingsFromConfig()
    '         ApplyThemeSettings()
    '     Catch ex As Exception
    '         MessageBox.Show($"เกิดข้อผิดพลาดในการเริ่มต้นหน้าจอ: {ex.Message}",
    '                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '     End Try
    ' End Sub

    ' Private Sub CreateButtons()
    '     Try
    '         ' ปุ่มการตั้งค่า
    '         If Me.Controls.Find("btnSettings", True).Length = 0 Then
    '             Dim btnSettings As New Button()
    '             btnSettings.Name = "btnSettings"
    '             btnSettings.Text = "การตั้งค่า"
    '             btnSettings.Location = New Point(450, 50)
    '             btnSettings.Size = New Size(100, 30)
    '             AddHandler btnSettings.Click, AddressOf OpenSettings
    '             Me.Controls.Add(btnSettings)
    '         End If

    '         ' ปุ่มทดสอบ
    '         If Me.Controls.Find("btnTest", True).Length = 0 Then
    '             Dim btnTest As New Button()
    '             btnTest.Name = "btnTest"
    '             btnTest.Text = "ทดสอบ"
    '             btnTest.Location = New Point(560, 50)
    '             btnTest.Size = New Size(100, 30)
    '             AddHandler btnTest.Click, AddressOf TestScan
    '             Me.Controls.Add(btnTest)
    '         End If

    '         ' ปุ่มล้างข้อมูล
    '         If Me.Controls.Find("btnClear", True).Length = 0 Then
    '             Dim btnClear As New Button()
    '             btnClear.Name = "btnClear"
    '             btnClear.Text = "ล้างข้อมูล"
    '             btnClear.Location = New Point(670, 50)
    '             btnClear.Size = New Size(100, 30)
    '             AddHandler btnClear.Click, AddressOf ClearData
    '             Me.Controls.Add(btnClear)
    '         End If

    '     Catch ex As Exception
    '         MessageBox.Show($"เกิดข้อผิดพลาดในการสร้างปุ่ม: {ex.Message}",
    '                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
    '     End Try
    ' End Sub

    ''' <summary>
    ''' เปิดหน้าประวัติการสแกน
    ''' </summary>
    Private Sub OpenHistoryForm()
        Try
            If Not DatabaseManager.IsConnected Then
                MessageBox.Show("ไม่สามารถเชื่อมต่อฐานข้อมูลได้" & vbNewLine &
                              "กรุณาตรวจสอบการตั้งค่าฐานข้อมูล",
                              "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim historyForm As New frmHistory()
            historyForm.ShowDialog()
            historyForm.Dispose()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเปิดหน้าประวัติ: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub LoadSettingsFromConfig()
        Try
            If File.Exists("Settings.config") Then
                Dim doc As New XmlDocument()
                doc.Load("Settings.config")

                scanTimeout = GetSettingValueFromXML(doc, "ScanTimeout", 100)
                showFullData = GetSettingValueFromXML(doc, "ShowFullData", False)
                autoExtract = GetSettingValueFromXML(doc, "AutoExtract", True)
                soundEnabled = GetSettingValueFromXML(doc, "SoundEnabled", True)
                extractPattern = GetSettingValueFromXML(doc, "ExtractPattern", "\+P([^+]+)\+D")
            End If
        Catch ex As Exception
            ResetToDefaultSettings()
        End Try
    End Sub

    Private Sub ResetToDefaultSettings()
        scanTimeout = 100
        showFullData = False
        autoExtract = True
        soundEnabled = True
        extractPattern = "\+P([^+]+)\+D"
    End Sub

    Private Function GetSettingValueFromXML(doc As XmlDocument, key As String, defaultValue As Object) As Object
        Try
            Dim node As XmlNode = doc.SelectSingleNode($"//Setting[@key='{key}']")
            If node IsNot Nothing Then
                Dim value As String = node.Attributes("value").Value
                Select Case defaultValue.GetType()
                    Case GetType(Boolean)
                        Return Boolean.Parse(value)
                    Case GetType(Integer)
                        Return Integer.Parse(value)
                    Case Else
                        Return value
                End Select
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    Private Sub ApplyThemeSettings()
        Try
            If File.Exists("Settings.config") Then
                Dim doc As New XmlDocument()
                doc.Load("Settings.config")

                Dim theme As String = GetSettingValueFromXML(doc, "Theme", "Light")
                Dim showStatusBar As Boolean = GetSettingValueFromXML(doc, "ShowStatusBar", True)

                ' ปรับธีม
                Select Case theme.ToLower()
                    Case "dark"
                        ApplyDarkTheme()
                    Case "light"
                        ApplyLightTheme()
                End Select

                ' แสดง/ซ่อน status bar
                statusStrip.Visible = showStatusBar
            End If
        Catch
            ApplyLightTheme()
        End Try
    End Sub

    Private Sub ApplyDarkTheme()
        Me.BackColor = Color.FromArgb(45, 45, 48)
        Me.ForeColor = Color.White
        pnlMain.BackColor = Color.FromArgb(37, 37, 38)
        pnlHeader.BackColor = Color.FromArgb(41, 128, 185)
    End Sub

    Private Sub ApplyLightTheme()
        Me.BackColor = SystemColors.Control
        Me.ForeColor = SystemColors.ControlText
        pnlMain.BackColor = Color.FromArgb(248, 249, 250)
        pnlHeader.BackColor = Color.FromArgb(41, 128, 185)
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
            toolStripStatusLabel.Text = $"{DateTime.Now:HH:mm:ss} - {message}"
        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub

    Private Sub WriteLog(message As String)
        Try
            If File.Exists("Settings.config") Then
                Dim doc As New XmlDocument()
                doc.Load("Settings.config")

                Dim enableLogging As Boolean = GetSettingValueFromXML(doc, "EnableLogging", True)
                Dim logPath As String = GetSettingValueFromXML(doc, "LogPath", Path.Combine(System.Windows.Forms.Application.StartupPath, "Logs"))

                If enableLogging AndAlso Not String.IsNullOrEmpty(logPath) Then
                    If Not Directory.Exists(logPath) Then
                        Directory.CreateDirectory(logPath)
                    End If

                    Dim logFile As String = Path.Combine(logPath, $"ScanLog_{DateTime.Now:yyyyMMdd}.log")
                    Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}"

                    File.AppendAllText(logFile, logEntry)
                End If
            End If
        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub

    ' Event Handlers สำหรับปุ่มต่างๆ
    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click
        Try
            Dim settingsForm As New frmSettings()
            If settingsForm.ShowDialog() = DialogResult.OK Then
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

    Private Sub btnClear_Click(sender As Object, e As EventArgs) Handles btnClear.Click
        Try
            txtBarcode.Clear()
            lblBarcodeValue.Text = "No barcode scanned"
            lblBarcodeValue.ForeColor = Color.FromArgb(46, 125, 50)
            lblScanTime.Text = "Never scanned"
            txtBarcode.BackColor = SystemColors.Window
            picStatusIcon.BackColor = Color.FromArgb(255, 159, 67)
            lblStatusValue.Text = "Ready to scan..."
            lblStatusValue.ForeColor = Color.FromArgb(255, 159, 67)
            UpdateStatusBar("ล้างข้อมูลแล้ว")
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการล้างข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnExit_Click(sender As Object, e As EventArgs) Handles btnExit.Click
        ' ออกจากโปรแกรม
        System.Windows.Forms.Application.Exit()
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

    ''' <summary>
    ''' อัปเดตชื่อโปรแกรมด้วยเวอร์ชันจาก Assembly
    ''' </summary>
    Private Sub UpdateFormTitleWithVersion()
        Try
            Dim version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Dim versionString As String = $"v{version.Major}.{version.Minor}.{version.Build}"

            ' อัปเดตชื่อในหัวข้อฟอร์ม
            Me.Text = $"QR Code Scanner System {versionString}"

            ' อัปเดตชื่อในหัวข้อหลัก
            lblTitle.Text = $"QR Code Scanner System {versionString}"

        Catch ex As Exception
            ' ถ้าอ่านเวอร์ชันไม่ได้ ให้ใช้ชื่อเดิม
            Me.Text = "QR Code Scanner System"
            lblTitle.Text = "QR Code Scanner System"
            WriteLog($"Error reading assembly version: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ประมวลผล barcode ที่สแกนได้ พร้อมการตรวจสอบและบันทึกฐานข้อมูล
    ''' </summary>
    Public Sub ProcessBarcode(barcode As String)
        Try
            ' ตรวจสอบความถูกต้องของ barcode
            Dim validation As BarcodeValidationResult = BarcodeValidator.ValidateBarcode(barcode)

            ' ดึงข้อมูลจาก barcode
            Dim extractionMode As ExtractionMode = GetExtractionMode()
            Dim extractedData As BarcodeExtractedData = BarcodeValidator.ExtractBarcodeData(barcode, extractionMode)

            ' แสดงผลในหน้าจอ
            Me.Invoke(Sub()
                          DisplayBarcodeResult(extractedData, validation)

                          ' เล่นเสียงตามผลการตรวจสอบ
                          If soundEnabled Then
                              PlaySoundByResult(validation)
                          End If

                          ' แสดงข้อความแจ้งเตือนตามการตั้งค่า
                          ShowBarcodeMessage(extractedData, validation)

                          ' บันทึกลงฐานข้อมูล
                          SaveToDatabase(extractedData, validation)

                          ' บันทึก log
                          WriteBarcodeLog(extractedData, validation)
                      End Sub)

        Catch ex As Exception
            Me.Invoke(Sub()
                          UpdateStatusBar($"เกิดข้อผิดพลาด: {ex.Message}")
                          WriteLog($"Error in ProcessBarcode: {ex.Message}")
                      End Sub)
        End Try
    End Sub

    ''' <summary>
    ''' บันทึกข้อมูลลงฐานข้อมูล
    ''' </summary>

    Private Sub SaveToDatabase(extractedData As BarcodeExtractedData, validation As BarcodeValidationResult)
        Try
            If DatabaseManager.IsConnected Then
                ' แปลง Quantity จาก String เป็น Integer
                Dim quantityValue As Integer = 0
                If Not String.IsNullOrEmpty(extractedData.Quantity) Then
                    If Not Integer.TryParse(extractedData.Quantity, quantityValue) Then
                        quantityValue = 0
                        WriteLog($"Warning: Could not parse quantity '{extractedData.Quantity}' to integer, using 0")
                    End If
                End If

                Dim scanRecord As New ScanDataRecord() With {
                    .ScanDateTime = DateTime.Now,
                    .OriginalData = extractedData.OriginalData,
                    .ExtractedData = extractedData.ExtractedValue,
                    .ProductCode = extractedData.ProductCode,
                    .ReferenceCode = extractedData.ReferenceCode,
                    .Quantity = quantityValue,
                    .DateCode = extractedData.DateCode,
                    .IsValid = validation.IsValid,
                    .ValidationMessages = If(validation.ValidationMessages IsNot Nothing,
                                           String.Join("; ", validation.ValidationMessages), ""),
                    .ComputerName = Environment.MachineName,
                    .UserName = Environment.UserName
                }

                Dim recordId As Integer = DatabaseManager.SaveScanData(scanRecord)
                If recordId > 0 Then
                    WriteLog($"Successfully saved scan record with ID: {recordId}")
                Else
                    WriteLog("Failed to save scan data to database")
                    MessageBox.Show("ไม่สามารถบันทึกข้อมูลลงฐานข้อมูลได้" & vbNewLine &
                                  "กรุณาตรวจสอบการเชื่อมต่อฐานข้อมูล",
                                  "ข้อผิดพลาดการบันทึกข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            Else
                WriteLog("Database is not connected - cannot save scan data")
                MessageBox.Show("ไม่สามารถเชื่อมต่อฐานข้อมูลได้" & vbNewLine &
                              "ข้อมูลการสแกนจะไม่ถูกบันทึก",
                              "ปัญหาการเชื่อมต่อฐานข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            WriteLog($"Error saving to database: {ex.Message}")
            ' แสดง error message ให้ผู้ใช้เห็น
            MessageBox.Show($"เกิดข้อผิดพลาดในการบันทึกข้อมูล:" & vbNewLine &
                          $"รายละเอียด: {ex.Message}" & vbNewLine & vbNewLine &
                          $"กรุณาตรวจสอบ:" & vbNewLine &
                          $"1. ไฟล์ฐานข้อมูล Access มีอยู่จริง" & vbNewLine &
                          $"2. มีสิทธิ์ในการเขียนไฟล์" & vbNewLine &
                          $"3. ไม่มีโปรแกรมอื่นเปิดไฟล์ฐานข้อมูลอยู่",
                          "ข้อผิดพลาดการบันทึกข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Me.Invoke(Sub()
                          UpdateStatusBar($"ไม่สามารถบันทึกข้อมูลลงฐานข้อมูลได้: {ex.Message}")
                      End Sub)
        End Try
    End Sub

    ''' <summary>
    ''' แสดงผลลัพธ์การสแกนในหน้าจอ
    ''' </summary>
    Private Sub DisplayBarcodeResult(extractedData As BarcodeExtractedData, validation As BarcodeValidationResult)
        Try
            ' อัปเดตข้อความในกล่องข้อความ
            txtBarcode.Text = extractedData.ExtractedValue
            lblBarcodeValue.Text = extractedData.ProductCode

            ' เปลี่ยนสีตามสถานะ
            If validation.IsValid Then
                txtBarcode.BackColor = Color.LightGreen
                lblBarcodeValue.ForeColor = Color.FromArgb(46, 125, 50)
                picStatusIcon.BackColor = Color.FromArgb(46, 125, 50)
                lblStatusValue.Text = "สแกนสำเร็จ"
                lblStatusValue.ForeColor = Color.FromArgb(46, 125, 50)
            ElseIf validation.IsPartiallyValid Then
                txtBarcode.BackColor = Color.LightYellow
                lblBarcodeValue.ForeColor = Color.FromArgb(255, 159, 67)
                picStatusIcon.BackColor = Color.FromArgb(255, 159, 67)
                lblStatusValue.Text = "สแกนสำเร็จ (มีคำเตือน)"
                lblStatusValue.ForeColor = Color.FromArgb(255, 159, 67)
            Else
                txtBarcode.BackColor = Color.LightPink
                lblBarcodeValue.ForeColor = Color.FromArgb(231, 76, 60)
                picStatusIcon.BackColor = Color.FromArgb(231, 76, 60)
                lblStatusValue.Text = "สแกนไม่สมบูรณ์"
                lblStatusValue.ForeColor = Color.FromArgb(231, 76, 60)
            End If

            ' อัปเดตเวลาการสแกนล่าสุด
            lblScanTime.Text = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss")

            ' อัปเดต status bar
            Dim statusMessage As String = ""
            If validation.IsValid Then
                statusMessage = $"สแกนสำเร็จ: {extractedData.ProductCode}"
            ElseIf validation.IsPartiallyValid Then
                statusMessage = $"สแกนสำเร็จ (มีคำเตือน): {extractedData.ProductCode}"
            Else
                statusMessage = $"สแกนไม่สมบูรณ์: {extractedData.ExtractedValue}"
            End If

            UpdateStatusBar(statusMessage)

        Catch ex As Exception
            WriteLog($"Error in DisplayBarcodeResult: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' แสดงข้อความแจ้งเตือนตามการตั้งค่า
    ''' </summary>
    Private Sub ShowBarcodeMessage(extractedData As BarcodeExtractedData, validation As BarcodeValidationResult)
        Try
            If showFullData Then
                Dim message As String = BuildDetailedMessage(extractedData, validation)
                Dim icon As MessageBoxIcon = MessageBoxIcon.Information

                If Not validation.IsValid Then
                    icon = MessageBoxIcon.Warning
                End If

                MessageBox.Show(message, "ผลการสแกน QR Code", MessageBoxButtons.OK, icon)
            End If

        Catch ex As Exception
            WriteLog($"Error in ShowBarcodeMessage: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' สร้างข้อความรายละเอียด
    ''' </summary>
    Private Function BuildDetailedMessage(extractedData As BarcodeExtractedData, validation As BarcodeValidationResult) As String
        Dim message As New System.Text.StringBuilder()

        message.AppendLine("=== ผลการสแกน QR Code ===")
        message.AppendLine()

        ' แสดงข้อมูลที่ดึงออกมา
        message.AppendLine("ข้อมูลที่ดึงออกมา:")
        If Not String.IsNullOrEmpty(extractedData.ProductCode) Then
            message.AppendLine($"  รหัสผลิตภัณฑ์: {extractedData.ProductCode}")
        End If
        If Not String.IsNullOrEmpty(extractedData.ReferenceCode) Then
            message.AppendLine($"  รหัสอ้างอิง: {extractedData.ReferenceCode}")
        End If
        If Not String.IsNullOrEmpty(extractedData.Quantity) Then
            message.AppendLine($"  จำนวน: {extractedData.Quantity}")
        End If
        If Not String.IsNullOrEmpty(extractedData.DateCode) Then
            message.AppendLine($"  วันที่: {extractedData.DateCode}")
        End If

        message.AppendLine()

        ' แสดงสถานะการตรวจสอบ
        If validation.IsValid Then
            message.AppendLine("✅ สถานะ: ข้อมูลถูกต้องสมบูรณ์")
        ElseIf validation.IsPartiallyValid Then
            message.AppendLine("⚠️ สถานะ: ข้อมูลถูกต้องบางส่วน")
        Else
            message.AppendLine("❌ สถานะ: ข้อมูลไม่สมบูรณ์")
        End If

        ' แสดงข้อความเตือน (ถ้ามี)
        If validation.ValidationMessages IsNot Nothing AndAlso validation.ValidationMessages.Count > 0 Then
            message.AppendLine()
            message.AppendLine("ข้อความเตือน:")
            For Each msg As String In validation.ValidationMessages
                message.AppendLine($"  • {msg}")
            Next
        End If

        ' แสดงข้อมูลต้นฉบับ
        message.AppendLine()
        message.AppendLine("ข้อมูลต้นฉบับ:")
        message.AppendLine(extractedData.OriginalData)

        Return message.ToString()
    End Function

    ''' <summary>
    ''' ดึงโหมดการ extract ข้อมูลจากการตั้งค่า
    ''' </summary>
    Private Function GetExtractionMode() As ExtractionMode
        Try
            If File.Exists("Settings.config") Then
                ' อ่านจากไฟล์ config
                ' สามารถเพิ่มการตั้งค่านี้ใน frmSettings ได้
                Return ExtractionMode.Intelligent
            End If
        Catch
            ' ไม่ต้องทำอะไร
        End Try

        ' ใช้โหมด Intelligent เป็นค่าเริ่มต้น
        Return ExtractionMode.Intelligent
    End Function

    ''' <summary>
    ''' เล่นเสียงสำเร็จ
    ''' </summary>
    ''' <summary>
    ''' เล่นเสียงตามผลการตรวจสอบ
    ''' </summary>
    Private Sub PlaySoundByResult(validation As BarcodeValidationResult)
        Try
            If validation.IsValid Then
                System.Media.SystemSounds.Asterisk.Play()
            ElseIf validation.IsPartiallyValid Then
                System.Media.SystemSounds.Exclamation.Play()
            Else
                System.Media.SystemSounds.Hand.Play()
            End If
        Catch
            System.Media.SystemSounds.Beep.Play()
        End Try
    End Sub

    ''' <summary>
    ''' บันทึก log การสแกน
    ''' </summary>
    Private Sub WriteBarcodeLog(extractedData As BarcodeExtractedData, validation As BarcodeValidationResult)
        Try
            Dim logMessage As String = $"Scanned: {extractedData.OriginalData} | " &
                                  $"Extracted: {extractedData.ExtractedValue} | " &
                                  $"Valid: {validation.IsValid} | " &
                                  $"Warnings: {If(validation.ValidationMessages?.Count, 0)}"

            WriteLog(logMessage)

        Catch ex As Exception
            WriteLog($"Error in WriteBarcodeLog: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ทดสอบการสแกนด้วยข้อมูลตัวอย่าง
    ''' </summary>
    ''' <summary>
    ''' ทดสอบการสแกนด้วยข้อมูลตัวอย่าง
    ''' </summary>
    Private Sub TestScanWithValidation()
        Try
            ' ข้อมูลทดสอบหลายแบบ
            Dim testData() As String = {
                "R00C-191604255012766+Q000060+P20414-007700A000+D20250527+LPT0000000+V00C-191604+U0000000", ' ข้อมูลสมบูรณ์
                "R00C-191604255012766+Q000060+P20414-007700A000+D20250527", ' ข้อมูลบางส่วน
                "P20414-007700A000+D20250527", ' ข้อมูลน้อย
                "InvalidBarcodeData123" ' ข้อมูลไม่ถูกต้อง
            }

            For Each data As String In testData
                ProcessBarcode(data)
                System.Threading.Thread.Sleep(1000) ' หน่วงเวลาเพื่อดูผล
            Next

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการทดสอบ: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' เพิ่มเมนูประวัติการสแกน (เรียกจาก Designer หรือ Load event)
    ''' </summary>
    Private Sub AddHistoryMenuItem()
        Try
            ' เพิ่มปุ่มประวัติในหน้าหลัก (ถ้าต้องการ)
            ' หรือเพิ่มในเมนูบาร์
        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub

    ''' <summary>
    ''' เริ่มต้นการเชื่อมต่อฐานข้อมูล
    ''' </summary>
    Private Sub InitializeDatabase()
        Try
            Console.WriteLine("Starting database initialization...")
            WriteLog("Starting database initialization...")
            
            ' เริ่มต้นฐานข้อมูล
            If DatabaseManager.Initialize() Then
                UpdateStatusBar("เชื่อมต่อฐานข้อมูลสำเร็จ")
                WriteLog("Database initialization successful")
                Console.WriteLine("Database initialization successful")
            Else
                UpdateStatusBar("ไม่สามารถเชื่อมต่อฐานข้อมูลได้ - ระบบจะทำงานแบบ Offline")
                WriteLog("Database initialization failed - running in offline mode")
                Console.WriteLine("Database initialization failed")
            End If
        Catch ex As Exception
            WriteLog($"Database initialization error: {ex.Message}")
            WriteLog($"Stack trace: {ex.StackTrace}")
            UpdateStatusBar("ระบบทำงานแบบ Offline")
            Console.WriteLine($"Database initialization error: {ex.Message}")
        End Try
    End Sub

    Private Sub btnHistory_Click(sender As Object, e As EventArgs) Handles btnHistory.Click
        OpenHistoryForm()
    End Sub

    ''' <summary>
    ''' ทดสอบการเชื่อมต่อฐานข้อมูลและการบันทึกข้อมูล
    ''' </summary>
    Private Sub TestDatabaseConnection()
        Try
            ' ทดสอบการเชื่อมต่อ
            If Not DatabaseManager.IsConnected Then
                MessageBox.Show("ไม่สามารถเชื่อมต่อฐานข้อมูลได้" & vbNewLine &
                              "กรุณาตรวจสอบ:" & vbNewLine &
                              "1. ไฟล์ฐานข้อมูล Access มีอยู่จริง" & vbNewLine &
                              "2. มีสิทธิ์ในการเขียนไฟล์" & vbNewLine &
                              "3. ไม่มีโปรแกรมอื่นเปิดไฟล์ฐานข้อมูลอยู่",
                              "ปัญหาการเชื่อมต่อฐานข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' ทดสอบการบันทึกข้อมูล
            Dim testRecord As New ScanDataRecord() With {
                .ScanDateTime = DateTime.Now,
                .OriginalData = "TEST_DATA_R00C-191604255012766+Q000001+P99999-TEST+D20250101",
                .ExtractedData = "99999-TEST",
                .ProductCode = "99999-TEST",
                .ReferenceCode = "00C-191604255012766",
                .Quantity = 1,
                .DateCode = "20250101",
                .IsValid = True,
                .ValidationMessages = "ทดสอบระบบ",
                .ComputerName = Environment.MachineName,
                .UserName = Environment.UserName
            }

            Dim recordId As Integer = DatabaseManager.SaveScanData(testRecord)

            If recordId > 0 Then
                MessageBox.Show($"ทดสอบสำเร็จ!" & vbNewLine &
                              $"บันทึกข้อมูลทดสอบได้ ID: {recordId}" & vbNewLine &
                              $"ระบบฐานข้อมูลทำงานปกติ",
                              "ผลการทดสอบ", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' ลบข้อมูลทดสอบ
                DatabaseManager.DeleteScanRecord(recordId)
                WriteLog($"Test record created and deleted successfully (ID: {recordId})")
            Else
                MessageBox.Show("ไม่สามารถบันทึกข้อมูลทดสอบได้" & vbNewLine &
                              "กรุณาตรวจสอบ log ไฟล์เพื่อดูรายละเอียดเพิ่มเติม" & vbNewLine & vbNewLine &
                              "สาเหตุที่เป็นไปได้:" & vbNewLine &
                              "1. ไฟล์ฐานข้อมูล Access ไม่มีอยู่" & vbNewLine &
                              "2. ไม่มีสิทธิ์ในการเขียนไฟล์" & vbNewLine &
                              "3. โครงสร้างตารางไม่ถูกต้อง" & vbNewLine &
                              "4. ไฟล์ฐานข้อมูลถูกล็อคโดยโปรแกรมอื่น",
                              "ทดสอบไม่สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการทดสอบ:" & vbNewLine &
                          $"รายละเอียด: {ex.Message}" & vbNewLine & vbNewLine &
                          $"กรุณาตรวจสอบ:" & vbNewLine &
                          $"1. ไฟล์ฐานข้อมูล Access มีอยู่จริง" & vbNewLine &
                          $"2. มีสิทธิ์ในการเขียนไฟล์" & vbNewLine &
                          $"3. ไม่มีโปรแกรมอื่นเปิดไฟล์ฐานข้อมูลอยู่" & vbNewLine &
                          $"4. Microsoft Access Database Engine ติดตั้งแล้ว",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            WriteLog($"Database test error: {ex.Message}")
            WriteLog($"Stack trace: {ex.StackTrace}")
        End Try
    End Sub

    ''' <summary>
    ''' ทดสอบการสแกนพร้อมบันทึกฐานข้อมูล
    ''' </summary>
    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        Try
            ' แสดงตัวเลือกการทดสอบ
            Dim result As DialogResult = MessageBox.Show(
                "เลือกประเภทการทดสอบ:" & vbNewLine &
                "Yes = ทดสอบการเชื่อมต่อฐานข้อมูล" & vbNewLine &
                "No = ทดสอบการสแกน QR Code" & vbNewLine &
                "Cancel = ยกเลิก",
                "เลือกการทดสอบ",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            Select Case result
                Case DialogResult.Yes
                    TestDatabaseConnection()
                Case DialogResult.No
                    TestScanWithValidation()
                Case DialogResult.Cancel
                    Return
            End Select

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการทดสอบ: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
End Class