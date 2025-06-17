Imports System.IO
Imports System.Text.RegularExpressions
Imports System.Data.OleDb
Imports System.Configuration
Imports System.Xml
Imports Microsoft.Win32


Public Class frmSettings

#Region "Constants & Variables"
    Private Const CONFIG_FILE As String = "Settings.config"
    Private originalSettings As Dictionary(Of String, Object)
    Private hasChanges As Boolean = False
    Private Const DEFAULT_ACCESS_PATH As String = "QRCodeScanner.accdb"
#End Region

#Region "Form Events"
    Private Sub frmSettings_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()
        LoadSettings()
        SetTooltips()
        BackupOriginalSettings()
    End Sub

    Private Sub frmSettings_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If hasChanges Then
            Dim result As DialogResult = MessageBox.Show(
                "มีการเปลี่ยนแปลงการตั้งค่าที่ยังไม่ได้บันทึก" & vbNewLine &
                "คุณต้องการบันทึกการเปลี่ยนแปลงหรือไม่?",
                "ยืนยันการบันทึก",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question)

            Select Case result
                Case DialogResult.Yes
                    SaveSettings()
                Case DialogResult.Cancel
                    e.Cancel = True
            End Select
        End If
    End Sub

    Private Sub frmSettings_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        WireAdditionalEvents()
        CheckDiskSpace()
    End Sub
#End Region

#Region "Initialization"
    Private Sub InitializeForm()
        Try
            ' ตั้งค่าเริ่มต้นสำหรับ ComboBox
            If cmbTheme.Items.Count > 0 Then cmbTheme.SelectedIndex = 0
            If cmbLanguage.Items.Count > 0 Then cmbLanguage.SelectedIndex = 0
            If cmbFontSize.Items.Count > 1 Then cmbFontSize.SelectedIndex = 1
            If cmbBackupInterval.Items.Count > 0 Then cmbBackupInterval.SelectedIndex = 0
            If cmbLogLevel.Items.Count > 2 Then cmbLogLevel.SelectedIndex = 2

            ' ตั้งค่าเริ่มต้นสำหรับ NumericUpDown
            numScanTimeout.Value = 100
            numNotificationDuration.Value = 5
            numMaxRecords.Value = 1000
            numCleanupDays.Value = 30

            ' ตั้งค่าเริ่มต้นสำหรับ TextBox
            txtDatabase.Text = DEFAULT_ACCESS_PATH
            txtExtractPattern.Text = "\+P([^+]+)\+D"
            txtLogPath.Text = Path.Combine(System.Windows.Forms.Application.StartupPath, "Logs")
            txtBackupPath.Text = Path.Combine(System.Windows.Forms.Application.StartupPath, "Backup")

            ' เปิดใช้งานควบคุมตามค่าเริ่มต้น
            UpdateControlStates()
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตั้งค่าเริ่มต้น: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    Private Sub SetTooltips()
        Try
            toolTip.SetToolTip(numScanTimeout, "ระยะเวลารอระหว่างการกดคีย์ในการสแกน (มิลลิวินาที)")
            toolTip.SetToolTip(txtExtractPattern, "รูปแบบ Regular Expression สำหรับดึงข้อมูลจาก QR Code")
            toolTip.SetToolTip(txtDatabase, "พาธของฐานข้อมูล Access")
            toolTip.SetToolTip(numMaxRecords, "จำนวนข้อมูลสูงสุดที่แสดงในหน้าเดียว")
            toolTip.SetToolTip(btnTestPattern, "ทดสอบรูปแบบการดึงข้อมูลกับข้อมูลตัวอย่าง")
        Catch
            ' ไม่ต้องทำอะไรถ้า tooltip ไม่สามารถตั้งได้
        End Try
    End Sub
#End Region

#Region "Settings Management"
    Private Sub LoadSettings()
        Try
            If File.Exists(CONFIG_FILE) Then
                Dim doc As New XmlDocument()
                doc.Load(CONFIG_FILE)

                ' Scanner Settings
                numScanTimeout.Value = GetSettingValue(doc, "ScanTimeout", 100)
                chkShowFullData.Checked = GetSettingValue(doc, "ShowFullData", False)
                chkAutoExtract.Checked = GetSettingValue(doc, "AutoExtract", True)
                chkSoundEnabled.Checked = GetSettingValue(doc, "SoundEnabled", True)
                txtExtractPattern.Text = GetSettingValue(doc, "ExtractPattern", "\+P([^+]+)\+D")

                ' Database Settings
                txtDatabase.Text = GetSettingValue(doc, "AccessDatabasePath", DEFAULT_ACCESS_PATH)
                txtPassword.Text = GetSettingValue(doc, "AccessPassword", "")

                ' Backup Settings
                chkAutoBackup.Checked = GetSettingValue(doc, "AutoBackup", False)
                txtBackupPath.Text = GetSettingValue(doc, "BackupPath", Path.Combine(System.Windows.Forms.Application.StartupPath, "Backup"))
                cmbBackupInterval.Text = GetSettingValue(doc, "BackupInterval", "ทุกวัน")

                ' Display Settings
                cmbTheme.Text = GetSettingValue(doc, "Theme", "Light")
                cmbLanguage.Text = GetSettingValue(doc, "Language", "ไทย")
                cmbFontSize.Text = GetSettingValue(doc, "FontSize", "ปกติ")
                chkShowStatusBar.Checked = GetSettingValue(doc, "ShowStatusBar", True)
                chkShowToolbar.Checked = GetSettingValue(doc, "ShowToolbar", True)

                ' Notification Settings
                chkShowNotifications.Checked = GetSettingValue(doc, "ShowNotifications", True)
                chkSoundNotifications.Checked = GetSettingValue(doc, "SoundNotifications", True)
                numNotificationDuration.Value = GetSettingValue(doc, "NotificationDuration", 5)

                ' Logging Settings
                chkEnableLogging.Checked = GetSettingValue(doc, "EnableLogging", True)
                cmbLogLevel.Text = GetSettingValue(doc, "LogLevel", "Info")
                txtLogPath.Text = GetSettingValue(doc, "LogPath", Path.Combine(System.Windows.Forms.Application.StartupPath, "Logs"))

                ' Performance Settings
                numMaxRecords.Value = GetSettingValue(doc, "MaxRecords", 1000)
                chkAutoCleanup.Checked = GetSettingValue(doc, "AutoCleanup", False)
                numCleanupDays.Value = GetSettingValue(doc, "CleanupDays", 30)
            End If

            UpdateControlStates()
            hasChanges = False

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดการตั้งค่า: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SaveSettings()
        Try
            Dim doc As New XmlDocument()
            Dim root As XmlElement = doc.CreateElement("Settings")
            doc.AppendChild(root)

            ' Scanner Settings
            AddSetting(doc, root, "ScanTimeout", numScanTimeout.Value.ToString())
            AddSetting(doc, root, "ShowFullData", chkShowFullData.Checked.ToString())
            AddSetting(doc, root, "AutoExtract", chkAutoExtract.Checked.ToString())
            AddSetting(doc, root, "SoundEnabled", chkSoundEnabled.Checked.ToString())
            AddSetting(doc, root, "ExtractPattern", txtExtractPattern.Text)

            ' Database Settings
            AddSetting(doc, root, "AccessDatabasePath", txtDatabase.Text)
            AddSetting(doc, root, "AccessPassword", EncryptPassword(txtPassword.Text))

            ' Backup Settings
            AddSetting(doc, root, "AutoBackup", chkAutoBackup.Checked.ToString())
            AddSetting(doc, root, "BackupPath", txtBackupPath.Text)
            AddSetting(doc, root, "BackupInterval", cmbBackupInterval.Text)

            ' Display Settings
            AddSetting(doc, root, "Theme", cmbTheme.Text)
            AddSetting(doc, root, "Language", cmbLanguage.Text)
            AddSetting(doc, root, "FontSize", cmbFontSize.Text)
            AddSetting(doc, root, "ShowStatusBar", chkShowStatusBar.Checked.ToString())
            AddSetting(doc, root, "ShowToolbar", chkShowToolbar.Checked.ToString())

            ' Notification Settings
            AddSetting(doc, root, "ShowNotifications", chkShowNotifications.Checked.ToString())
            AddSetting(doc, root, "SoundNotifications", chkSoundNotifications.Checked.ToString())
            AddSetting(doc, root, "NotificationDuration", numNotificationDuration.Value.ToString())

            ' Logging Settings
            AddSetting(doc, root, "EnableLogging", chkEnableLogging.Checked.ToString())
            AddSetting(doc, root, "LogLevel", cmbLogLevel.Text)
            AddSetting(doc, root, "LogPath", txtLogPath.Text)

            ' Performance Settings
            AddSetting(doc, root, "MaxRecords", numMaxRecords.Value.ToString())
            AddSetting(doc, root, "AutoCleanup", chkAutoCleanup.Checked.ToString()) 
            AddSetting(doc, root, "CleanupDays", numCleanupDays.Value.ToString())

            doc.Save(CONFIG_FILE)
            hasChanges = False

            MessageBox.Show("บันทึกการตั้งค่าเรียบร้อยแล้ว", "สำเร็จ",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการบันทึกการตั้งค่า: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function GetSettingValue(doc As XmlDocument, key As String, defaultValue As Object) As Object
        Try
            Dim node As XmlNode = doc.SelectSingleNode($"//Setting[@key='{key}']")
            If node IsNot Nothing Then
                Dim value As String = node.Attributes("value").Value

                Select Case defaultValue.GetType()
                    Case GetType(Boolean)
                        Return Boolean.Parse(value)
                    Case GetType(Integer)
                        Return Integer.Parse(value)
                    Case GetType(Decimal)
                        Return Decimal.Parse(value)
                    Case Else
                        If key = "Password" Then
                            Return DecryptPassword(value)
                        End If
                        Return value
                End Select
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    Private Sub AddSetting(doc As XmlDocument, parent As XmlElement, key As String, value As String)
        Dim setting As XmlElement = doc.CreateElement("Setting")
        setting.SetAttribute("key", key)
        setting.SetAttribute("value", value)
        parent.AppendChild(setting)
    End Sub

    Private Sub BackupOriginalSettings()
        Try
            originalSettings = New Dictionary(Of String, Object)
            ' เก็บค่าเดิมไว้สำหรับการ Reset
            originalSettings("ScanTimeout") = numScanTimeout.Value
            originalSettings("ShowFullData") = chkShowFullData.Checked
            originalSettings("AutoExtract") = chkAutoExtract.Checked
            ' เพิ่มการเก็บค่าอื่นๆ ตามต้องการ
        Catch
            ' ไม่ต้องทำอะไรถ้า backup ไม่สำเร็จ
        End Try
    End Sub
#End Region

#Region "Control State Management"
    Private Sub UpdateControlStates()
        Try
            ' เปิด/ปิดควบคุมตามเงื่อนไข
            txtPassword.Enabled = True
            lblPassword.Enabled = True

            txtBackupPath.Enabled = chkAutoBackup.Checked
            btnBrowseBackup.Enabled = chkAutoBackup.Checked
            cmbBackupInterval.Enabled = chkAutoBackup.Checked
            lblBackupPath.Enabled = chkAutoBackup.Checked
            lblBackupInterval.Enabled = chkAutoBackup.Checked

            txtLogPath.Enabled = chkEnableLogging.Checked
            btnBrowseLog.Enabled = chkEnableLogging.Checked
            btnOpenLogFolder.Enabled = chkEnableLogging.Checked
            cmbLogLevel.Enabled = chkEnableLogging.Checked
            lblLogLevel.Enabled = chkEnableLogging.Checked
            lblLogPath.Enabled = chkEnableLogging.Checked

            numNotificationDuration.Enabled = chkShowNotifications.Checked
            chkSoundNotifications.Enabled = chkShowNotifications.Checked
            lblNotificationDuration.Enabled = chkShowNotifications.Checked
            lblSeconds.Enabled = chkShowNotifications.Checked

            numCleanupDays.Enabled = chkAutoCleanup.Checked
            lblCleanupDays.Enabled = chkAutoCleanup.Checked
            lblDays.Enabled = chkAutoCleanup.Checked
        Catch
            ' ไม่ต้องทำอะไรถ้าเกิดข้อผิดพลาด
        End Try
    End Sub
#End Region

#Region "Event Handlers - CheckBox Changes"
    Private Sub chkAutoBackup_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoBackup.CheckedChanged
        UpdateControlStates()
        MarkAsChanged()
    End Sub

    Private Sub chkEnableLogging_CheckedChanged(sender As Object, e As EventArgs) Handles chkEnableLogging.CheckedChanged
        UpdateControlStates()
        MarkAsChanged()
    End Sub

    Private Sub chkShowNotifications_CheckedChanged(sender As Object, e As EventArgs) Handles chkShowNotifications.CheckedChanged
        UpdateControlStates()
        MarkAsChanged()
    End Sub

    Private Sub chkAutoCleanup_CheckedChanged(sender As Object, e As EventArgs) Handles chkAutoCleanup.CheckedChanged
        UpdateControlStates()
        MarkAsChanged()
    End Sub
#End Region

#Region "Event Handlers - Value Changes"
    Private Sub MarkAsChanged()
        hasChanges = True
        If btnApply IsNot Nothing Then
            btnApply.Enabled = True
        End If
    End Sub

    ' เพิ่ม Event Handler สำหรับการเปลี่ยนแปลงค่าต่างๆ
    Private Sub numScanTimeout_ValueChanged(sender As Object, e As EventArgs) Handles numScanTimeout.ValueChanged
        MarkAsChanged()
    End Sub

    Private Sub txtExtractPattern_TextChanged(sender As Object, e As EventArgs) Handles txtExtractPattern.TextChanged
        MarkAsChanged()
    End Sub

    Private Sub cmbTheme_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbTheme.SelectedIndexChanged
        ApplyTheme()
        MarkAsChanged()
    End Sub

    Private Sub cmbFontSize_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbFontSize.SelectedIndexChanged
        ApplyFontSize()
        MarkAsChanged()
    End Sub
#End Region

#Region "Button Events"
    Private Sub btnOK_Click(sender As Object, e As EventArgs) Handles btnOK.Click
        If ValidateSettings() Then
            SaveSettings()
            DialogResult = DialogResult.OK
            Close()
        End If
    End Sub

    Private Sub btnCancel_Click(sender As Object, e As EventArgs) Handles btnCancel.Click
        DialogResult = DialogResult.Cancel
        Close()
    End Sub

    Private Sub btnApply_Click(sender As Object, e As EventArgs) Handles btnApply.Click
        If ValidateSettings() Then
            SaveSettings()
            btnApply.Enabled = False
        End If
    End Sub

    Private Sub btnReset_Click(sender As Object, e As EventArgs) Handles btnReset.Click
        Dim result As DialogResult = MessageBox.Show(
            "คุณต้องการรีเซ็ตการตั้งค่าทั้งหมดกลับเป็นค่าเริ่มต้นหรือไม่?",
            "ยืนยันการรีเซ็ต",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

        If result = DialogResult.Yes Then
            ResetToDefaults()
        End If
    End Sub

    Private Sub btnTestPattern_Click(sender As Object, e As EventArgs) Handles btnTestPattern.Click
        TestExtractPattern()
    End Sub

    Private Sub btnRunTest_Click(sender As Object, e As EventArgs) Handles btnRunTest.Click
        RunPatternTest()
    End Sub

    Private Sub btnTestConnection_Click(sender As Object, e As EventArgs) Handles btnTestConnection.Click
        TestDatabaseConnection()
    End Sub

    Private Sub btnBrowseBackup_Click(sender As Object, e As EventArgs) Handles btnBrowseBackup.Click
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            txtBackupPath.Text = folderBrowserDialog.SelectedPath
            MarkAsChanged()
        End If
    End Sub

    Private Sub btnBrowseLog_Click(sender As Object, e As EventArgs) Handles btnBrowseLog.Click
        If folderBrowserDialog.ShowDialog() = DialogResult.OK Then
            txtLogPath.Text = folderBrowserDialog.SelectedPath
            MarkAsChanged()
        End If
    End Sub

    Private Sub btnBrowseDatabase_Click(sender As Object, e As EventArgs) Handles btnBrowseDatabase.Click
        Try
            ' สร้าง OpenFileDialog สำหรับเลือกไฟล์ฐานข้อมูล Access
            Dim fileDialog As New OpenFileDialog()
            fileDialog.Title = "เลือกไฟล์ฐานข้อมูล Access"
            fileDialog.Filter = "ไฟล์ฐานข้อมูล Access (*.accdb;*.mdb)|*.accdb;*.mdb|ไฟล์ทั้งหมด (*.*)|*.*"
            fileDialog.CheckFileExists = True
            
            If fileDialog.ShowDialog() = DialogResult.OK Then
                txtDatabase.Text = fileDialog.FileName
                MarkAsChanged()
            End If
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเลือกไฟล์: {ex.Message}", 
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnOpenLogFolder_Click(sender As Object, e As EventArgs) Handles btnOpenLogFolder.Click
        Try
            If Directory.Exists(txtLogPath.Text) Then
                Process.Start("explorer.exe", txtLogPath.Text)
            Else
                MessageBox.Show("โฟลเดอร์ไม่มีอยู่", "ข้อผิดพลาด",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"ไม่สามารถเปิดโฟลเดอร์ได้: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnBackupNow_Click(sender As Object, e As EventArgs) Handles btnBackupNow.Click
        PerformBackupNow()
    End Sub

    Private Sub btnImport_Click(sender As Object, e As EventArgs) Handles btnImport.Click
        ImportSettings()
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        ExportSettings()
    End Sub
#End Region

#Region "Validation & Testing"
    Private Function ValidateSettings() As Boolean
        ' ตรวจสอบรูปแบบ Regex
        Try
            If Not String.IsNullOrEmpty(txtExtractPattern.Text) Then
                Dim regex As New Regex(txtExtractPattern.Text)
            End If
        Catch ex As Exception
            MessageBox.Show($"รูปแบบการดึงข้อมูลไม่ถูกต้อง: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            tabSettings.SelectedTab = tabScanner
            txtExtractPattern.Focus()
            Return False
        End Try

        ' ตรวจสอบพาธฐานข้อมูล Access
        If Not ValidateAccessDatabasePath() Then
            tabSettings.SelectedTab = tabDatabase
            txtDatabase.Focus()
            Return False
        End If

        ' ตรวจสอบโฟลเดอร์
        If chkAutoBackup.Checked AndAlso Not String.IsNullOrEmpty(txtBackupPath.Text) Then
            Try
                Dim parentDir As String = Path.GetDirectoryName(txtBackupPath.Text)
                If Not Directory.Exists(parentDir) Then
                    MessageBox.Show("โฟลเดอร์สำรองข้อมูลไม่ถูกต้อง", "ข้อผิดพลาด",
                                  MessageBoxButtons.OK, MessageBoxIcon.Error)
                    tabSettings.SelectedTab = tabDatabase
                    txtBackupPath.Focus()
                    Return False
                End If
            Catch
                ' ไม่ต้องทำอะไรถ้าตรวจสอบไม่ได้
            End Try
        End If

        Return True
    End Function

    Private Function ValidateAccessDatabasePath() As Boolean
        If String.IsNullOrEmpty(txtDatabase.Text) Then
            MessageBox.Show("กรุณาระบุพาธของฐานข้อมูล Access", "ข้อผิดพลาด",
                         MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        ' ตรวจสอบว่าไฟล์มีนามสกุลที่ถูกต้อง
        Dim extension As String = Path.GetExtension(txtDatabase.Text).ToLower()
        If extension <> ".accdb" AndAlso extension <> ".mdb" Then
            MessageBox.Show("ไฟล์ฐานข้อมูลต้องเป็นนามสกุล .accdb หรือ .mdb เท่านั้น", "ข้อผิดพลาด",
                         MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If

        ' ตรวจสอบว่าไฟล์มีอยู่จริง
        If Not File.Exists(txtDatabase.Text) Then
            Dim result As DialogResult = MessageBox.Show(
                "ไม่พบไฟล์ฐานข้อมูลตามพาธที่ระบุ" & vbNewLine &
                "คุณต้องการใช้พาธนี้ต่อไปหรือไม่? ระบบจะสร้างฐานข้อมูลใหม่เมื่อเริ่มใช้งาน",
                "คำเตือน", MessageBoxButtons.YesNo, MessageBoxIcon.Warning)
            
            If result = DialogResult.No Then
                Return False
            End If
        End If

        Return True
    End Function

    Private Sub TestExtractPattern()
        Try
            Dim pattern As String = txtExtractPattern.Text
            Dim testData As String = "R00C-191604255012766+Q000060+P20414-007700A000+D20250527+LPT0000000+V00C-191604+U0000000"

            If String.IsNullOrEmpty(pattern) Then
                MessageBox.Show("กรุณาใส่รูปแบบการดึงข้อมูล", "ข้อผิดพลาด",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Dim regex As New Regex(pattern)
            Dim match As Match = regex.Match(testData)

            If match.Success AndAlso match.Groups.Count > 1 Then
                MessageBox.Show($"ทดสอบสำเร็จ!" & vbNewLine &
                              $"ผลลัพธ์: {match.Groups(1).Value}",
                              "ผลการทดสอบ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("ไม่พบข้อมูลที่ตรงกับรูปแบบ", "ผลการทดสอบ",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาด: {ex.Message}", "ข้อผิดพลาด",
                          MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RunPatternTest()
        Try
            Dim pattern As String = txtExtractPattern.Text
            Dim testData As String = txtTestInput.Text

            If String.IsNullOrEmpty(pattern) Then
                txtTestResult.Text = "ข้อผิดพลาด: ไม่มีรูปแบบการดึงข้อมูล"
                Return
            End If

            If String.IsNullOrEmpty(testData) Then
                txtTestResult.Text = "ข้อผิดพลาด: ไม่มีข้อมูลทดสอบ"
                Return
            End If

            Dim regex As New Regex(pattern)
            Dim match As Match = regex.Match(testData)

            If match.Success AndAlso match.Groups.Count > 1 Then
                txtTestResult.Text = match.Groups(1).Value
            Else
                txtTestResult.Text = "ไม่พบข้อมูลที่ตรงกับรูปแบบ"
            End If

        Catch ex As Exception
            txtTestResult.Text = $"ข้อผิดพลาด: {ex.Message}"
        End Try
    End Sub

    Private Sub TestDatabaseConnection()
        Try
            Dim connectionString As String = BuildConnectionString()

            Using conn As New OleDbConnection(connectionString)
                conn.Open()
                lblConnectionStatus.Text = "สถานะ: เชื่อมต่อสำเร็จ"
                lblConnectionStatus.ForeColor = Color.Green
                MessageBox.Show("เชื่อมต่อฐานข้อมูลสำเร็จ", "สำเร็จ",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            End Using

        Catch ex As Exception
            lblConnectionStatus.Text = "สถานะ: เชื่อมต่อไม่สำเร็จ"
            lblConnectionStatus.ForeColor = Color.Red
            MessageBox.Show($"ไม่สามารถเชื่อมต่อฐานข้อมูลได้: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Function BuildConnectionString() As String
        Try
            Dim connectionString As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={txtDatabase.Text};"

            ' เพิ่มรหัสผ่านถ้ามี
            If Not String.IsNullOrEmpty(txtPassword.Text) Then
                connectionString += $"Jet OLEDB:Database Password={txtPassword.Text};"
            End If

            Return connectionString
        Catch ex As Exception
            Return ""
        End Try
    End Function
#End Region

#Region "Import/Export & Backup"
    Private Sub ImportSettings()
        Try
            If openFileDialog.ShowDialog() = DialogResult.OK Then
                If File.Exists(openFileDialog.FileName) Then
                    File.Copy(openFileDialog.FileName, CONFIG_FILE, True)
                    LoadSettings()
                    MessageBox.Show("นำเข้าการตั้งค่าเรียบร้อยแล้ว", "สำเร็จ",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการนำเข้า: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportSettings()
        Try
            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                If File.Exists(CONFIG_FILE) Then
                    File.Copy(CONFIG_FILE, saveFileDialog.FileName, True)
                    MessageBox.Show("ส่งออกการตั้งค่าเรียบร้อยแล้ว", "สำเร็จ",
                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                Else
                    MessageBox.Show("ไม่พบไฟล์การตั้งค่า", "ข้อผิดพลาด",
                                  MessageBoxButtons.OK, MessageBoxIcon.Warning)
                End If
            End If
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออก: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub PerformBackupNow()
        Try
            If String.IsNullOrEmpty(txtBackupPath.Text) Then
                MessageBox.Show("กรุณาระบุโฟลเดอร์สำรองข้อมูล", "ข้อผิดพลาด",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            If Not Directory.Exists(txtBackupPath.Text) Then
                Directory.CreateDirectory(txtBackupPath.Text)
            End If

            Dim backupFileName As String = $"Backup_{DateTime.Now:yyyyMMdd_HHmmss}.config"
            Dim backupPath As String = Path.Combine(txtBackupPath.Text, backupFileName)

            If File.Exists(CONFIG_FILE) Then
                File.Copy(CONFIG_FILE, backupPath, True)
                MessageBox.Show($"สำรองข้อมูลเรียบร้อยแล้ว" & vbNewLine &
                              $"ไฟล์: {backupFileName}", "สำเร็จ",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("ไม่พบไฟล์การตั้งค่าที่จะสำรอง", "ข้อผิดพลาด",
                              MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการสำรองข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ResetToDefaults()
        Try
            ' รีเซ็ตค่าทั้งหมดกลับเป็นค่าเริ่มต้น
            numScanTimeout.Value = 100
            chkShowFullData.Checked = False
            chkAutoExtract.Checked = True
            chkSoundEnabled.Checked = True
            txtExtractPattern.Text = "\+P([^+]+)\+D"

            txtDatabase.Text = DEFAULT_ACCESS_PATH
            txtPassword.Text = ""
            
            chkAutoBackup.Checked = False
            txtBackupPath.Text = Path.Combine(System.Windows.Forms.Application.StartupPath, "Backup")
            If cmbBackupInterval.Items.Count > 0 Then cmbBackupInterval.SelectedIndex = 0

            If cmbTheme.Items.Count > 0 Then cmbTheme.SelectedIndex = 0
            If cmbLanguage.Items.Count > 0 Then cmbLanguage.SelectedIndex = 0
            If cmbFontSize.Items.Count > 1 Then cmbFontSize.SelectedIndex = 1
            chkShowStatusBar.Checked = True
            chkShowToolbar.Checked = True

            chkShowNotifications.Checked = True
            chkSoundNotifications.Checked = True
            numNotificationDuration.Value = 5

            chkEnableLogging.Checked = True
            If cmbLogLevel.Items.Count > 2 Then cmbLogLevel.SelectedIndex = 2
            txtLogPath.Text = Path.Combine(System.Windows.Forms.Application.StartupPath, "Logs")

            numMaxRecords.Value = 1000
            chkAutoCleanup.Checked = False
            numCleanupDays.Value = 30

            UpdateControlStates()
            MarkAsChanged()

            MessageBox.Show("รีเซ็ตการตั้งค่าเรียบร้อยแล้ว", "สำเร็จ",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการรีเซ็ต: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Password Encryption/Decryption"
    Private Function EncryptPassword(password As String) As String
        Try
            If String.IsNullOrEmpty(password) Then Return ""

            ' การเข้ารหัสแบบง่าย (ใช้ Base64)
            ' หมายเหตุ: ในการใช้งานจริงควรใช้การเข้ารหัสที่แข็งแกร่งกว่า
            Dim bytes As Byte() = System.Text.Encoding.UTF8.GetBytes(password)
            Return Convert.ToBase64String(bytes)
        Catch
            Return password
        End Try
    End Function

    Private Function DecryptPassword(encryptedPassword As String) As String
        Try
            If String.IsNullOrEmpty(encryptedPassword) Then Return ""

            ' การถอดรหัสแบบง่าย (จาก Base64)
            Dim bytes As Byte() = Convert.FromBase64String(encryptedPassword)
            Return System.Text.Encoding.UTF8.GetString(bytes)
        Catch
            Return encryptedPassword
        End Try
    End Function
#End Region

#Region "Public Methods for External Access"
    ''' <summary>
    ''' ดึงค่าการตั้งค่าสำหรับใช้งานภายนอก
    ''' </summary>
    Public Function GetSetting(key As String) As Object
        Try
            Select Case key.ToLower()
                Case "scantimeout"
                    Return CInt(numScanTimeout.Value)
                Case "showfulldata"
                    Return chkShowFullData.Checked
                Case "autoextract"
                    Return chkAutoExtract.Checked
                Case "soundenabled"
                    Return chkSoundEnabled.Checked
                Case "extractpattern"
                    Return txtExtractPattern.Text
                Case "accessdatabasepath"
                    Return txtDatabase.Text
                Case "accesspassword"
                    Return txtPassword.Text
                Case "theme"
                    Return cmbTheme.Text
                Case "language"
                    Return cmbLanguage.Text
                Case "enablelogging"
                    Return chkEnableLogging.Checked
                Case "loglevel"
                    Return cmbLogLevel.Text
                Case "logpath"
                    Return txtLogPath.Text
                Case "maxrecords"
                    Return CInt(numMaxRecords.Value)
                Case Else
                    Return Nothing
            End Select
        Catch
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' ตรวจสอบว่ามีการเปลี่ยนแปลงหรือไม่
    ''' </summary>
    Public ReadOnly Property HasUnsavedChanges As Boolean
        Get
            Return hasChanges
        End Get
    End Property

    ''' <summary>
    ''' ทดสอบรูปแบบการดึงข้อมูลจากภายนอก
    ''' </summary>
    Public Function TestPatternExtraction(testData As String) As String
        Try
            Dim pattern As String = txtExtractPattern.Text
            If String.IsNullOrEmpty(pattern) Then Return testData

            Dim regex As New Regex(pattern)
            Dim match As Match = regex.Match(testData)

            If match.Success AndAlso match.Groups.Count > 1 Then
                Return match.Groups(1).Value
            End If

            Return testData
        Catch
            Return testData
        End Try
    End Function

    ''' <summary>
    ''' ดึงพาธของฐานข้อมูล Access
    ''' </summary>
    Public Function GetAccessDatabasePath() As String
        Return txtDatabase.Text
    End Function

    ''' <summary>
    ''' ดึงรหัสผ่านของฐานข้อมูล Access
    ''' </summary>
    Public Function GetAccessDatabasePassword() As String
        Return txtPassword.Text
    End Function
#End Region

#Region "Theme and UI Management"
    ''' <summary>
    ''' ปรับธีมของฟอร์ม
    ''' </summary>
    Private Sub ApplyTheme()
        Try
            Select Case cmbTheme.Text.ToLower()
                Case "dark"
                    ApplyDarkTheme()
                Case "light"
                    ApplyLightTheme()
                Case "auto"
                    If IsSystemDarkMode() Then
                        ApplyDarkTheme()
                    Else
                        ApplyLightTheme()
                    End If
            End Select
        Catch
            ' ไม่ต้องทำอะไรถ้าเกิดข้อผิดพลาด
        End Try
    End Sub

    Private Sub ApplyDarkTheme()
        Try
            Me.BackColor = Color.FromArgb(45, 45, 48)
            Me.ForeColor = Color.White

            For Each tab As TabPage In tabSettings.TabPages
                tab.BackColor = Color.FromArgb(37, 37, 38)
                tab.ForeColor = Color.White
            Next
        Catch
            ' ไม่ต้องทำอะไรถ้าเกิดข้อผิดพลาด
        End Try
    End Sub

    Private Sub ApplyLightTheme()
        Try
            Me.BackColor = SystemColors.Control
            Me.ForeColor = SystemColors.ControlText

            For Each tab As TabPage In tabSettings.TabPages
                tab.BackColor = SystemColors.Control
                tab.ForeColor = SystemColors.ControlText
            Next
        Catch
            ' ไม่ต้องทำอะไรถ้าเกิดข้อผิดพลาด
        End Try
    End Sub

    Private Function IsSystemDarkMode() As Boolean
        Try
            Using key As RegistryKey = Registry.CurrentUser.OpenSubKey("SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                If key IsNot Nothing Then
                    Dim value As Object = key.GetValue("AppsUseLightTheme")
                    If value IsNot Nothing Then
                        Return CInt(value) = 0
                    End If
                End If
            End Using
        Catch
        End Try
        Return False
    End Function

    ''' <summary>
    ''' ปรับขนาดตัวอักษร
    ''' </summary>
    Private Sub ApplyFontSize()
        Try
            Dim newSize As Single = 9.0F

            Select Case cmbFontSize.Text
                Case "เล็ก"
                    newSize = 8.0F
                Case "ปกติ"
                    newSize = 9.0F
                Case "ใหญ่"
                    newSize = 11.0F
                Case "ใหญ่พิเศษ"
                    newSize = 13.0F
            End Select

            Me.Font = New Font(Me.Font.FontFamily, newSize)

            ' ปรับขนาดตัวอักษรของควบคุมทั้งหมด
            ApplyFontToControls(Me, newSize)
        Catch
            ' ไม่ต้องทำอะไรถ้าเกิดข้อผิดพลาด
        End Try
    End Sub

    Private Sub ApplyFontToControls(parent As Control, fontSize As Single)
        Try
            For Each ctrl As Control In parent.Controls
                ctrl.Font = New Font(ctrl.Font.FontFamily, fontSize)
                If ctrl.HasChildren Then
                    ApplyFontToControls(ctrl, fontSize)
                End If
            Next
        Catch
            ' ไม่ต้องทำอะไรถ้าเกิดข้อผิดพลาด
        End Try
    End Sub
#End Region

#Region "Logging Management"
    ''' <summary>
    ''' สร้างโฟลเดอร์ Log หากยังไม่มี
    ''' </summary>
    Private Sub EnsureLogDirectoryExists()
        Try
            If Not String.IsNullOrEmpty(txtLogPath.Text) AndAlso Not Directory.Exists(txtLogPath.Text) Then
                Directory.CreateDirectory(txtLogPath.Text)
            End If
        Catch ex As Exception
            MessageBox.Show($"ไม่สามารถสร้างโฟลเดอร์ Log ได้: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
        End Try
    End Sub

    ''' <summary>
    ''' เขียน Log ทดสอบ
    ''' </summary>
    Private Sub WriteTestLog()
        Try
            EnsureLogDirectoryExists()

            Dim logFile As String = Path.Combine(txtLogPath.Text, $"TestLog_{DateTime.Now:yyyyMMdd}.log")
            Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - [TEST] การทดสอบระบบ Log เรียบร้อย{Environment.NewLine}"

            File.AppendAllText(logFile, logEntry)

            MessageBox.Show("เขียน Test Log เรียบร้อยแล้ว", "สำเร็จ",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"ไม่สามารถเขียน Log ได้: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Advanced Features"
    ''' <summary>
    ''' ตรวจสอบและทำความสะอาดข้อมูลเก่า
    ''' </summary>
    Private Sub PerformDataCleanup()
        If Not chkAutoCleanup.Checked Then Return

        Try
            Dim cutoffDate As DateTime = DateTime.Now.AddDays(-CInt(numCleanupDays.Value))

            ' ตรวจสอบและลบไฟล์ Log เก่า
            If Directory.Exists(txtLogPath.Text) Then
                Dim logFiles As String() = Directory.GetFiles(txtLogPath.Text, "*.log")
                Dim deletedCount As Integer = 0

                For Each logFile As String In logFiles
                    Dim fileInfo As New FileInfo(logFile)
                    If fileInfo.CreationTime < cutoffDate Then
                        File.Delete(logFile)
                        deletedCount += 1
                    End If
                Next

                If deletedCount > 0 Then
                    MessageBox.Show($"ลบไฟล์ Log เก่า {deletedCount} ไฟล์เรียบร้อยแล้ว",
                                  "ทำความสะอาดข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการทำความสะอาด: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ตรวจสอบพื้นที่ดิสก์
    ''' </summary>
    Private Sub CheckDiskSpace()
        Try
            Dim drive As New DriveInfo(Path.GetPathRoot(System.Windows.Forms.Application.StartupPath))
            Dim freeSpaceGB As Double = drive.AvailableFreeSpace / (1024 * 1024 * 1024)

            If freeSpaceGB < 1 Then ' เตือนเมื่อพื้นที่ว่างน้อยกว่า 1 GB
                MessageBox.Show($"พื้นที่ดิสก์เหลือน้อย: {freeSpaceGB:F2} GB",
                              "คำเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch
            ' ไม่ต้องแสดงข้อผิดพลาดถ้าตรวจสอบไม่ได้
        End Try
    End Sub
#End Region

#Region "Form Designer Event Wiring"
    ' เพิ่ม Event Handler สำหรับการเปลี่ยนแปลงค่าต่างๆ ที่เหลือ
    Private Sub WireAdditionalEvents()
        Try
            ' เชื่อม Event Handler กับ Control ต่างๆ ที่ยังไม่ได้เชื่อม
            AddHandler chkShowFullData.CheckedChanged, AddressOf GenericChangeHandler
            AddHandler chkAutoExtract.CheckedChanged, AddressOf GenericChangeHandler
            AddHandler chkSoundEnabled.CheckedChanged, AddressOf GenericChangeHandler
            AddHandler txtDatabase.TextChanged, AddressOf GenericChangeHandler
            AddHandler txtPassword.TextChanged, AddressOf GenericChangeHandler
            AddHandler txtBackupPath.TextChanged, AddressOf GenericChangeHandler
            AddHandler cmbBackupInterval.SelectedIndexChanged, AddressOf GenericChangeHandler
            AddHandler cmbLanguage.SelectedIndexChanged, AddressOf GenericChangeHandler
            AddHandler chkShowStatusBar.CheckedChanged, AddressOf GenericChangeHandler
            AddHandler chkShowToolbar.CheckedChanged, AddressOf GenericChangeHandler
            AddHandler chkSoundNotifications.CheckedChanged, AddressOf GenericChangeHandler
            AddHandler numNotificationDuration.ValueChanged, AddressOf GenericChangeHandler
            AddHandler cmbLogLevel.SelectedIndexChanged, AddressOf GenericChangeHandler
            AddHandler txtLogPath.TextChanged, AddressOf GenericChangeHandler
            AddHandler numMaxRecords.ValueChanged, AddressOf GenericChangeHandler
            AddHandler numCleanupDays.ValueChanged, AddressOf GenericChangeHandler
        Catch
            ' ไม่ต้องทำอะไรถ้าไม่สามารถ wire events ได้
        End Try
    End Sub

    Private Sub GenericChangeHandler(sender As Object, e As EventArgs)
        MarkAsChanged()
    End Sub
#End Region

End Class