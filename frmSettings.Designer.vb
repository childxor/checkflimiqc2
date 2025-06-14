<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmSettings
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmSettings))
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.tabSettings = New System.Windows.Forms.TabControl()
        Me.tabScanner = New System.Windows.Forms.TabPage()
        Me.grpScannerSettings = New System.Windows.Forms.GroupBox()
        Me.lblScanTimeout = New System.Windows.Forms.Label()
        Me.numScanTimeout = New System.Windows.Forms.NumericUpDown()
        Me.lblMs = New System.Windows.Forms.Label()
        Me.chkShowFullData = New System.Windows.Forms.CheckBox()
        Me.chkAutoExtract = New System.Windows.Forms.CheckBox()
        Me.chkSoundEnabled = New System.Windows.Forms.CheckBox()
        Me.lblExtractPattern = New System.Windows.Forms.Label()
        Me.txtExtractPattern = New System.Windows.Forms.TextBox()
        Me.btnTestPattern = New System.Windows.Forms.Button()
        Me.grpTestArea = New System.Windows.Forms.GroupBox()
        Me.lblTestInput = New System.Windows.Forms.Label()
        Me.txtTestInput = New System.Windows.Forms.TextBox()
        Me.lblTestResult = New System.Windows.Forms.Label()
        Me.txtTestResult = New System.Windows.Forms.TextBox()
        Me.btnRunTest = New System.Windows.Forms.Button()
        Me.tabDatabase = New System.Windows.Forms.TabPage()
        Me.grpDatabaseSettings = New System.Windows.Forms.GroupBox()
        Me.lblDatabase = New System.Windows.Forms.Label()
        Me.txtDatabase = New System.Windows.Forms.TextBox()
        Me.lblPassword = New System.Windows.Forms.Label()
        Me.txtPassword = New System.Windows.Forms.TextBox()
        Me.btnBrowseDatabase = New System.Windows.Forms.Button()
        Me.btnTestConnection = New System.Windows.Forms.Button()
        Me.lblConnectionStatus = New System.Windows.Forms.Label()
        Me.grpBackup = New System.Windows.Forms.GroupBox()
        Me.chkAutoBackup = New System.Windows.Forms.CheckBox()
        Me.lblBackupPath = New System.Windows.Forms.Label()
        Me.txtBackupPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseBackup = New System.Windows.Forms.Button()
        Me.lblBackupInterval = New System.Windows.Forms.Label()
        Me.cmbBackupInterval = New System.Windows.Forms.ComboBox()
        Me.btnBackupNow = New System.Windows.Forms.Button()
        Me.tabDisplay = New System.Windows.Forms.TabPage()
        Me.grpAppearance = New System.Windows.Forms.GroupBox()
        Me.lblTheme = New System.Windows.Forms.Label()
        Me.cmbTheme = New System.Windows.Forms.ComboBox()
        Me.lblLanguage = New System.Windows.Forms.Label()
        Me.cmbLanguage = New System.Windows.Forms.ComboBox()
        Me.lblFontSize = New System.Windows.Forms.Label()
        Me.cmbFontSize = New System.Windows.Forms.ComboBox()
        Me.chkShowStatusBar = New System.Windows.Forms.CheckBox()
        Me.chkShowToolbar = New System.Windows.Forms.CheckBox()
        Me.grpNotifications = New System.Windows.Forms.GroupBox()
        Me.chkShowNotifications = New System.Windows.Forms.CheckBox()
        Me.chkSoundNotifications = New System.Windows.Forms.CheckBox()
        Me.lblNotificationDuration = New System.Windows.Forms.Label()
        Me.numNotificationDuration = New System.Windows.Forms.NumericUpDown()
        Me.lblSeconds = New System.Windows.Forms.Label()
        Me.tabAdvanced = New System.Windows.Forms.TabPage()
        Me.grpLogging = New System.Windows.Forms.GroupBox()
        Me.chkEnableLogging = New System.Windows.Forms.CheckBox()
        Me.lblLogLevel = New System.Windows.Forms.Label()
        Me.cmbLogLevel = New System.Windows.Forms.ComboBox()
        Me.lblLogPath = New System.Windows.Forms.Label()
        Me.txtLogPath = New System.Windows.Forms.TextBox()
        Me.btnBrowseLog = New System.Windows.Forms.Button()
        Me.btnOpenLogFolder = New System.Windows.Forms.Button()
        Me.grpPerformance = New System.Windows.Forms.GroupBox()
        Me.lblMaxRecords = New System.Windows.Forms.Label()
        Me.numMaxRecords = New System.Windows.Forms.NumericUpDown()
        Me.chkAutoCleanup = New System.Windows.Forms.CheckBox()
        Me.lblCleanupDays = New System.Windows.Forms.Label()
        Me.numCleanupDays = New System.Windows.Forms.NumericUpDown()
        Me.lblDays = New System.Windows.Forms.Label()
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.btnOK = New System.Windows.Forms.Button()
        Me.btnCancel = New System.Windows.Forms.Button()
        Me.btnApply = New System.Windows.Forms.Button()
        Me.btnReset = New System.Windows.Forms.Button()
        Me.btnImport = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.folderBrowserDialog = New System.Windows.Forms.FolderBrowserDialog()
        Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.toolTip = New System.Windows.Forms.ToolTip(Me.components)
        Me.pnlMain.SuspendLayout()
        Me.tabSettings.SuspendLayout()
        Me.tabScanner.SuspendLayout()
        Me.grpScannerSettings.SuspendLayout()
        CType(Me.numScanTimeout, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.grpTestArea.SuspendLayout()
        Me.tabDatabase.SuspendLayout()
        Me.grpDatabaseSettings.SuspendLayout()
        Me.grpBackup.SuspendLayout()
        Me.tabDisplay.SuspendLayout()
        Me.grpAppearance.SuspendLayout()
        Me.grpNotifications.SuspendLayout()
        CType(Me.numNotificationDuration, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.tabAdvanced.SuspendLayout()
        Me.grpLogging.SuspendLayout()
        Me.grpPerformance.SuspendLayout()
        CType(Me.numMaxRecords, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.numCleanupDays, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlButtons.SuspendLayout()
        Me.SuspendLayout()
        '
        'pnlMain
        '
        Me.pnlMain.Controls.Add(Me.tabSettings)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 0)
        Me.pnlMain.Margin = New System.Windows.Forms.Padding(4)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Padding = New System.Windows.Forms.Padding(10)
        Me.pnlMain.Size = New System.Drawing.Size(900, 650)
        Me.pnlMain.TabIndex = 0
        '
        'tabSettings
        '
        Me.tabSettings.Controls.Add(Me.tabScanner)
        Me.tabSettings.Controls.Add(Me.tabDatabase)
        Me.tabSettings.Controls.Add(Me.tabDisplay)
        Me.tabSettings.Controls.Add(Me.tabAdvanced)
        Me.tabSettings.Dock = System.Windows.Forms.DockStyle.Fill
        Me.tabSettings.Location = New System.Drawing.Point(10, 10)
        Me.tabSettings.Margin = New System.Windows.Forms.Padding(4)
        Me.tabSettings.Name = "tabSettings"
        Me.tabSettings.SelectedIndex = 0
        Me.tabSettings.Size = New System.Drawing.Size(880, 630)
        Me.tabSettings.TabIndex = 0
        '
        'tabScanner
        '
        Me.tabScanner.Controls.Add(Me.grpScannerSettings)
        Me.tabScanner.Controls.Add(Me.grpTestArea)
        Me.tabScanner.Location = New System.Drawing.Point(4, 28)
        Me.tabScanner.Margin = New System.Windows.Forms.Padding(4)
        Me.tabScanner.Name = "tabScanner"
        Me.tabScanner.Padding = New System.Windows.Forms.Padding(4)
        Me.tabScanner.Size = New System.Drawing.Size(872, 598)
        Me.tabScanner.TabIndex = 0
        Me.tabScanner.Text = "การตั้งค่าเครื่องสแกน"
        Me.tabScanner.UseVisualStyleBackColor = True
        '
        'grpScannerSettings
        '
        Me.grpScannerSettings.Controls.Add(Me.lblScanTimeout)
        Me.grpScannerSettings.Controls.Add(Me.numScanTimeout)
        Me.grpScannerSettings.Controls.Add(Me.lblMs)
        Me.grpScannerSettings.Controls.Add(Me.chkShowFullData)
        Me.grpScannerSettings.Controls.Add(Me.chkAutoExtract)
        Me.grpScannerSettings.Controls.Add(Me.chkSoundEnabled)
        Me.grpScannerSettings.Controls.Add(Me.lblExtractPattern)
        Me.grpScannerSettings.Controls.Add(Me.txtExtractPattern)
        Me.grpScannerSettings.Controls.Add(Me.btnTestPattern)
        Me.grpScannerSettings.Location = New System.Drawing.Point(8, 8)
        Me.grpScannerSettings.Margin = New System.Windows.Forms.Padding(4)
        Me.grpScannerSettings.Name = "grpScannerSettings"
        Me.grpScannerSettings.Padding = New System.Windows.Forms.Padding(4)
        Me.grpScannerSettings.Size = New System.Drawing.Size(856, 280)
        Me.grpScannerSettings.TabIndex = 0
        Me.grpScannerSettings.TabStop = False
        Me.grpScannerSettings.Text = "ตั้งค่าการสแกน"
        '
        'lblScanTimeout
        '
        Me.lblScanTimeout.AutoSize = True
        Me.lblScanTimeout.Location = New System.Drawing.Point(20, 40)
        Me.lblScanTimeout.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblScanTimeout.Name = "lblScanTimeout"
        Me.lblScanTimeout.Size = New System.Drawing.Size(152, 19)
        Me.lblScanTimeout.TabIndex = 0
        Me.lblScanTimeout.Text = "ระยะเวลารอการสแกน:"
        '
        'numScanTimeout
        '
        Me.numScanTimeout.Location = New System.Drawing.Point(180, 38)
        Me.numScanTimeout.Margin = New System.Windows.Forms.Padding(4)
        Me.numScanTimeout.Maximum = New Decimal(New Integer() {1000, 0, 0, 0})
        Me.numScanTimeout.Minimum = New Decimal(New Integer() {50, 0, 0, 0})
        Me.numScanTimeout.Name = "numScanTimeout"
        Me.numScanTimeout.Size = New System.Drawing.Size(100, 27)
        Me.numScanTimeout.TabIndex = 1
        Me.numScanTimeout.Value = New Decimal(New Integer() {100, 0, 0, 0})
        '
        'lblMs
        '
        Me.lblMs.AutoSize = True
        Me.lblMs.Location = New System.Drawing.Point(288, 42)
        Me.lblMs.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMs.Name = "lblMs"
        Me.lblMs.Size = New System.Drawing.Size(83, 19)
        Me.lblMs.TabIndex = 2
        Me.lblMs.Text = "มิลลิวินาที"
        '
        'chkShowFullData
        '
        Me.chkShowFullData.AutoSize = True
        Me.chkShowFullData.Location = New System.Drawing.Point(24, 80)
        Me.chkShowFullData.Margin = New System.Windows.Forms.Padding(4)
        Me.chkShowFullData.Name = "chkShowFullData"
        Me.chkShowFullData.Size = New System.Drawing.Size(239, 23)
        Me.chkShowFullData.TabIndex = 3
        Me.chkShowFullData.Text = "แสดงข้อมูลทั้งหมดใน MessageBox"
        Me.chkShowFullData.UseVisualStyleBackColor = True
        '
        'chkAutoExtract
        '
        Me.chkAutoExtract.AutoSize = True
        Me.chkAutoExtract.Checked = True
        Me.chkAutoExtract.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkAutoExtract.Location = New System.Drawing.Point(24, 110)
        Me.chkAutoExtract.Margin = New System.Windows.Forms.Padding(4)
        Me.chkAutoExtract.Name = "chkAutoExtract"
        Me.chkAutoExtract.Size = New System.Drawing.Size(197, 23)
        Me.chkAutoExtract.TabIndex = 4
        Me.chkAutoExtract.Text = "ดึงข้อมูลโดยอัตโนมัติ"
        Me.chkAutoExtract.UseVisualStyleBackColor = True
        '
        'chkSoundEnabled
        '
        Me.chkSoundEnabled.AutoSize = True
        Me.chkSoundEnabled.Checked = True
        Me.chkSoundEnabled.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSoundEnabled.Location = New System.Drawing.Point(24, 140)
        Me.chkSoundEnabled.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSoundEnabled.Name = "chkSoundEnabled"
        Me.chkSoundEnabled.Size = New System.Drawing.Size(166, 23)
        Me.chkSoundEnabled.TabIndex = 5
        Me.chkSoundEnabled.Text = "เปิดเสียงเมื่อสแกนสำเร็จ"
        Me.chkSoundEnabled.UseVisualStyleBackColor = True
        '
        'lblExtractPattern
        '
        Me.lblExtractPattern.AutoSize = True
        Me.lblExtractPattern.Location = New System.Drawing.Point(20, 180)
        Me.lblExtractPattern.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblExtractPattern.Name = "lblExtractPattern"
        Me.lblExtractPattern.Size = New System.Drawing.Size(152, 19)
        Me.lblExtractPattern.TabIndex = 6
        Me.lblExtractPattern.Text = "รูปแบบการดึงข้อมูล:"
        '
        'txtExtractPattern
        '
        Me.txtExtractPattern.Location = New System.Drawing.Point(180, 176)
        Me.txtExtractPattern.Margin = New System.Windows.Forms.Padding(4)
        Me.txtExtractPattern.Name = "txtExtractPattern"
        Me.txtExtractPattern.Size = New System.Drawing.Size(500, 27)
        Me.txtExtractPattern.TabIndex = 7
        Me.txtExtractPattern.Text = "\+P([^+]+)\+D"
        '
        'btnTestPattern
        '
        Me.btnTestPattern.Location = New System.Drawing.Point(688, 174)
        Me.btnTestPattern.Margin = New System.Windows.Forms.Padding(4)
        Me.btnTestPattern.Name = "btnTestPattern"
        Me.btnTestPattern.Size = New System.Drawing.Size(100, 32)
        Me.btnTestPattern.TabIndex = 8
        Me.btnTestPattern.Text = "ทดสอบ"
        Me.btnTestPattern.UseVisualStyleBackColor = True
        '
        'grpTestArea
        '
        Me.grpTestArea.Controls.Add(Me.lblTestInput)
        Me.grpTestArea.Controls.Add(Me.txtTestInput)
        Me.grpTestArea.Controls.Add(Me.lblTestResult)
        Me.grpTestArea.Controls.Add(Me.txtTestResult)
        Me.grpTestArea.Controls.Add(Me.btnRunTest)
        Me.grpTestArea.Location = New System.Drawing.Point(8, 300)
        Me.grpTestArea.Margin = New System.Windows.Forms.Padding(4)
        Me.grpTestArea.Name = "grpTestArea"
        Me.grpTestArea.Padding = New System.Windows.Forms.Padding(4)
        Me.grpTestArea.Size = New System.Drawing.Size(856, 200)
        Me.grpTestArea.TabIndex = 1
        Me.grpTestArea.TabStop = False
        Me.grpTestArea.Text = "พื้นที่ทดสอบ"
        '
        'lblTestInput
        '
        Me.lblTestInput.AutoSize = True
        Me.lblTestInput.Location = New System.Drawing.Point(20, 40)
        Me.lblTestInput.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTestInput.Name = "lblTestInput"
        Me.lblTestInput.Size = New System.Drawing.Size(105, 19)
        Me.lblTestInput.TabIndex = 0
        Me.lblTestInput.Text = "ข้อมูลทดสอบ:"
        '
        'txtTestInput
        '
        Me.txtTestInput.Location = New System.Drawing.Point(140, 36)
        Me.txtTestInput.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTestInput.Multiline = True
        Me.txtTestInput.Name = "txtTestInput"
        Me.txtTestInput.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtTestInput.Size = New System.Drawing.Size(600, 50)
        Me.txtTestInput.TabIndex = 1
        Me.txtTestInput.Text = "R00C-191604255012766+Q000060+P20414-007700A000+D20250527+LPT0000000+V00C-191604+U0000000"
        '
        'lblTestResult
        '
        Me.lblTestResult.AutoSize = True
        Me.lblTestResult.Location = New System.Drawing.Point(20, 100)
        Me.lblTestResult.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTestResult.Name = "lblTestResult"
        Me.lblTestResult.Size = New System.Drawing.Size(68, 19)
        Me.lblTestResult.TabIndex = 2
        Me.lblTestResult.Text = "ผลลัพธ์:"
        '
        'txtTestResult
        '
        Me.txtTestResult.BackColor = System.Drawing.SystemColors.Control
        Me.txtTestResult.Location = New System.Drawing.Point(140, 96)
        Me.txtTestResult.Margin = New System.Windows.Forms.Padding(4)
        Me.txtTestResult.Multiline = True
        Me.txtTestResult.Name = "txtTestResult"
        Me.txtTestResult.ReadOnly = True
        Me.txtTestResult.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal
        Me.txtTestResult.Size = New System.Drawing.Size(600, 50)
        Me.txtTestResult.TabIndex = 3
        '
        'btnRunTest
        '
        Me.btnRunTest.Location = New System.Drawing.Point(750, 36)
        Me.btnRunTest.Margin = New System.Windows.Forms.Padding(4)
        Me.btnRunTest.Name = "btnRunTest"
        Me.btnRunTest.Size = New System.Drawing.Size(80, 32)
        Me.btnRunTest.TabIndex = 4
        Me.btnRunTest.Text = "ทดสอบ"
        Me.btnRunTest.UseVisualStyleBackColor = True
        '
        'tabDatabase
        '
        Me.tabDatabase.Controls.Add(Me.grpDatabaseSettings)
        Me.tabDatabase.Controls.Add(Me.grpBackup)
        Me.tabDatabase.Location = New System.Drawing.Point(4, 28)
        Me.tabDatabase.Margin = New System.Windows.Forms.Padding(4)
        Me.tabDatabase.Name = "tabDatabase"
        Me.tabDatabase.Padding = New System.Windows.Forms.Padding(4)
        Me.tabDatabase.Size = New System.Drawing.Size(872, 598)
        Me.tabDatabase.TabIndex = 1
        Me.tabDatabase.Text = "ฐานข้อมูล"
        Me.tabDatabase.UseVisualStyleBackColor = True
        '
        'grpDatabaseSettings
        '
        Me.grpDatabaseSettings.Controls.Add(Me.lblDatabase)
        Me.grpDatabaseSettings.Controls.Add(Me.txtDatabase)
        Me.grpDatabaseSettings.Controls.Add(Me.lblPassword)
        Me.grpDatabaseSettings.Controls.Add(Me.txtPassword)
        Me.grpDatabaseSettings.Controls.Add(Me.btnBrowseDatabase)
        Me.grpDatabaseSettings.Controls.Add(Me.btnTestConnection)
        Me.grpDatabaseSettings.Controls.Add(Me.lblConnectionStatus)
        Me.grpDatabaseSettings.Location = New System.Drawing.Point(8, 8)
        Me.grpDatabaseSettings.Margin = New System.Windows.Forms.Padding(4)
        Me.grpDatabaseSettings.Name = "grpDatabaseSettings"
        Me.grpDatabaseSettings.Padding = New System.Windows.Forms.Padding(4)
        Me.grpDatabaseSettings.Size = New System.Drawing.Size(856, 280)
        Me.grpDatabaseSettings.TabIndex = 0
        Me.grpDatabaseSettings.TabStop = False
        Me.grpDatabaseSettings.Text = "การตั้งค่าฐานข้อมูล Access"
        '
        'lblDatabase
        '
        Me.lblDatabase.AutoSize = True
        Me.lblDatabase.Location = New System.Drawing.Point(20, 40)
        Me.lblDatabase.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDatabase.Name = "lblDatabase"
        Me.lblDatabase.Size = New System.Drawing.Size(94, 19)
        Me.lblDatabase.TabIndex = 0
        Me.lblDatabase.Text = "ไฟล์ฐานข้อมูล:"
        '
        'txtDatabase
        '
        Me.txtDatabase.Location = New System.Drawing.Point(120, 36)
        Me.txtDatabase.Margin = New System.Windows.Forms.Padding(4)
        Me.txtDatabase.Name = "txtDatabase"
        Me.txtDatabase.Size = New System.Drawing.Size(500, 27)
        Me.txtDatabase.TabIndex = 1
        '
        'lblPassword
        '
        Me.lblPassword.AutoSize = True
        Me.lblPassword.Location = New System.Drawing.Point(20, 80)
        Me.lblPassword.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblPassword.Name = "lblPassword"
        Me.lblPassword.Size = New System.Drawing.Size(65, 19)
        Me.lblPassword.TabIndex = 2
        Me.lblPassword.Text = "รหัสผ่าน:"
        '
        'txtPassword
        '
        Me.txtPassword.Location = New System.Drawing.Point(120, 76)
        Me.txtPassword.Margin = New System.Windows.Forms.Padding(4)
        Me.txtPassword.Name = "txtPassword"
        Me.txtPassword.PasswordChar = Global.Microsoft.VisualBasic.ChrW(42)
        Me.txtPassword.Size = New System.Drawing.Size(300, 27)
        Me.txtPassword.TabIndex = 3
        '
        'btnBrowseDatabase
        '
        Me.btnBrowseDatabase.Location = New System.Drawing.Point(630, 36)
        Me.btnBrowseDatabase.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBrowseDatabase.Name = "btnBrowseDatabase"
        Me.btnBrowseDatabase.Size = New System.Drawing.Size(80, 32)
        Me.btnBrowseDatabase.TabIndex = 4
        Me.btnBrowseDatabase.Text = "เลือก..."
        Me.btnBrowseDatabase.UseVisualStyleBackColor = True
        '
        'btnTestConnection
        '
        Me.btnTestConnection.Location = New System.Drawing.Point(450, 76)
        Me.btnTestConnection.Margin = New System.Windows.Forms.Padding(4)
        Me.btnTestConnection.Name = "btnTestConnection"
        Me.btnTestConnection.Size = New System.Drawing.Size(120, 32)
        Me.btnTestConnection.TabIndex = 5
        Me.btnTestConnection.Text = "ทดสอบการเชื่อมต่อ"
        Me.btnTestConnection.UseVisualStyleBackColor = True
        '
        'lblConnectionStatus
        '
        Me.lblConnectionStatus.AutoSize = True
        Me.lblConnectionStatus.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.lblConnectionStatus.ForeColor = System.Drawing.Color.Red
        Me.lblConnectionStatus.Location = New System.Drawing.Point(120, 120)
        Me.lblConnectionStatus.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblConnectionStatus.Name = "lblConnectionStatus"
        Me.lblConnectionStatus.Size = New System.Drawing.Size(143, 20)
        Me.lblConnectionStatus.TabIndex = 6
        Me.lblConnectionStatus.Text = "สถานะ: ไม่ได้เชื่อมต่อ"
        '
        'grpBackup
        '
        Me.grpBackup.Controls.Add(Me.chkAutoBackup)
        Me.grpBackup.Controls.Add(Me.lblBackupPath)
        Me.grpBackup.Controls.Add(Me.txtBackupPath)
        Me.grpBackup.Controls.Add(Me.btnBrowseBackup)
        Me.grpBackup.Controls.Add(Me.lblBackupInterval)
        Me.grpBackup.Controls.Add(Me.cmbBackupInterval)
        Me.grpBackup.Controls.Add(Me.btnBackupNow)
        Me.grpBackup.Location = New System.Drawing.Point(8, 300)
        Me.grpBackup.Margin = New System.Windows.Forms.Padding(4)
        Me.grpBackup.Name = "grpBackup"
        Me.grpBackup.Padding = New System.Windows.Forms.Padding(4)
        Me.grpBackup.Size = New System.Drawing.Size(856, 200)
        Me.grpBackup.TabIndex = 1
        Me.grpBackup.TabStop = False
        Me.grpBackup.Text = "การสำรองข้อมูล"
        '
        'chkAutoBackup
        '
        Me.chkAutoBackup.AutoSize = True
        Me.chkAutoBackup.Location = New System.Drawing.Point(24, 40)
        Me.chkAutoBackup.Margin = New System.Windows.Forms.Padding(4)
        Me.chkAutoBackup.Name = "chkAutoBackup"
        Me.chkAutoBackup.Size = New System.Drawing.Size(189, 23)
        Me.chkAutoBackup.TabIndex = 0
        Me.chkAutoBackup.Text = "เปิดการสำรองอัตโนมัติ"
        Me.chkAutoBackup.UseVisualStyleBackColor = True
        '
        'lblBackupPath
        '
        Me.lblBackupPath.AutoSize = True
        Me.lblBackupPath.Location = New System.Drawing.Point(20, 80)
        Me.lblBackupPath.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBackupPath.Name = "lblBackupPath"
        Me.lblBackupPath.Size = New System.Drawing.Size(141, 19)
        Me.lblBackupPath.TabIndex = 1
        Me.lblBackupPath.Text = "โฟลเดอร์สำรองข้อมูล:"
        '
        'txtBackupPath
        '
        Me.txtBackupPath.Location = New System.Drawing.Point(170, 76)
        Me.txtBackupPath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtBackupPath.Name = "txtBackupPath"
        Me.txtBackupPath.Size = New System.Drawing.Size(500, 27)
        Me.txtBackupPath.TabIndex = 2
        '
        'btnBrowseBackup
        '
        Me.btnBrowseBackup.Location = New System.Drawing.Point(680, 74)
        Me.btnBrowseBackup.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBrowseBackup.Name = "btnBrowseBackup"
        Me.btnBrowseBackup.Size = New System.Drawing.Size(80, 32)
        Me.btnBrowseBackup.TabIndex = 3
        Me.btnBrowseBackup.Text = "เลือก..."
        Me.btnBrowseBackup.UseVisualStyleBackColor = True
        '
        'lblBackupInterval
        '
        Me.lblBackupInterval.AutoSize = True
        Me.lblBackupInterval.Location = New System.Drawing.Point(20, 120)
        Me.lblBackupInterval.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblBackupInterval.Name = "lblBackupInterval"
        Me.lblBackupInterval.Size = New System.Drawing.Size(141, 19)
        Me.lblBackupInterval.TabIndex = 4
        Me.lblBackupInterval.Text = "ความถี่ในการสำรอง:"
        '
        'cmbBackupInterval
        '
        Me.cmbBackupInterval.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbBackupInterval.FormattingEnabled = True
        Me.cmbBackupInterval.Items.AddRange(New Object() {"ทุกวัน", "ทุกสัปดาห์", "ทุกเดือน"})
        Me.cmbBackupInterval.Location = New System.Drawing.Point(170, 116)
        Me.cmbBackupInterval.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbBackupInterval.Name = "cmbBackupInterval"
        Me.cmbBackupInterval.Size = New System.Drawing.Size(200, 27)
        Me.cmbBackupInterval.TabIndex = 5
        '
        'btnBackupNow
        '
        Me.btnBackupNow.Location = New System.Drawing.Point(400, 114)
        Me.btnBackupNow.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBackupNow.Name = "btnBackupNow"
        Me.btnBackupNow.Size = New System.Drawing.Size(120, 32)
        Me.btnBackupNow.TabIndex = 6
        Me.btnBackupNow.Text = "สำรองข้อมูลทันที"
        Me.btnBackupNow.UseVisualStyleBackColor = True
        '
        'tabDisplay
        '
        Me.tabDisplay.Controls.Add(Me.grpAppearance)
        Me.tabDisplay.Controls.Add(Me.grpNotifications)
        Me.tabDisplay.Location = New System.Drawing.Point(4, 28)
        Me.tabDisplay.Margin = New System.Windows.Forms.Padding(4)
        Me.tabDisplay.Name = "tabDisplay"
        Me.tabDisplay.Size = New System.Drawing.Size(872, 598)
        Me.tabDisplay.TabIndex = 2
        Me.tabDisplay.Text = "การแสดงผล"
        Me.tabDisplay.UseVisualStyleBackColor = True
        '
        'grpAppearance
        '
        Me.grpAppearance.Controls.Add(Me.lblTheme)
        Me.grpAppearance.Controls.Add(Me.cmbTheme)
        Me.grpAppearance.Controls.Add(Me.lblLanguage)
        Me.grpAppearance.Controls.Add(Me.cmbLanguage)
        Me.grpAppearance.Controls.Add(Me.lblFontSize)
        Me.grpAppearance.Controls.Add(Me.cmbFontSize)
        Me.grpAppearance.Controls.Add(Me.chkShowStatusBar)
        Me.grpAppearance.Controls.Add(Me.chkShowToolbar)
        Me.grpAppearance.Location = New System.Drawing.Point(8, 8)
        Me.grpAppearance.Margin = New System.Windows.Forms.Padding(4)
        Me.grpAppearance.Name = "grpAppearance"
        Me.grpAppearance.Padding = New System.Windows.Forms.Padding(4)
        Me.grpAppearance.Size = New System.Drawing.Size(856, 280)
        Me.grpAppearance.TabIndex = 0
        Me.grpAppearance.TabStop = False
        Me.grpAppearance.Text = "รูปแบบการแสดงผล"
        '
        'lblTheme
        '
        Me.lblTheme.AutoSize = True
        Me.lblTheme.Location = New System.Drawing.Point(20, 40)
        Me.lblTheme.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblTheme.Name = "lblTheme"
        Me.lblTheme.Size = New System.Drawing.Size(46, 19)
        Me.lblTheme.TabIndex = 0
        Me.lblTheme.Text = "ธีม:"
        '
        'cmbTheme
        '
        Me.cmbTheme.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbTheme.FormattingEnabled = True
        Me.cmbTheme.Items.AddRange(New Object() {"Light", "Dark", "Auto"})
        Me.cmbTheme.Location = New System.Drawing.Point(120, 36)
        Me.cmbTheme.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbTheme.Name = "cmbTheme"
        Me.cmbTheme.Size = New System.Drawing.Size(200, 27)
        Me.cmbTheme.TabIndex = 1
        '
        'lblLanguage
        '
        Me.lblLanguage.AutoSize = True
        Me.lblLanguage.Location = New System.Drawing.Point(20, 80)
        Me.lblLanguage.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblLanguage.Name = "lblLanguage"
        Me.lblLanguage.Size = New System.Drawing.Size(56, 19)
        Me.lblLanguage.TabIndex = 2
        Me.lblLanguage.Text = "ภาษา:"
        '
        'cmbLanguage
        '
        Me.cmbLanguage.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLanguage.FormattingEnabled = True
        Me.cmbLanguage.Items.AddRange(New Object() {"ไทย", "English"})
        Me.cmbLanguage.Location = New System.Drawing.Point(120, 76)
        Me.cmbLanguage.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbLanguage.Name = "cmbLanguage"
        Me.cmbLanguage.Size = New System.Drawing.Size(200, 27)
        Me.cmbLanguage.TabIndex = 3
        '
        'lblFontSize
        '
        Me.lblFontSize.AutoSize = True
        Me.lblFontSize.Location = New System.Drawing.Point(20, 120)
        Me.lblFontSize.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblFontSize.Name = "lblFontSize"
        Me.lblFontSize.Size = New System.Drawing.Size(94, 19)
        Me.lblFontSize.TabIndex = 4
        Me.lblFontSize.Text = "ขนาดตัวอักษร:"
        '
        'cmbFontSize
        '
        Me.cmbFontSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbFontSize.FormattingEnabled = True
        Me.cmbFontSize.Items.AddRange(New Object() {"เล็ก", "ปกติ", "ใหญ่", "ใหญ่พิเศษ"})
        Me.cmbFontSize.Location = New System.Drawing.Point(120, 116)
        Me.cmbFontSize.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbFontSize.Name = "cmbFontSize"
        Me.cmbFontSize.Size = New System.Drawing.Size(200, 27)
        Me.cmbFontSize.TabIndex = 5
        '
        'chkShowStatusBar
        '
        Me.chkShowStatusBar.AutoSize = True
        Me.chkShowStatusBar.Checked = True
        Me.chkShowStatusBar.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowStatusBar.Location = New System.Drawing.Point(24, 160)
        Me.chkShowStatusBar.Margin = New System.Windows.Forms.Padding(4)
        Me.chkShowStatusBar.Name = "chkShowStatusBar"
        Me.chkShowStatusBar.Size = New System.Drawing.Size(127, 23)
        Me.chkShowStatusBar.TabIndex = 6
        Me.chkShowStatusBar.Text = "แสดง Status Bar"
        Me.chkShowStatusBar.UseVisualStyleBackColor = True
        '
        'chkShowToolbar
        '
        Me.chkShowToolbar.AutoSize = True
        Me.chkShowToolbar.Checked = True
        Me.chkShowToolbar.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowToolbar.Location = New System.Drawing.Point(24, 190)
        Me.chkShowToolbar.Margin = New System.Windows.Forms.Padding(4)
        Me.chkShowToolbar.Name = "chkShowToolbar"
        Me.chkShowToolbar.Size = New System.Drawing.Size(115, 23)
        Me.chkShowToolbar.TabIndex = 7
        Me.chkShowToolbar.Text = "แสดง Toolbar"
        Me.chkShowToolbar.UseVisualStyleBackColor = True
        '
        'grpNotifications
        '
        Me.grpNotifications.Controls.Add(Me.chkShowNotifications)
        Me.grpNotifications.Controls.Add(Me.chkSoundNotifications)
        Me.grpNotifications.Controls.Add(Me.lblNotificationDuration)
        Me.grpNotifications.Controls.Add(Me.numNotificationDuration)
        Me.grpNotifications.Controls.Add(Me.lblSeconds)
        Me.grpNotifications.Location = New System.Drawing.Point(8, 300)
        Me.grpNotifications.Margin = New System.Windows.Forms.Padding(4)
        Me.grpNotifications.Name = "grpNotifications"
        Me.grpNotifications.Padding = New System.Windows.Forms.Padding(4)
        Me.grpNotifications.Size = New System.Drawing.Size(856, 150)
        Me.grpNotifications.TabIndex = 1
        Me.grpNotifications.TabStop = False
        Me.grpNotifications.Text = "การแจ้งเตือน"
        '
        'chkShowNotifications
        '
        Me.chkShowNotifications.AutoSize = True
        Me.chkShowNotifications.Checked = True
        Me.chkShowNotifications.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkShowNotifications.Location = New System.Drawing.Point(24, 40)
        Me.chkShowNotifications.Margin = New System.Windows.Forms.Padding(4)
        Me.chkShowNotifications.Name = "chkShowNotifications"
        Me.chkShowNotifications.Size = New System.Drawing.Size(135, 23)
        Me.chkShowNotifications.TabIndex = 0
        Me.chkShowNotifications.Text = "แสดงการแจ้งเตือน"
        Me.chkShowNotifications.UseVisualStyleBackColor = True
        '
        'chkSoundNotifications
        '
        Me.chkSoundNotifications.AutoSize = True
        Me.chkSoundNotifications.Checked = True
        Me.chkSoundNotifications.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkSoundNotifications.Location = New System.Drawing.Point(24, 70)
        Me.chkSoundNotifications.Margin = New System.Windows.Forms.Padding(4)
        Me.chkSoundNotifications.Name = "chkSoundNotifications"
        Me.chkSoundNotifications.Size = New System.Drawing.Size(166, 23)
        Me.chkSoundNotifications.TabIndex = 1
        Me.chkSoundNotifications.Text = "เปิดเสียงการแจ้งเตือน"
        Me.chkSoundNotifications.UseVisualStyleBackColor = True
        '
        'lblNotificationDuration
        '
        Me.lblNotificationDuration.AutoSize = True
        Me.lblNotificationDuration.Location = New System.Drawing.Point(20, 105)
        Me.lblNotificationDuration.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblNotificationDuration.Name = "lblNotificationDuration"
        Me.lblNotificationDuration.Size = New System.Drawing.Size(168, 19)
        Me.lblNotificationDuration.TabIndex = 2
        Me.lblNotificationDuration.Text = "ระยะเวลาแสดงการแจ้งเตือน:"
        '
        'numNotificationDuration
        '
        Me.numNotificationDuration.Location = New System.Drawing.Point(200, 103)
        Me.numNotificationDuration.Margin = New System.Windows.Forms.Padding(4)
        Me.numNotificationDuration.Maximum = New Decimal(New Integer() {30, 0, 0, 0})
        Me.numNotificationDuration.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numNotificationDuration.Name = "numNotificationDuration"
        Me.numNotificationDuration.Size = New System.Drawing.Size(80, 27)
        Me.numNotificationDuration.TabIndex = 3
        Me.numNotificationDuration.Value = New Decimal(New Integer() {5, 0, 0, 0})
        '
        'lblSeconds
        '
        Me.lblSeconds.AutoSize = True
        Me.lblSeconds.Location = New System.Drawing.Point(288, 107)
        Me.lblSeconds.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblSeconds.Name = "lblSeconds"
        Me.lblSeconds.Size = New System.Drawing.Size(47, 19)
        Me.lblSeconds.TabIndex = 4
        Me.lblSeconds.Text = "วินาที"
        '
        'tabAdvanced
        '
        Me.tabAdvanced.Controls.Add(Me.grpLogging)
        Me.tabAdvanced.Controls.Add(Me.grpPerformance)
        Me.tabAdvanced.Location = New System.Drawing.Point(4, 28)
        Me.tabAdvanced.Margin = New System.Windows.Forms.Padding(4)
        Me.tabAdvanced.Name = "tabAdvanced"
        Me.tabAdvanced.Size = New System.Drawing.Size(872, 598)
        Me.tabAdvanced.TabIndex = 3
        Me.tabAdvanced.Text = "ขั้นสูง"
        Me.tabAdvanced.UseVisualStyleBackColor = True
        '
        'grpLogging
        '
        Me.grpLogging.Controls.Add(Me.chkEnableLogging)
        Me.grpLogging.Controls.Add(Me.lblLogLevel)
        Me.grpLogging.Controls.Add(Me.cmbLogLevel)
        Me.grpLogging.Controls.Add(Me.lblLogPath)
        Me.grpLogging.Controls.Add(Me.txtLogPath)
        Me.grpLogging.Controls.Add(Me.btnBrowseLog)
        Me.grpLogging.Controls.Add(Me.btnOpenLogFolder)
        Me.grpLogging.Location = New System.Drawing.Point(8, 8)
        Me.grpLogging.Margin = New System.Windows.Forms.Padding(4)
        Me.grpLogging.Name = "grpLogging"
        Me.grpLogging.Padding = New System.Windows.Forms.Padding(4)
        Me.grpLogging.Size = New System.Drawing.Size(856, 200)
        Me.grpLogging.TabIndex = 0
        Me.grpLogging.TabStop = False
        Me.grpLogging.Text = "การบันทึก Log"
        '
        'chkEnableLogging
        '
        Me.chkEnableLogging.AutoSize = True
        Me.chkEnableLogging.Checked = True
        Me.chkEnableLogging.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkEnableLogging.Location = New System.Drawing.Point(24, 40)
        Me.chkEnableLogging.Margin = New System.Windows.Forms.Padding(4)
        Me.chkEnableLogging.Name = "chkEnableLogging"
        Me.chkEnableLogging.Size = New System.Drawing.Size(149, 23)
        Me.chkEnableLogging.TabIndex = 0
        Me.chkEnableLogging.Text = "เปิดการบันทึก Log"
        Me.chkEnableLogging.UseVisualStyleBackColor = True
        '
        'lblLogLevel
        '
        Me.lblLogLevel.AutoSize = True
        Me.lblLogLevel.Location = New System.Drawing.Point(20, 80)
        Me.lblLogLevel.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblLogLevel.Name = "lblLogLevel"
        Me.lblLogLevel.Size = New System.Drawing.Size(78, 19)
        Me.lblLogLevel.TabIndex = 1
        Me.lblLogLevel.Text = "ระดับ Log:"
        '
        'cmbLogLevel
        '
        Me.cmbLogLevel.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbLogLevel.FormattingEnabled = True
        Me.cmbLogLevel.Items.AddRange(New Object() {"Error", "Warning", "Info", "Debug"})
        Me.cmbLogLevel.Location = New System.Drawing.Point(120, 76)
        Me.cmbLogLevel.Margin = New System.Windows.Forms.Padding(4)
        Me.cmbLogLevel.Name = "cmbLogLevel"
        Me.cmbLogLevel.Size = New System.Drawing.Size(150, 27)
        Me.cmbLogLevel.TabIndex = 2
        '
        'lblLogPath
        '
        Me.lblLogPath.AutoSize = True
        Me.lblLogPath.Location = New System.Drawing.Point(20, 120)
        Me.lblLogPath.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblLogPath.Name = "lblLogPath"
        Me.lblLogPath.Size = New System.Drawing.Size(94, 19)
        Me.lblLogPath.TabIndex = 3
        Me.lblLogPath.Text = "โฟลเดอร์ Log:"
        '
        'txtLogPath
        '
        Me.txtLogPath.Location = New System.Drawing.Point(120, 116)
        Me.txtLogPath.Margin = New System.Windows.Forms.Padding(4)
        Me.txtLogPath.Name = "txtLogPath"
        Me.txtLogPath.Size = New System.Drawing.Size(500, 27)
        Me.txtLogPath.TabIndex = 4
        '
        'btnBrowseLog
        '
        Me.btnBrowseLog.Location = New System.Drawing.Point(630, 114)
        Me.btnBrowseLog.Margin = New System.Windows.Forms.Padding(4)
        Me.btnBrowseLog.Name = "btnBrowseLog"
        Me.btnBrowseLog.Size = New System.Drawing.Size(80, 32)
        Me.btnBrowseLog.TabIndex = 5
        Me.btnBrowseLog.Text = "เลือก..."
        Me.btnBrowseLog.UseVisualStyleBackColor = True
        '
        'btnOpenLogFolder
        '
        Me.btnOpenLogFolder.Location = New System.Drawing.Point(720, 114)
        Me.btnOpenLogFolder.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOpenLogFolder.Name = "btnOpenLogFolder"
        Me.btnOpenLogFolder.Size = New System.Drawing.Size(100, 32)
        Me.btnOpenLogFolder.TabIndex = 6
        Me.btnOpenLogFolder.Text = "เปิดโฟลเดอร์"
        Me.btnOpenLogFolder.UseVisualStyleBackColor = True
        '
        'grpPerformance
        '
        Me.grpPerformance.Controls.Add(Me.lblMaxRecords)
        Me.grpPerformance.Controls.Add(Me.numMaxRecords)
        Me.grpPerformance.Controls.Add(Me.chkAutoCleanup)
        Me.grpPerformance.Controls.Add(Me.lblCleanupDays)
        Me.grpPerformance.Controls.Add(Me.numCleanupDays)
        Me.grpPerformance.Controls.Add(Me.lblDays)
        Me.grpPerformance.Location = New System.Drawing.Point(8, 220)
        Me.grpPerformance.Margin = New System.Windows.Forms.Padding(4)
        Me.grpPerformance.Name = "grpPerformance"
        Me.grpPerformance.Padding = New System.Windows.Forms.Padding(4)
        Me.grpPerformance.Size = New System.Drawing.Size(856, 150)
        Me.grpPerformance.TabIndex = 1
        Me.grpPerformance.TabStop = False
        Me.grpPerformance.Text = "ประสิทธิภาพ"
        '
        'lblMaxRecords
        '
        Me.lblMaxRecords.AutoSize = True
        Me.lblMaxRecords.Location = New System.Drawing.Point(20, 40)
        Me.lblMaxRecords.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblMaxRecords.Name = "lblMaxRecords"
        Me.lblMaxRecords.Size = New System.Drawing.Size(169, 19)
        Me.lblMaxRecords.TabIndex = 0
        Me.lblMaxRecords.Text = "จำนวนข้อมูลสูงสุดในหน้า:"
        '
        'numMaxRecords
        '
        Me.numMaxRecords.Location = New System.Drawing.Point(200, 38)
        Me.numMaxRecords.Margin = New System.Windows.Forms.Padding(4)
        Me.numMaxRecords.Maximum = New Decimal(New Integer() {10000, 0, 0, 0})
        Me.numMaxRecords.Minimum = New Decimal(New Integer() {10, 0, 0, 0})
        Me.numMaxRecords.Name = "numMaxRecords"
        Me.numMaxRecords.Size = New System.Drawing.Size(100, 27)
        Me.numMaxRecords.TabIndex = 1
        Me.numMaxRecords.Value = New Decimal(New Integer() {1000, 0, 0, 0})
        '
        'chkAutoCleanup
        '
        Me.chkAutoCleanup.AutoSize = True
        Me.chkAutoCleanup.Location = New System.Drawing.Point(24, 80)
        Me.chkAutoCleanup.Margin = New System.Windows.Forms.Padding(4)
        Me.chkAutoCleanup.Name = "chkAutoCleanup"
        Me.chkAutoCleanup.Size = New System.Drawing.Size(179, 23)
        Me.chkAutoCleanup.TabIndex = 2
        Me.chkAutoCleanup.Text = "ลบข้อมูลเก่าอัตโนมัติ"
        Me.chkAutoCleanup.UseVisualStyleBackColor = True
        '
        'lblCleanupDays
        '
        Me.lblCleanupDays.AutoSize = True
        Me.lblCleanupDays.Location = New System.Drawing.Point(250, 82)
        Me.lblCleanupDays.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblCleanupDays.Name = "lblCleanupDays"
        Me.lblCleanupDays.Size = New System.Drawing.Size(98, 19)
        Me.lblCleanupDays.TabIndex = 3
        Me.lblCleanupDays.Text = "ข้อมูลเก่ากว่า:"
        '
        'numCleanupDays
        '
        Me.numCleanupDays.Location = New System.Drawing.Point(360, 80)
        Me.numCleanupDays.Margin = New System.Windows.Forms.Padding(4)
        Me.numCleanupDays.Maximum = New Decimal(New Integer() {365, 0, 0, 0})
        Me.numCleanupDays.Minimum = New Decimal(New Integer() {1, 0, 0, 0})
        Me.numCleanupDays.Name = "numCleanupDays"
        Me.numCleanupDays.Size = New System.Drawing.Size(80, 27)
        Me.numCleanupDays.TabIndex = 4
        Me.numCleanupDays.Value = New Decimal(New Integer() {30, 0, 0, 0})
        '
        'lblDays
        '
        Me.lblDays.AutoSize = True
        Me.lblDays.Location = New System.Drawing.Point(448, 84)
        Me.lblDays.Margin = New System.Windows.Forms.Padding(4, 0, 4, 0)
        Me.lblDays.Name = "lblDays"
        Me.lblDays.Size = New System.Drawing.Size(27, 19)
        Me.lblDays.TabIndex = 5
        Me.lblDays.Text = "วัน"
        ' 
        'pnlButtons
        '
        Me.pnlButtons.Controls.Add(Me.btnOK)
        Me.pnlButtons.Controls.Add(Me.btnCancel)
        Me.pnlButtons.Controls.Add(Me.btnApply)
        Me.pnlButtons.Controls.Add(Me.btnReset)
        Me.pnlButtons.Controls.Add(Me.btnImport)
        Me.pnlButtons.Controls.Add(Me.btnExport)
        Me.pnlButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlButtons.Location = New System.Drawing.Point(0, 650)
        Me.pnlButtons.Margin = New System.Windows.Forms.Padding(4)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Padding = New System.Windows.Forms.Padding(10)
        Me.pnlButtons.Size = New System.Drawing.Size(900, 60)
        Me.pnlButtons.TabIndex = 1
        '
        'btnOK
        '
        Me.btnOK.Location = New System.Drawing.Point(500, 15)
        Me.btnOK.Margin = New System.Windows.Forms.Padding(4)
        Me.btnOK.Name = "btnOK"
        Me.btnOK.Size = New System.Drawing.Size(80, 32)
        Me.btnOK.TabIndex = 0
        Me.btnOK.Text = "ตกลง"
        Me.btnOK.UseVisualStyleBackColor = True
        '
        'btnCancel
        '
        Me.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnCancel.Location = New System.Drawing.Point(590, 15)
        Me.btnCancel.Margin = New System.Windows.Forms.Padding(4)
        Me.btnCancel.Name = "btnCancel"
        Me.btnCancel.Size = New System.Drawing.Size(80, 32)
        Me.btnCancel.TabIndex = 1
        Me.btnCancel.Text = "ยกเลิก"
        Me.btnCancel.UseVisualStyleBackColor = True
        '
        'btnApply
        '
        Me.btnApply.Location = New System.Drawing.Point(680, 15)
        Me.btnApply.Margin = New System.Windows.Forms.Padding(4)
        Me.btnApply.Name = "btnApply"
        Me.btnApply.Size = New System.Drawing.Size(80, 32)
        Me.btnApply.TabIndex = 2
        Me.btnApply.Text = "นำไปใช้"
        Me.btnApply.UseVisualStyleBackColor = True
        '
        'btnReset
        '
        Me.btnReset.Location = New System.Drawing.Point(770, 15)
        Me.btnReset.Margin = New System.Windows.Forms.Padding(4)
        Me.btnReset.Name = "btnReset"
        Me.btnReset.Size = New System.Drawing.Size(80, 32)
        Me.btnReset.TabIndex = 3
        Me.btnReset.Text = "รีเซ็ต"
        Me.btnReset.UseVisualStyleBackColor = True
        '
        'btnImport
        '
        Me.btnImport.Location = New System.Drawing.Point(15, 15)
        Me.btnImport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnImport.Name = "btnImport"
        Me.btnImport.Size = New System.Drawing.Size(100, 32)
        Me.btnImport.TabIndex = 4
        Me.btnImport.Text = "นำเข้าการตั้งค่า"
        Me.btnImport.UseVisualStyleBackColor = True
        '
        'btnExport
        '
        Me.btnExport.Location = New System.Drawing.Point(125, 15)
        Me.btnExport.Margin = New System.Windows.Forms.Padding(4)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(100, 32)
        Me.btnExport.TabIndex = 5
        Me.btnExport.Text = "ส่งออกการตั้งค่า"
        Me.btnExport.UseVisualStyleBackColor = True
        '
        'folderBrowserDialog
        '
        Me.folderBrowserDialog.Description = "เลือกโฟลเดอร์"
        '
        'openFileDialog
        '
        Me.openFileDialog.Filter = "Config Files (*.config)|*.config|All Files (*.*)|*.*"
        Me.openFileDialog.Title = "เลือกไฟล์การตั้งค่า"
        '
        'saveFileDialog
        '
        Me.saveFileDialog.Filter = "Config Files (*.config)|*.config|All Files (*.*)|*.*"
        Me.saveFileDialog.Title = "บันทึกไฟล์การตั้งค่า"
        '
        'frmSettings
        '
        Me.AcceptButton = Me.btnOK
        Me.AutoScaleDimensions = New System.Drawing.SizeF(8.0!, 19.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.CancelButton = Me.btnCancel
        Me.ClientSize = New System.Drawing.Size(900, 710)
        Me.Controls.Add(Me.pnlMain)
        Me.Controls.Add(Me.pnlButtons)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frmSettings"
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "การตั้งค่าระบบ - QR Code Scanner"
        Me.pnlMain.ResumeLayout(False)
        Me.tabSettings.ResumeLayout(False)
        Me.tabScanner.ResumeLayout(False)
        Me.grpScannerSettings.ResumeLayout(False)
        Me.grpScannerSettings.PerformLayout()
        CType(Me.numScanTimeout, System.ComponentModel.ISupportInitialize).EndInit()
        Me.grpTestArea.ResumeLayout(False)
        Me.grpTestArea.PerformLayout()
        Me.tabDatabase.ResumeLayout(False)
        Me.grpDatabaseSettings.ResumeLayout(False)
        Me.grpDatabaseSettings.PerformLayout()
        Me.grpBackup.ResumeLayout(False)
        Me.grpBackup.PerformLayout()
        Me.tabDisplay.ResumeLayout(False)
        Me.grpAppearance.ResumeLayout(False)
        Me.grpAppearance.PerformLayout()
        Me.grpNotifications.ResumeLayout(False)
        Me.grpNotifications.PerformLayout()
        CType(Me.numNotificationDuration, System.ComponentModel.ISupportInitialize).EndInit()
        Me.tabAdvanced.ResumeLayout(False)
        Me.grpLogging.ResumeLayout(False)
        Me.grpLogging.PerformLayout()
        Me.grpPerformance.ResumeLayout(False)
        Me.grpPerformance.PerformLayout()
        CType(Me.numMaxRecords, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.numCleanupDays, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlButtons.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    ' Control declarations
    Friend WithEvents pnlMain As Panel
    Friend WithEvents tabSettings As TabControl
    Friend WithEvents tabScanner As TabPage
    Friend WithEvents grpScannerSettings As GroupBox
    Friend WithEvents lblScanTimeout As Label
    Friend WithEvents numScanTimeout As NumericUpDown
    Friend WithEvents lblMs As Label
    Friend WithEvents chkShowFullData As CheckBox
    Friend WithEvents chkAutoExtract As CheckBox
    Friend WithEvents chkSoundEnabled As CheckBox
    Friend WithEvents lblExtractPattern As Label
    Friend WithEvents txtExtractPattern As TextBox
    Friend WithEvents btnTestPattern As Button
    Friend WithEvents grpTestArea As GroupBox
    Friend WithEvents lblTestInput As Label
    Friend WithEvents txtTestInput As TextBox
    Friend WithEvents lblTestResult As Label
    Friend WithEvents txtTestResult As TextBox
    Friend WithEvents btnRunTest As Button
    Friend WithEvents tabDatabase As TabPage
    Friend WithEvents grpDatabaseSettings As GroupBox
    Friend WithEvents lblDatabase As Label
    Friend WithEvents txtDatabase As TextBox
    Friend WithEvents lblPassword As Label
    Friend WithEvents txtPassword As TextBox
    Friend WithEvents btnBrowseDatabase As Button
    Friend WithEvents btnTestConnection As Button
    Friend WithEvents lblConnectionStatus As Label
    Friend WithEvents grpBackup As GroupBox
    Friend WithEvents chkAutoBackup As CheckBox
    Friend WithEvents lblBackupPath As Label
    Friend WithEvents txtBackupPath As TextBox
    Friend WithEvents btnBrowseBackup As Button
    Friend WithEvents lblBackupInterval As Label
    Friend WithEvents cmbBackupInterval As ComboBox
    Friend WithEvents btnBackupNow As Button
    Friend WithEvents tabDisplay As TabPage
    Friend WithEvents grpAppearance As GroupBox
    Friend WithEvents lblTheme As Label
    Friend WithEvents cmbTheme As ComboBox
    Friend WithEvents lblLanguage As Label
    Friend WithEvents cmbLanguage As ComboBox
    Friend WithEvents lblFontSize As Label
    Friend WithEvents cmbFontSize As ComboBox
    Friend WithEvents chkShowStatusBar As CheckBox
    Friend WithEvents chkShowToolbar As CheckBox
    Friend WithEvents grpNotifications As GroupBox
    Friend WithEvents chkShowNotifications As CheckBox
    Friend WithEvents chkSoundNotifications As CheckBox
    Friend WithEvents lblNotificationDuration As Label
    Friend WithEvents numNotificationDuration As NumericUpDown
    Friend WithEvents lblSeconds As Label
    Friend WithEvents tabAdvanced As TabPage
    Friend WithEvents grpLogging As GroupBox
    Friend WithEvents chkEnableLogging As CheckBox
    Friend WithEvents lblLogLevel As Label
    Friend WithEvents cmbLogLevel As ComboBox
    Friend WithEvents lblLogPath As Label
    Friend WithEvents txtLogPath As TextBox
    Friend WithEvents btnBrowseLog As Button
    Friend WithEvents btnOpenLogFolder As Button
    Friend WithEvents grpPerformance As GroupBox
    Friend WithEvents lblMaxRecords As Label
    Friend WithEvents numMaxRecords As NumericUpDown
    Friend WithEvents chkAutoCleanup As CheckBox
    Friend WithEvents lblCleanupDays As Label
    Friend WithEvents numCleanupDays As NumericUpDown
    Friend WithEvents lblDays As Label
    Friend WithEvents pnlButtons As Panel
    Friend WithEvents btnOK As Button
    Friend WithEvents btnCancel As Button
    Friend WithEvents btnApply As Button
    Friend WithEvents btnReset As Button
    Friend WithEvents btnImport As Button
    Friend WithEvents btnExport As Button
    Friend WithEvents folderBrowserDialog As FolderBrowserDialog
    Friend WithEvents openFileDialog As OpenFileDialog
    Friend WithEvents saveFileDialog As SaveFileDialog
    Friend WithEvents toolTip As ToolTip

End Class