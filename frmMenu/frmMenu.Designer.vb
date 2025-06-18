<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMenu
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(disposing As Boolean)
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
        components = New ComponentModel.Container()
        pnlHeader = New Panel()
        lblTitle = New Label()
        picScanIcon = New PictureBox()
        pnlMain = New Panel()
        grpBarcodeInfo = New GroupBox()
        lblBarcodeValue = New Label()
        txtBarcode = New TextBox()
        lblBarcode = New Label()
        lblLastScanned = New Label()
        lblScanTime = New Label()
        grpStatus = New GroupBox()
        lblStatusValue = New Label()
        picStatusIcon = New PictureBox()
        lblStatus = New Label()
        pnlButtons = New Panel()
        btnClear = New Button()
        btnSettings = New Button()
        btnExit = New Button()
        btnHistory = New Button()
        btnTest = New Button()
        btnCheckUpdate = New Button()
        statusStrip = New StatusStrip()
        toolStripStatusLabel = New ToolStripStatusLabel()
        toolStripProgressBar = New ToolStripProgressBar()
        timerStatus = New Timer(components)
        pnlHeader.SuspendLayout()
        CType(picScanIcon, ComponentModel.ISupportInitialize).BeginInit()
        pnlMain.SuspendLayout()
        grpBarcodeInfo.SuspendLayout()
        grpStatus.SuspendLayout()
        CType(picStatusIcon, ComponentModel.ISupportInitialize).BeginInit()
        pnlButtons.SuspendLayout()
        statusStrip.SuspendLayout()
        SuspendLayout()
        ' 
        ' pnlHeader
        ' 
        pnlHeader.BackColor = Color.FromArgb(CByte(52), CByte(152), CByte(219))
        pnlHeader.Controls.Add(lblTitle)
        pnlHeader.Controls.Add(picScanIcon)
        pnlHeader.Dock = DockStyle.Top
        pnlHeader.Location = New Point(0, 0)
        pnlHeader.Name = "pnlHeader"
        pnlHeader.Size = New Size(680, 80)
        pnlHeader.TabIndex = 0
        ' 
        ' lblTitle
        ' 
        lblTitle.AutoSize = True
        lblTitle.Font = New Font("Segoe UI", 18F, FontStyle.Bold)
        lblTitle.ForeColor = Color.White
        lblTitle.Location = New Point(80, 25)
        lblTitle.Name = "lblTitle"
        lblTitle.Size = New Size(286, 32)
        lblTitle.TabIndex = 1
        lblTitle.Text = "QR Code Scanner System"
        ' 
        ' picScanIcon
        ' 
        picScanIcon.BackColor = Color.White
        picScanIcon.Location = New Point(20, 20)
        picScanIcon.Name = "picScanIcon"
        picScanIcon.Size = New Size(40, 40)
        picScanIcon.TabIndex = 0
        picScanIcon.TabStop = False
        ' 
        ' pnlMain
        ' 
        pnlMain.BackColor = Color.White
        pnlMain.Controls.Add(grpBarcodeInfo)
        pnlMain.Controls.Add(grpStatus)
        pnlMain.Dock = DockStyle.Fill
        pnlMain.Location = New Point(0, 80)
        pnlMain.Name = "pnlMain"
        pnlMain.Padding = New Padding(20)
        pnlMain.Size = New Size(680, 300)
        pnlMain.TabIndex = 1
        ' 
        ' grpBarcodeInfo
        ' 
        grpBarcodeInfo.Controls.Add(lblBarcodeValue)
        grpBarcodeInfo.Controls.Add(txtBarcode)
        grpBarcodeInfo.Controls.Add(lblBarcode)
        grpBarcodeInfo.Controls.Add(lblLastScanned)
        grpBarcodeInfo.Controls.Add(lblScanTime)
        grpBarcodeInfo.Font = New Font("Segoe UI", 12F, FontStyle.Bold)
        grpBarcodeInfo.Location = New Point(20, 20)
        grpBarcodeInfo.Name = "grpBarcodeInfo"
        grpBarcodeInfo.Size = New Size(640, 130)
        grpBarcodeInfo.TabIndex = 0
        grpBarcodeInfo.TabStop = False
        grpBarcodeInfo.Text = "ข้อมูล QR Code"
        ' 
        ' lblBarcodeValue
        ' 
        lblBarcodeValue.AutoSize = True
        lblBarcodeValue.Font = New Font("Segoe UI", 11F)
        lblBarcodeValue.ForeColor = Color.FromArgb(CByte(46), CByte(125), CByte(50))
        lblBarcodeValue.Location = New Point(20, 85)
        lblBarcodeValue.Name = "lblBarcodeValue"
        lblBarcodeValue.Size = New Size(135, 20)
        lblBarcodeValue.TabIndex = 4
        lblBarcodeValue.Text = "No barcode scanned"
        ' 
        ' txtBarcode
        ' 
        txtBarcode.Font = New Font("Segoe UI", 11F)
        txtBarcode.Location = New Point(120, 35)
        txtBarcode.Name = "txtBarcode"
        txtBarcode.Size = New Size(500, 27)
        txtBarcode.TabIndex = 1
        txtBarcode.ReadOnly = True
        ' 
        ' lblBarcode
        ' 
        lblBarcode.AutoSize = True
        lblBarcode.Font = New Font("Segoe UI", 11F)
        lblBarcode.Location = New Point(20, 38)
        lblBarcode.Name = "lblBarcode"
        lblBarcode.Size = New Size(94, 20)
        lblBarcode.TabIndex = 0
        lblBarcode.Text = "QR Code/รหัส:"
        ' 
        ' lblLastScanned
        ' 
        lblLastScanned.AutoSize = True
        lblLastScanned.Font = New Font("Segoe UI", 10F)
        lblLastScanned.Location = New Point(280, 87)
        lblLastScanned.Name = "lblLastScanned"
        lblLastScanned.Size = New Size(102, 19)
        lblLastScanned.TabIndex = 2
        lblLastScanned.Text = "สแกนล่าสุดเมื่อ:"
        ' 
        ' lblScanTime
        ' 
        lblScanTime.AutoSize = True
        lblScanTime.Font = New Font("Segoe UI", 10F)
        lblScanTime.ForeColor = Color.FromArgb(CByte(255), CByte(159), CByte(67))
        lblScanTime.Location = New Point(390, 87)
        lblScanTime.Name = "lblScanTime"
        lblScanTime.Size = New Size(97, 19)
        lblScanTime.TabIndex = 3
        lblScanTime.Text = "Never scanned"
        ' 
        ' grpStatus
        ' 
        grpStatus.Controls.Add(lblStatusValue)
        grpStatus.Controls.Add(picStatusIcon)
        grpStatus.Controls.Add(lblStatus)
        grpStatus.Font = New Font("Segoe UI", 12F, FontStyle.Bold)
        grpStatus.Location = New Point(20, 170)
        grpStatus.Name = "grpStatus"
        grpStatus.Size = New Size(640, 100)
        grpStatus.TabIndex = 1
        grpStatus.TabStop = False
        grpStatus.Text = "สถานะ"
        ' 
        ' lblStatusValue
        ' 
        lblStatusValue.AutoSize = True
        lblStatusValue.Font = New Font("Segoe UI", 11F, FontStyle.Bold)
        lblStatusValue.ForeColor = Color.FromArgb(CByte(255), CByte(159), CByte(67))
        lblStatusValue.Location = New Point(120, 38)
        lblStatusValue.Name = "lblStatusValue"
        lblStatusValue.Size = New Size(119, 20)
        lblStatusValue.TabIndex = 2
        lblStatusValue.Text = "Ready to scan..."
        ' 
        ' picStatusIcon
        ' 
        picStatusIcon.BackColor = Color.FromArgb(CByte(255), CByte(159), CByte(67))
        picStatusIcon.Location = New Point(80, 35)
        picStatusIcon.Name = "picStatusIcon"
        picStatusIcon.Size = New Size(20, 20)
        picStatusIcon.TabIndex = 1
        picStatusIcon.TabStop = False
        ' 
        ' lblStatus
        ' 
        lblStatus.AutoSize = True
        lblStatus.Font = New Font("Segoe UI", 10F)
        lblStatus.Location = New Point(20, 38)
        lblStatus.Name = "lblStatus"
        lblStatus.Size = New Size(50, 19)
        lblStatus.TabIndex = 0
        lblStatus.Text = "Status:"
        ' 
        ' pnlButtons
        ' 
        pnlButtons.BackColor = Color.White
        pnlButtons.Controls.Add(btnClear)
        pnlButtons.Controls.Add(btnSettings)
        pnlButtons.Controls.Add(btnExit)
        pnlButtons.Controls.Add(btnHistory)
        pnlButtons.Controls.Add(btnTest)
        pnlButtons.Controls.Add(btnCheckUpdate)
        pnlButtons.Dock = DockStyle.Bottom
        pnlButtons.Location = New Point(0, 400)
        pnlButtons.Name = "pnlButtons"
        pnlButtons.Padding = New Padding(20, 10, 20, 10)
        pnlButtons.Size = New Size(650, 60)
        pnlButtons.TabIndex = 2
        ' 
        ' btnClear
        ' 
        btnClear.BackColor = Color.FromArgb(CByte(108), CByte(117), CByte(125))
        btnClear.FlatAppearance.BorderSize = 0
        btnClear.FlatStyle = FlatStyle.Flat
        btnClear.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        btnClear.ForeColor = Color.White
        btnClear.Location = New Point(20, 15)
        btnClear.Name = "btnClear"
        btnClear.Size = New Size(100, 35)
        btnClear.TabIndex = 0
        btnClear.Text = "🗑️ Clear"
        btnClear.UseVisualStyleBackColor = False
        ' 
        ' btnTest
        ' 
        btnTest.BackColor = Color.FromArgb(CByte(52), CByte(152), CByte(219))
        btnTest.FlatAppearance.BorderSize = 0
        btnTest.FlatStyle = FlatStyle.Flat
        btnTest.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        btnTest.ForeColor = Color.White
        btnTest.Location = New Point(130, 15)
        btnTest.Name = "btnTest"
        btnTest.Size = New Size(100, 35)
        btnTest.TabIndex = 4
        btnTest.Text = "🧪 ทดสอบ"
        btnTest.UseVisualStyleBackColor = False
        ' 
        ' btnSettings
        ' 
        btnSettings.BackColor = Color.FromArgb(CByte(52), CByte(152), CByte(219))
        btnSettings.FlatAppearance.BorderSize = 0
        btnSettings.FlatStyle = FlatStyle.Flat
        btnSettings.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        btnSettings.ForeColor = Color.White
        btnSettings.Location = New Point(240, 15)
        btnSettings.Name = "btnSettings"
        btnSettings.Size = New Size(100, 35)
        btnSettings.TabIndex = 1
        btnSettings.Text = "⚙️ Settings"
        btnSettings.UseVisualStyleBackColor = False
        ' 
        ' btnCheckUpdate
        ' 
        btnCheckUpdate.BackColor = Color.FromArgb(CByte(155), CByte(89), CByte(182))
        btnCheckUpdate.FlatAppearance.BorderSize = 0
        btnCheckUpdate.FlatStyle = FlatStyle.Flat
        btnCheckUpdate.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        btnCheckUpdate.ForeColor = Color.White
        btnCheckUpdate.Location = New Point(350, 15)
        btnCheckUpdate.Name = "btnCheckUpdate"
        btnCheckUpdate.Size = New Size(100, 35)
        btnCheckUpdate.TabIndex = 5
        btnCheckUpdate.Text = "🔄 Update"
        btnCheckUpdate.UseVisualStyleBackColor = False
        ' 
        ' btnHistory
        ' 
        btnHistory.BackColor = Color.FromArgb(CByte(52), CByte(152), CByte(219))
        btnHistory.FlatAppearance.BorderSize = 0
        btnHistory.FlatStyle = FlatStyle.Flat
        btnHistory.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        btnHistory.ForeColor = Color.White
        btnHistory.Location = New Point(460, 15)
        btnHistory.Name = "btnHistory"
        btnHistory.Size = New Size(100, 35)
        btnHistory.TabIndex = 3
        btnHistory.Text = "📋 ประวัติ"
        btnHistory.UseVisualStyleBackColor = False
        ' 
        ' btnExit
        ' 
        btnExit.BackColor = Color.FromArgb(CByte(231), CByte(76), CByte(60))
        btnExit.FlatAppearance.BorderSize = 0
        btnExit.FlatStyle = FlatStyle.Flat
        btnExit.Font = New Font("Segoe UI", 10F, FontStyle.Bold)
        btnExit.ForeColor = Color.White
        btnExit.Location = New Point(570, 15)
        btnExit.Name = "btnExit"
        btnExit.Size = New Size(100, 35)
        btnExit.TabIndex = 2
        btnExit.Text = "❌ Exit"
        btnExit.UseVisualStyleBackColor = False
        ' 
        ' statusStrip
        ' 
        statusStrip.BackColor = Color.FromArgb(CByte(236), CByte(240), CByte(241))
        statusStrip.Items.AddRange(New ToolStripItem() {toolStripStatusLabel, toolStripProgressBar})
        statusStrip.Location = New Point(0, 460)
        statusStrip.Name = "statusStrip"
        statusStrip.Size = New Size(680, 22)
        statusStrip.TabIndex = 3
        statusStrip.Text = "StatusStrip1"
        ' 
        ' toolStripStatusLabel
        ' 
        toolStripStatusLabel.Name = "toolStripStatusLabel"
        toolStripStatusLabel.Size = New Size(585, 17)
        toolStripStatusLabel.Spring = True
        toolStripStatusLabel.Text = "Ready"
        toolStripStatusLabel.TextAlign = ContentAlignment.MiddleLeft
        ' 
        ' toolStripProgressBar
        ' 
        toolStripProgressBar.Name = "toolStripProgressBar"
        toolStripProgressBar.Size = New Size(100, 16)
        toolStripProgressBar.Visible = False
        ' 
        ' timerStatus
        ' 
        timerStatus.Interval = 1000
        ' 
        ' frmMenu
        ' 
        AutoScaleDimensions = New SizeF(7F, 15F)
        AutoScaleMode = AutoScaleMode.Font
        BackColor = Color.White
        ClientSize = New Size(680, 482)
        Controls.Add(pnlMain)
        Controls.Add(pnlButtons)
        Controls.Add(pnlHeader)
        Controls.Add(statusStrip)
        Font = New Font("Segoe UI", 9F)
        FormBorderStyle = FormBorderStyle.FixedSingle
        MaximizeBox = False
        Name = "frmMenu"
        StartPosition = FormStartPosition.CenterScreen
        Text = "QR Code Scanner System"
        pnlHeader.ResumeLayout(False)
        pnlHeader.PerformLayout()
        CType(picScanIcon, ComponentModel.ISupportInitialize).EndInit()
        pnlMain.ResumeLayout(False)
        grpBarcodeInfo.ResumeLayout(False)
        grpBarcodeInfo.PerformLayout()
        grpStatus.ResumeLayout(False)
        grpStatus.PerformLayout()
        CType(picStatusIcon, ComponentModel.ISupportInitialize).BeginInit()
        pnlButtons.ResumeLayout(False)
        statusStrip.ResumeLayout(False)
        statusStrip.PerformLayout()
        ResumeLayout(False)
        PerformLayout()

    End Sub

    Friend WithEvents pnlHeader As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents picScanIcon As PictureBox
    Friend WithEvents pnlMain As Panel
    Friend WithEvents grpBarcodeInfo As GroupBox
    Friend WithEvents lblBarcodeValue As Label
    Friend WithEvents txtBarcode As TextBox
    Friend WithEvents lblBarcode As Label
    Friend WithEvents lblLastScanned As Label
    Friend WithEvents lblScanTime As Label
    Friend WithEvents grpStatus As GroupBox
    Friend WithEvents lblStatusValue As Label
    Friend WithEvents picStatusIcon As PictureBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents pnlButtons As Panel
    Friend WithEvents btnClear As Button
    Friend WithEvents btnSettings As Button
    Friend WithEvents btnExit As Button
    Friend WithEvents statusStrip As StatusStrip
    Friend WithEvents toolStripStatusLabel As ToolStripStatusLabel
    Friend WithEvents toolStripProgressBar As ToolStripProgressBar
    Friend WithEvents timerStatus As Timer
    Friend WithEvents btnHistory As Button
    Friend WithEvents btnTest As Button
    Friend WithEvents btnCheckUpdate As Button

End Class