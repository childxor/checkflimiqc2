<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmHistory
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
        Me.pnlHeader = New System.Windows.Forms.Panel()
        Me.lblTitle = New System.Windows.Forms.Label()
        Me.picIcon = New System.Windows.Forms.PictureBox()
        Me.pnlFilter = New System.Windows.Forms.Panel()
        Me.grpFilter = New System.Windows.Forms.GroupBox()
        Me.lblSearch = New System.Windows.Forms.Label()
        Me.txtSearch = New System.Windows.Forms.TextBox()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.cmbStatus = New System.Windows.Forms.ComboBox()
        Me.lblFromDate = New System.Windows.Forms.Label()
        Me.dtpFromDate = New System.Windows.Forms.DateTimePicker()
        Me.lblToDate = New System.Windows.Forms.Label()
        Me.dtpToDate = New System.Windows.Forms.DateTimePicker()
        Me.btnRefresh = New System.Windows.Forms.Button()
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.dgvHistory = New System.Windows.Forms.DataGridView()
        Me.pnlButtons = New System.Windows.Forms.Panel()
        Me.btnViewDetail = New System.Windows.Forms.Button()
        Me.btnDelete = New System.Windows.Forms.Button()
        Me.btnExport = New System.Windows.Forms.Button()
        Me.btnExportExcel = New System.Windows.Forms.Button()
        Me.btnClose = New System.Windows.Forms.Button()
        Me.statusStrip = New System.Windows.Forms.StatusStrip()
        Me.lblCount = New System.Windows.Forms.ToolStripStatusLabel()
        Me.toolStripProgressBar = New System.Windows.Forms.ToolStripProgressBar()
        Me.saveFileDialog = New System.Windows.Forms.SaveFileDialog()
        Me.openFileDialog = New System.Windows.Forms.OpenFileDialog()
        Me.btnSettings = New System.Windows.Forms.Button()

        Me.pnlHeader.SuspendLayout()
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlFilter.SuspendLayout()
        Me.grpFilter.SuspendLayout()
        Me.pnlMain.SuspendLayout()
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.pnlButtons.SuspendLayout()
        Me.statusStrip.SuspendLayout()
        Me.SuspendLayout()

        '
        'pnlHeader
        '
        Me.pnlHeader.BackColor = System.Drawing.Color.FromArgb(41, 128, 185)
        Me.pnlHeader.Controls.Add(Me.lblTitle)
        Me.pnlHeader.Controls.Add(Me.picIcon)
        Me.pnlHeader.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlHeader.Location = New System.Drawing.Point(0, 0)
        Me.pnlHeader.Name = "pnlHeader"
        Me.pnlHeader.Size = New System.Drawing.Size(1200, 60)
        Me.pnlHeader.TabIndex = 0

        '
        'lblTitle
        '
        Me.lblTitle.AutoSize = True
        Me.lblTitle.Font = New System.Drawing.Font("Segoe UI", 16.0!, System.Drawing.FontStyle.Bold)
        Me.lblTitle.ForeColor = System.Drawing.Color.White
        Me.lblTitle.Location = New System.Drawing.Point(60, 18)
        Me.lblTitle.Name = "lblTitle"
        Me.lblTitle.Size = New System.Drawing.Size(298, 30)
        Me.lblTitle.TabIndex = 1
        Me.lblTitle.Text = "ประวัติการสแกน QR Code"

        '
        'picIcon
        '
        Me.picIcon.BackColor = System.Drawing.Color.White
        Me.picIcon.Location = New System.Drawing.Point(15, 15)
        Me.picIcon.Name = "picIcon"
        Me.picIcon.Size = New System.Drawing.Size(30, 30)
        Me.picIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picIcon.TabIndex = 0
        Me.picIcon.TabStop = False

        '
        'pnlFilter
        '
        Me.pnlFilter.BackColor = System.Drawing.Color.FromArgb(248, 249, 250)
        Me.pnlFilter.Controls.Add(Me.grpFilter)
        Me.pnlFilter.Dock = System.Windows.Forms.DockStyle.Top
        Me.pnlFilter.Location = New System.Drawing.Point(0, 60)
        Me.pnlFilter.Name = "pnlFilter"
        Me.pnlFilter.Padding = New System.Windows.Forms.Padding(10)
        Me.pnlFilter.Size = New System.Drawing.Size(1200, 100)
        Me.pnlFilter.TabIndex = 1

        '
        'grpFilter
        '
        Me.grpFilter.Controls.Add(Me.lblSearch)
        Me.grpFilter.Controls.Add(Me.txtSearch)
        Me.grpFilter.Controls.Add(Me.lblStatus)
        Me.grpFilter.Controls.Add(Me.cmbStatus)
        Me.grpFilter.Controls.Add(Me.lblFromDate)
        Me.grpFilter.Controls.Add(Me.dtpFromDate)
        Me.grpFilter.Controls.Add(Me.lblToDate)
        Me.grpFilter.Controls.Add(Me.dtpToDate)
        Me.grpFilter.Controls.Add(Me.btnRefresh)
        Me.grpFilter.Dock = System.Windows.Forms.DockStyle.Fill
        Me.grpFilter.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.grpFilter.ForeColor = System.Drawing.Color.FromArgb(52, 73, 94)
        Me.grpFilter.Location = New System.Drawing.Point(10, 10)
        Me.grpFilter.Name = "grpFilter"
        Me.grpFilter.Size = New System.Drawing.Size(1180, 80)
        Me.grpFilter.TabIndex = 0
        Me.grpFilter.TabStop = False
        Me.grpFilter.Text = "กรองข้อมูล"

        '
        'lblSearch
        '
        Me.lblSearch.AutoSize = True
        Me.lblSearch.Location = New System.Drawing.Point(15, 30)
        Me.lblSearch.Name = "lblSearch"
        Me.lblSearch.Size = New System.Drawing.Size(32, 15)
        Me.lblSearch.TabIndex = 0
        Me.lblSearch.Text = "ค้นหา:"

        '
        'txtSearch
        '
        Me.txtSearch.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.txtSearch.Location = New System.Drawing.Point(15, 48)
        Me.txtSearch.Name = "txtSearch"
        Me.txtSearch.Size = New System.Drawing.Size(200, 23)
        Me.txtSearch.TabIndex = 1

        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Location = New System.Drawing.Point(235, 30)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Size = New System.Drawing.Size(39, 15)
        Me.lblStatus.TabIndex = 2
        Me.lblStatus.Text = "สถานะ:"

        '
        'cmbStatus
        '
        Me.cmbStatus.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbStatus.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.cmbStatus.FormattingEnabled = True
        Me.cmbStatus.Items.AddRange(New Object() {"ทั้งหมด", "ถูกต้อง", "ไม่ถูกต้อง"})
        Me.cmbStatus.Location = New System.Drawing.Point(235, 48)
        Me.cmbStatus.Name = "cmbStatus"
        Me.cmbStatus.Size = New System.Drawing.Size(120, 23)
        Me.cmbStatus.TabIndex = 3

        '
        'lblFromDate
        '
        Me.lblFromDate.AutoSize = True
        Me.lblFromDate.Location = New System.Drawing.Point(375, 30)
        Me.lblFromDate.Name = "lblFromDate"
        Me.lblFromDate.Size = New System.Drawing.Size(50, 15)
        Me.lblFromDate.TabIndex = 4
        Me.lblFromDate.Text = "วันที่เริ่ม:"

        '
        'dtpFromDate
        '
        Me.dtpFromDate.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.dtpFromDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpFromDate.Location = New System.Drawing.Point(375, 48)
        Me.dtpFromDate.Name = "dtpFromDate"
        Me.dtpFromDate.Size = New System.Drawing.Size(120, 23)
        Me.dtpFromDate.TabIndex = 5

        '
        'lblToDate
        '
        Me.lblToDate.AutoSize = True
        Me.lblToDate.Location = New System.Drawing.Point(515, 30)
        Me.lblToDate.Name = "lblToDate"
        Me.lblToDate.Size = New System.Drawing.Size(53, 15)
        Me.lblToDate.TabIndex = 6
        Me.lblToDate.Text = "วันที่สิ้นสุด:"

        '
        'dtpToDate
        '
        Me.dtpToDate.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.dtpToDate.Format = System.Windows.Forms.DateTimePickerFormat.Short
        Me.dtpToDate.Location = New System.Drawing.Point(515, 48)
        Me.dtpToDate.Name = "dtpToDate"
        Me.dtpToDate.Size = New System.Drawing.Size(120, 23)
        Me.dtpToDate.TabIndex = 7

        '
        'btnRefresh
        '
        Me.btnRefresh.BackColor = System.Drawing.Color.FromArgb(52, 152, 219)
        Me.btnRefresh.FlatAppearance.BorderSize = 0
        Me.btnRefresh.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnRefresh.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnRefresh.ForeColor = System.Drawing.Color.White
        Me.btnRefresh.Location = New System.Drawing.Point(655, 45)
        Me.btnRefresh.Name = "btnRefresh"
        Me.btnRefresh.Size = New System.Drawing.Size(100, 28)
        Me.btnRefresh.TabIndex = 8
        Me.btnRefresh.Text = "รีเฟรช"
        Me.btnRefresh.UseVisualStyleBackColor = False

        '
        'pnlMain
        '
        Me.pnlMain.Controls.Add(Me.dgvHistory)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Location = New System.Drawing.Point(0, 160)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Padding = New System.Windows.Forms.Padding(10, 0, 10, 0)
        Me.pnlMain.Size = New System.Drawing.Size(1200, 400)
        Me.pnlMain.TabIndex = 2

        '
        'dgvHistory
        '
        Me.dgvHistory.AllowUserToAddRows = False
        Me.dgvHistory.AllowUserToDeleteRows = False
        Me.dgvHistory.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill
        Me.dgvHistory.BackgroundColor = System.Drawing.Color.White
        Me.dgvHistory.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvHistory.ColumnHeadersHeight = 35
        Me.dgvHistory.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvHistory.Dock = System.Windows.Forms.DockStyle.Fill
        Me.dgvHistory.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.dgvHistory.GridColor = System.Drawing.Color.FromArgb(224, 224, 224)
        Me.dgvHistory.Location = New System.Drawing.Point(10, 0)
        Me.dgvHistory.MultiSelect = False
        Me.dgvHistory.Name = "dgvHistory"
        Me.dgvHistory.ReadOnly = True
        Me.dgvHistory.RowHeadersVisible = False
        Me.dgvHistory.RowTemplate.Height = 30
        Me.dgvHistory.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect
        Me.dgvHistory.Size = New System.Drawing.Size(1180, 400)
        Me.dgvHistory.TabIndex = 0

        '
        'pnlButtons
        '
        Me.pnlButtons.BackColor = System.Drawing.Color.White
        Me.pnlButtons.Controls.Add(Me.btnViewDetail)
        Me.pnlButtons.Controls.Add(Me.btnDelete)
        Me.pnlButtons.Controls.Add(Me.btnExport)
        Me.pnlButtons.Controls.Add(Me.btnExportExcel)
        Me.pnlButtons.Controls.Add(Me.btnSettings)
        Me.pnlButtons.Controls.Add(Me.btnClose)
        Me.pnlButtons.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pnlButtons.Location = New System.Drawing.Point(0, 560)
        Me.pnlButtons.Name = "pnlButtons"
        Me.pnlButtons.Padding = New System.Windows.Forms.Padding(10)
        Me.pnlButtons.Size = New System.Drawing.Size(1200, 60)
        Me.pnlButtons.TabIndex = 3

        '
        'btnViewDetail
        '
        Me.btnViewDetail.BackColor = System.Drawing.Color.FromArgb(52, 152, 219)
        Me.btnViewDetail.Enabled = False
        Me.btnViewDetail.FlatAppearance.BorderSize = 0
        Me.btnViewDetail.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnViewDetail.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnViewDetail.ForeColor = System.Drawing.Color.White
        Me.btnViewDetail.Location = New System.Drawing.Point(15, 15)
        Me.btnViewDetail.Name = "btnViewDetail"
        Me.btnViewDetail.Size = New System.Drawing.Size(120, 35)
        Me.btnViewDetail.TabIndex = 0
        Me.btnViewDetail.Text = "ดูรายละเอียด"
        Me.btnViewDetail.UseVisualStyleBackColor = False

        '
        'btnDelete
        '
        Me.btnDelete.BackColor = System.Drawing.Color.FromArgb(231, 76, 60)
        Me.btnDelete.Enabled = False
        Me.btnDelete.FlatAppearance.BorderSize = 0
        Me.btnDelete.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnDelete.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnDelete.ForeColor = System.Drawing.Color.White
        Me.btnDelete.Location = New System.Drawing.Point(145, 15)
        Me.btnDelete.Name = "btnDelete"
        Me.btnDelete.Size = New System.Drawing.Size(100, 35)
        Me.btnDelete.TabIndex = 1
        Me.btnDelete.Text = "ลบ"
        Me.btnDelete.UseVisualStyleBackColor = False

        '
        'btnExport
        '
        Me.btnExport.BackColor = System.Drawing.Color.FromArgb(46, 125, 50)
        Me.btnExport.FlatAppearance.BorderSize = 0
        Me.btnExport.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExport.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnExport.ForeColor = System.Drawing.Color.White
        Me.btnExport.Location = New System.Drawing.Point(255, 15)
        Me.btnExport.Name = "btnExport"
        Me.btnExport.Size = New System.Drawing.Size(100, 35)
        Me.btnExport.TabIndex = 2
        Me.btnExport.Text = "Export CSV"
        Me.btnExport.UseVisualStyleBackColor = False

        '
        'btnExportExcel
        '
        Me.btnExportExcel.BackColor = System.Drawing.Color.FromArgb(46, 125, 50)
        Me.btnExportExcel.FlatAppearance.BorderSize = 0
        Me.btnExportExcel.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnExportExcel.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnExportExcel.ForeColor = System.Drawing.Color.White
        Me.btnExportExcel.Location = New System.Drawing.Point(365, 15)
        Me.btnExportExcel.Name = "btnExportExcel"
        Me.btnExportExcel.Size = New System.Drawing.Size(110, 35)
        Me.btnExportExcel.TabIndex = 3
        Me.btnExportExcel.Text = "Export Excel"
        Me.btnExportExcel.UseVisualStyleBackColor = False

        '
        'btnSettings
        '
        Me.btnSettings.BackColor = System.Drawing.Color.FromArgb(41, 128, 185)
        Me.btnSettings.FlatAppearance.BorderSize = 0
        Me.btnSettings.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnSettings.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnSettings.ForeColor = System.Drawing.Color.White
        Me.btnSettings.Location = New System.Drawing.Point(485, 15)
        Me.btnSettings.Name = "btnSettings"
        Me.btnSettings.Size = New System.Drawing.Size(110, 35)
        Me.btnSettings.TabIndex = 4
        Me.btnSettings.Text = "ตั้งค่าฐานข้อมูล"
        Me.btnSettings.UseVisualStyleBackColor = False

        '
        'btnClose
        '
        Me.btnClose.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btnClose.BackColor = System.Drawing.Color.FromArgb(108, 117, 125)
        Me.btnClose.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btnClose.FlatAppearance.BorderSize = 0
        Me.btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
        Me.btnClose.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Bold)
        Me.btnClose.ForeColor = System.Drawing.Color.White
        Me.btnClose.Location = New System.Drawing.Point(1085, 15)
        Me.btnClose.Name = "btnClose"
        Me.btnClose.Size = New System.Drawing.Size(100, 35)
        Me.btnClose.TabIndex = 4
        Me.btnClose.Text = "ปิด"
        Me.btnClose.UseVisualStyleBackColor = False

        '
        'statusStrip
        '
        Me.statusStrip.BackColor = System.Drawing.Color.FromArgb(236, 240, 241)
        Me.statusStrip.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.lblCount, Me.toolStripProgressBar})
        Me.statusStrip.Location = New System.Drawing.Point(0, 620)
        Me.statusStrip.Name = "statusStrip"
        Me.statusStrip.Size = New System.Drawing.Size(1200, 22)
        Me.statusStrip.TabIndex = 4

        '
        'lblCount
        '
        Me.lblCount.Name = "lblCount"
        Me.lblCount.Size = New System.Drawing.Size(1085, 17)
        Me.lblCount.Spring = True
        Me.lblCount.Text = "จำนวนรายการ: 0"
        Me.lblCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft

        '
        'toolStripProgressBar
        '
        Me.toolStripProgressBar.Name = "toolStripProgressBar"
        Me.toolStripProgressBar.Size = New System.Drawing.Size(100, 16)
        Me.toolStripProgressBar.Visible = False

        '
        'saveFileDialog
        '
        Me.saveFileDialog.DefaultExt = "csv"
        Me.saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        Me.saveFileDialog.Title = "ส่งออกข้อมูล"

        '
        'openFileDialog
        '
        Me.openFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"
        Me.openFileDialog.Title = "เลือกไฟล์"

        '
        'frmHistory
        '
        Me.AcceptButton = Me.btnRefresh
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.White
        Me.CancelButton = Me.btnClose
        Me.ClientSize = New System.Drawing.Size(1200, 642)
        Me.Controls.Add(Me.pnlMain)
        Me.Controls.Add(Me.pnlButtons)
        Me.Controls.Add(Me.pnlFilter)
        Me.Controls.Add(Me.pnlHeader)
        Me.Controls.Add(Me.statusStrip)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!)
        Me.MinimumSize = New System.Drawing.Size(1000, 600)
        Me.Name = "frmHistory"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "ประวัติการสแกน QR Code - QR Code Scanner System"

        Me.pnlHeader.ResumeLayout(False)
        Me.pnlHeader.PerformLayout()
        CType(Me.picIcon, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlFilter.ResumeLayout(False)
        Me.grpFilter.ResumeLayout(False)
        Me.grpFilter.PerformLayout()
        Me.pnlMain.ResumeLayout(False)
        CType(Me.dgvHistory, System.ComponentModel.ISupportInitialize).EndInit()
        Me.pnlButtons.ResumeLayout(False)
        Me.statusStrip.ResumeLayout(False)
        Me.statusStrip.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()
    End Sub

    ' Control declarations
    Friend WithEvents pnlHeader As Panel
    Friend WithEvents lblTitle As Label
    Friend WithEvents picIcon As PictureBox
    Friend WithEvents pnlFilter As Panel
    Friend WithEvents grpFilter As GroupBox
    Friend WithEvents lblSearch As Label
    Friend WithEvents txtSearch As TextBox
    Friend WithEvents lblStatus As Label
    Friend WithEvents cmbStatus As ComboBox
    Friend WithEvents lblFromDate As Label
    Friend WithEvents dtpFromDate As DateTimePicker
    Friend WithEvents lblToDate As Label
    Friend WithEvents dtpToDate As DateTimePicker
    Friend WithEvents btnRefresh As Button
    Friend WithEvents pnlMain As Panel
    Friend WithEvents dgvHistory As DataGridView
    Friend WithEvents pnlButtons As Panel
    Friend WithEvents btnViewDetail As Button
    Friend WithEvents btnDelete As Button
    Friend WithEvents btnExport As Button
    Friend WithEvents btnExportExcel As Button
    Friend WithEvents btnClose As Button
    Friend WithEvents statusStrip As StatusStrip
    Friend WithEvents lblCount As ToolStripStatusLabel
    Friend WithEvents toolStripProgressBar As ToolStripProgressBar
    Friend WithEvents saveFileDialog As SaveFileDialog
    Friend WithEvents openFileDialog As OpenFileDialog
    Friend WithEvents btnSettings As Button

End Class