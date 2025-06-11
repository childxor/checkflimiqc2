Imports System.Net.NetworkInformation
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports Microsoft.Office.Interop.Excel

Public Class frmHistory

#Region "Variables"
    Private scanHistory As List(Of ScanDataRecord)
    Private filteredHistory As List(Of ScanDataRecord)
    Private isLoading As Boolean = False
    Private backgroundWorker As System.ComponentModel.BackgroundWorker
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Console.WriteLine("frmHistory_Load started")

            InitializeForm()
            SetupDataGridView()
            SetupBackgroundWorker()
            LoadScanHistory()

            Console.WriteLine("frmHistory_Load completed")

        Catch ex As Exception
            Console.WriteLine($"Error in frmHistory_Load: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดฟอร์ม: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        LoadScanHistory()
    End Sub

    Private Sub btnViewDetail_Click(sender As Object, e As EventArgs) Handles btnViewDetail.Click
        ViewSelectedRecord()
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        DeleteSelectedRecord()
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        ExportToCSV()
    End Sub

    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        ExportToExcel()
    End Sub

    Private Sub dgvHistory_SelectionChanged(sender As Object, e As EventArgs) Handles dgvHistory.SelectionChanged
        UpdateButtonStates()
    End Sub

    Private Sub dgvHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvHistory.CellContentClick
        ' จัดการการคลิกปุ่มในเซลล์
        If e.RowIndex >= 0 AndAlso e.ColumnIndex = 0 Then ' คอลัมน์ปุ่มตรวจสอบ Excel
            CheckExcelFile(e.RowIndex)
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        ApplyFilters()
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatus.SelectedIndexChanged
        ApplyFilters()
    End Sub

    Private Sub dtpFromDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpFromDate.ValueChanged
        ApplyFilters()
    End Sub

    Private Sub dtpToDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpToDate.ValueChanged
        ApplyFilters()
    End Sub
#End Region

#Region "Initialization"
    Private Sub InitializeForm()
        Try
            Console.WriteLine("InitializeForm started")

            ' ตั้งค่าเริ่มต้นสำหรับ ComboBox
            cmbStatus.Items.Clear()
            cmbStatus.Items.AddRange(New String() {"ทั้งหมด", "ถูกต้อง", "ไม่ถูกต้อง"})
            cmbStatus.SelectedIndex = 0

            ' ตั้งค่าวันที่เริ่มต้น
            dtpFromDate.Value = DateTime.Now.AddDays(-7)
            dtpToDate.Value = DateTime.Now

            ' ตั้งค่าเริ่มต้นสำหรับสถานะปุ่ม
            btnViewDetail.Enabled = False
            btnDelete.Enabled = False

            ' ซ่อน progress bar เริ่มต้น
            toolStripProgressBar.Visible = False

            Console.WriteLine("InitializeForm completed")

        Catch ex As Exception
            Console.WriteLine($"Error in InitializeForm: {ex.Message}")
            Throw
        End Try
    End Sub

    Private Sub SetupDataGridView()
        Try
            Console.WriteLine("SetupDataGridView started")

            ' เคลียร์คอลัมน์เดิม
            dgvHistory.Columns.Clear()
            dgvHistory.DataSource = Nothing

            ' ตั้งค่าพื้นฐานของ DataGridView
            dgvHistory.AutoGenerateColumns = False
            dgvHistory.AllowUserToAddRows = False
            dgvHistory.AllowUserToDeleteRows = False
            dgvHistory.ReadOnly = True
            dgvHistory.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgvHistory.MultiSelect = False
            dgvHistory.RowHeadersVisible = False

            ' สร้างคอลัมน์ปุ่มเช็คไฟล์ใน Excel
            Dim btnCol As New DataGridViewButtonColumn()
            btnCol.Name = "btnCheckExcel"
            btnCol.HeaderText = "ตรวจสอบ Excel"
            btnCol.Text = "🔍 ตรวจสอบ"
            btnCol.UseColumnTextForButtonValue = True
            btnCol.Width = 120
            dgvHistory.Columns.Add(btnCol)

            ' สร้างคอลัมน์ปุ่มสร้าง Mission
            Dim btnCreateMission As New DataGridViewButtonColumn()
            btnCreateMission.Name = "btnCreateMission"
            btnCreateMission.HeaderText = "สร้าง Mission"
            btnCreateMission.Text = "🚀 สร้าง"
            btnCreateMission.UseColumnTextForButtonValue = True
            btnCreateMission.Width = 65
            dgvHistory.Columns.Add(btnCreateMission)

            ' สร้างคอลัมน์วันที่/เวลา
            Dim colDateTime As New DataGridViewTextBoxColumn()
            colDateTime.Name = "ScanDateTime"
            colDateTime.HeaderText = "วันที่/เวลา"
            colDateTime.DataPropertyName = "ScanDateTime"
            colDateTime.Width = 150
            colDateTime.DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss"
            dgvHistory.Columns.Add(colDateTime)

            ' สร้างคอลัมน์รหัสผลิตภัณฑ์
            Dim colProduct As New DataGridViewTextBoxColumn()
            colProduct.Name = "ProductCode"
            colProduct.HeaderText = "รหัสผลิตภัณฑ์"
            colProduct.DataPropertyName = "ProductCode"
            colProduct.Width = 180
            dgvHistory.Columns.Add(colProduct)

            ' สร้างคอลัมน์รหัสอ้างอิง
            Dim colRef As New DataGridViewTextBoxColumn()
            colRef.Name = "ReferenceCode"
            colRef.HeaderText = "รหัสอ้างอิง"
            colRef.DataPropertyName = "ReferenceCode"
            colRef.Width = 150
            dgvHistory.Columns.Add(colRef)

            ' สร้างคอลัมน์จำนวน
            Dim colQty As New DataGridViewTextBoxColumn()
            colQty.Name = "Quantity"
            colQty.HeaderText = "จำนวน"
            colQty.DataPropertyName = "Quantity"
            colQty.Width = 80
            dgvHistory.Columns.Add(colQty)

            ' สร้างคอลัมน์วันที่ผลิต
            Dim colDate As New DataGridViewTextBoxColumn()
            colDate.Name = "DateCode"
            colDate.HeaderText = "วันที่ผลิต"
            colDate.DataPropertyName = "DateCode"
            colDate.Width = 100
            dgvHistory.Columns.Add(colDate)

            ' สร้างคอลัมน์สถานะ
            Dim colStatus As New DataGridViewTextBoxColumn()
            colStatus.Name = "StatusDisplay"
            colStatus.HeaderText = "สถานะ"
            colStatus.Width = 100
            dgvHistory.Columns.Add(colStatus)

            ' สร้างคอลัมน์เครื่อง
            Dim colComputer As New DataGridViewTextBoxColumn()
            colComputer.Name = "ComputerName"
            colComputer.HeaderText = "เครื่อง"
            colComputer.DataPropertyName = "ComputerName"
            colComputer.Width = 100
            dgvHistory.Columns.Add(colComputer)

            ' สร้างคอลัมน์ผู้ใช้
            Dim colUser As New DataGridViewTextBoxColumn()
            colUser.Name = "UserName"
            colUser.HeaderText = "ผู้ใช้"
            colUser.DataPropertyName = "UserName"
            colUser.Width = 100
            dgvHistory.Columns.Add(colUser)

            Console.WriteLine($"SetupDataGridView completed with {dgvHistory.Columns.Count} columns")

        Catch ex As Exception
            Console.WriteLine($"Error in SetupDataGridView: {ex.Message}")
            Throw
        End Try
    End Sub

    Private Sub SetupBackgroundWorker()
        Try
            backgroundWorker = New System.ComponentModel.BackgroundWorker()
            backgroundWorker.WorkerReportsProgress = True
            backgroundWorker.WorkerSupportsCancellation = True

            AddHandler backgroundWorker.DoWork, AddressOf BackgroundWorker_DoWork
            AddHandler backgroundWorker.ProgressChanged, AddressOf BackgroundWorker_ProgressChanged
            AddHandler backgroundWorker.RunWorkerCompleted, AddressOf BackgroundWorker_RunWorkerCompleted

        Catch ex As Exception
            Console.WriteLine($"Error in SetupBackgroundWorker: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Data Loading"
    Private Sub LoadScanHistory()
        Try
            If isLoading Then Return

            isLoading = True
            toolStripProgressBar.Visible = True
            toolStripProgressBar.Style = ProgressBarStyle.Marquee
            lblCount.Text = "กำลังโหลดข้อมูล..."

            ' ปิดปุ่มขณะโหลด
            btnRefresh.Enabled = False
            btnViewDetail.Enabled = False
            btnDelete.Enabled = False
            btnExport.Enabled = False
            btnExportExcel.Enabled = False

            If backgroundWorker IsNot Nothing AndAlso Not backgroundWorker.IsBusy Then
                backgroundWorker.RunWorkerAsync()
            Else
                ' ถ้า background worker ไม่พร้อม ให้โหลดแบบ synchronous
            End If
            LoadDataSynchronous()

        Catch ex As Exception
            Console.WriteLine($"Error in LoadScanHistory: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ResetLoadingState()
        End Try
    End Sub

    Private Sub BackgroundWorker_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs)
        Try
            Console.WriteLine("Background worker started")

            ' โหลดข้อมูลจากฐานข้อมูล
            Dim data As List(Of ScanDataRecord) = DatabaseManager.GetScanHistory(1000)

            backgroundWorker.ReportProgress(50, "กำลังประมวลผลข้อมูล...")

            e.Result = data

        Catch ex As Exception
            e.Result = ex
            Console.WriteLine($"Error in BackgroundWorker_DoWork: {ex.Message}")
        End Try
    End Sub

    Private Sub BackgroundWorker_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs)
        Try
            If e.UserState IsNot Nothing Then
                lblCount.Text = e.UserState.ToString()
            End If
        Catch ex As Exception
            Console.WriteLine($"Error in BackgroundWorker_ProgressChanged: {ex.Message}")
        End Try
    End Sub

    Private Sub BackgroundWorker_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs)
        Try
            If TypeOf e.Result Is Exception Then
                Dim ex As Exception = CType(e.Result, Exception)
                MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}",
                              "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                scanHistory = New List(Of ScanDataRecord)()
            ElseIf TypeOf e.Result Is List(Of ScanDataRecord) Then
                scanHistory = CType(e.Result, List(Of ScanDataRecord))
                Console.WriteLine($"Loaded {scanHistory.Count} records from background worker")
            Else
                scanHistory = New List(Of ScanDataRecord)()
            End If

            ApplyFilters()
            ResetLoadingState()

        Catch ex As Exception
            Console.WriteLine($"Error in BackgroundWorker_RunWorkerCompleted: {ex.Message}")
            scanHistory = New List(Of ScanDataRecord)()
            ResetLoadingState()
        End Try
    End Sub

    Private Sub LoadDataSynchronous()
        Try
            Console.WriteLine("Loading data synchronously")
            scanHistory = DatabaseManager.GetScanHistory(1000)
            Console.WriteLine($"Loaded {scanHistory.Count} records synchronously")

            ApplyFilters()
            ResetLoadingState()

        Catch ex As Exception
            Console.WriteLine($"Error in LoadDataSynchronous: {ex.Message}")
            scanHistory = New List(Of ScanDataRecord)()
            ResetLoadingState()
            Throw
        End Try
    End Sub

    Private Sub ResetLoadingState()
        Try
            isLoading = False
            toolStripProgressBar.Visible = False

            ' เปิดปุ่มกลับมา
            btnRefresh.Enabled = True
            btnExport.Enabled = True
            btnExportExcel.Enabled = True

            UpdateButtonStates()

        Catch ex As Exception
            Console.WriteLine($"Error in ResetLoadingState: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Data Filtering and Display"
    Private Sub ApplyFilters()
        Try
            If scanHistory Is Nothing Then
                scanHistory = New List(Of ScanDataRecord)()
            End If

            Dim filtered As IEnumerable(Of ScanDataRecord) = scanHistory

            ' กรองตามข้อความค้นหา
            If Not String.IsNullOrEmpty(txtSearch.Text.Trim()) Then
                Dim searchText As String = txtSearch.Text.Trim().ToLower()
                filtered = filtered.Where(Function(r) _
                    (Not String.IsNullOrEmpty(r.ProductCode) AndAlso r.ProductCode.ToLower().Contains(searchText)) OrElse
                    (Not String.IsNullOrEmpty(r.OriginalData) AndAlso r.OriginalData.ToLower().Contains(searchText)) OrElse
                    (Not String.IsNullOrEmpty(r.ReferenceCode) AndAlso r.ReferenceCode.ToLower().Contains(searchText))
                )
            End If

            ' กรองตามสถานะ
            If cmbStatus.SelectedIndex > 0 Then
                Dim isValid As Boolean = (cmbStatus.SelectedIndex = 1) ' 1 = ถูกต้อง, 2 = ไม่ถูกต้อง
                filtered = filtered.Where(Function(r) r.IsValid = isValid)
            End If

            ' กรองตามวันที่
            filtered = filtered.Where(Function(r)
                                          Dim scanDate As DateTime = r.ScanDateTime
                                          Dim fromDate As DateTime = dtpFromDate.Value
                                          Dim toDate As DateTime = dtpToDate.Value
                                          Return scanDate.Date >= fromDate.Date AndAlso scanDate.Date <= toDate.Date
                                      End Function)

            filteredHistory = filtered.OrderByDescending(Function(r) r.ScanDateTime).ToList()

            DisplayData()

        Catch ex As Exception
            Console.WriteLine($"Error in ApplyFilters: {ex.Message}")
            filteredHistory = New List(Of ScanDataRecord)()
            DisplayData()
        End Try
    End Sub

    Private Sub DisplayData()
        Try
            dgvHistory.Rows.Clear()

            If filteredHistory IsNot Nothing AndAlso filteredHistory.Count > 0 Then
                For Each record As ScanDataRecord In filteredHistory
                    Dim row As Integer = dgvHistory.Rows.Add()

                    With dgvHistory.Rows(row)
                        .Cells("ScanDateTime").Value = record.ScanDateTime
                        .Cells("ProductCode").Value = record.ProductCode
                        .Cells("ReferenceCode").Value = record.ReferenceCode
                        .Cells("Quantity").Value = record.Quantity
                        .Cells("DateCode").Value = record.DateCode
                        .Cells("StatusDisplay").Value = If(record.IsValid, "✅ ถูกต้อง", "❌ ไม่ถูกต้อง")
                        .Cells("ComputerName").Value = record.ComputerName
                        .Cells("UserName").Value = record.UserName

                        ' เก็บข้อมูลต้นฉบับใน Tag
                        .Tag = record

                        ' กำหนดสีตามสถานะ
                        If record.IsValid Then
                            .DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(46, 125, 50)
                        Else
                            .DefaultCellStyle.ForeColor = System.Drawing.Color.FromArgb(231, 76, 60)
                        End If
                    End With
                Next
            End If

            ' อัปเดตจำนวนรายการ
            Dim totalCount As Integer = If(scanHistory?.Count, 0)
            Dim filteredCount As Integer = If(filteredHistory?.Count, 0)

            If filteredCount = totalCount Then
                lblCount.Text = $"จำนวนรายการ: {totalCount}"
            Else
                lblCount.Text = $"จำนวนรายการ: {filteredCount} จาก {totalCount} รายการ"
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in DisplayData: {ex.Message}")
            lblCount.Text = "เกิดข้อผิดพลาดในการแสดงข้อมูล"
        End Try
    End Sub

    Private Sub UpdateButtonStates()
        Try
            Dim hasSelection As Boolean = dgvHistory.SelectedRows.Count > 0
            btnViewDetail.Enabled = hasSelection
            btnDelete.Enabled = hasSelection

        Catch ex As Exception
            Console.WriteLine($"Error in UpdateButtonStates: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Excel Integration"
    Private Sub CheckExcelFile(rowIndex As Integer)
        Try
            If rowIndex < 0 OrElse rowIndex >= dgvHistory.Rows.Count Then
                Return
            End If

            Dim record As ScanDataRecord = CType(dgvHistory.Rows(rowIndex).Tag, ScanDataRecord)
            If record Is Nothing OrElse String.IsNullOrEmpty(record.ProductCode) Then
                MessageBox.Show("ไม่พบรหัสผลิตภัณฑ์ในรายการนี้", "แจ้งเตือน",
                              System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If

            ' แสดงหน้าต่างสถานะ
            Dim statusForm As New System.Windows.Forms.Form()
            statusForm.Text = "กำลังตรวจสอบการเชื่อมต่อ"
            statusForm.Size = New Size(400, 120)
            statusForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            statusForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            statusForm.ControlBox = False
            statusForm.ShowInTaskbar = False

            Dim lblStatus As New System.Windows.Forms.Label()
            lblStatus.Text = "กำลังตรวจสอบการเชื่อมต่อกับเซิร์ฟเวอร์..."
            lblStatus.Location = New System.Drawing.Point(20, 30)
            lblStatus.AutoSize = True

            statusForm.Controls.Add(lblStatus)
            statusForm.Show(Me)
            System.Windows.Forms.Application.DoEvents()

            ' ตรวจสอบการเชื่อมต่อเครือข่าย
            Dim networkResult As NetworkCheckResult = CheckNetworkConnection()

            statusForm.Close()

            If networkResult.IsConnected Then
                HandleExcelFileAccess(record.ProductCode, networkResult)
            Else
                MessageBox.Show($"ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้{vbNewLine}{networkResult.ErrorMessage}",
                              "แจ้งเตือน", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตรวจสอบไฟล์ Excel: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in CheckExcelFile: {ex.Message}")
        End Try
    End Sub

    Private Function CheckNetworkConnection() As NetworkCheckResult
        Dim result As New NetworkCheckResult()

        Try
            Dim ping As New Ping()

            ' ทดสอบเครือข่าย OA ก่อน
            Try
                Dim replyOa As PingReply = ping.Send("10.24.179.2", 3000)
                If replyOa.Status = IPStatus.Success Then
                    result.IsConnected = True
                    result.NetworkType = "OA"
                    Return result
                End If
            Catch ex As Exception
                Console.WriteLine($"OA network test failed: {ex.Message}")
            End Try

            ' ทดสอบเครือข่าย FAB
            Try
                Dim replyFab As PingReply = ping.Send("172.24.0.3", 3000)
                If replyFab.Status = IPStatus.Success Then
                    result.IsConnected = True
                    result.NetworkType = "FAB"
                    Return result
                End If
            Catch ex As Exception
                Console.WriteLine($"FAB network test failed: {ex.Message}")
                result.ErrorMessage = $"ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้: {ex.Message}"
            End Try

            result.IsConnected = False
            If String.IsNullOrEmpty(result.ErrorMessage) Then
                result.ErrorMessage = "ไม่สามารถเชื่อมต่อกับเครือข่าย OA หรือ FAB ได้"
            End If

        Catch ex As Exception
            result.IsConnected = False
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการตรวจสอบเครือข่าย: {ex.Message}"
        End Try

        Return result
    End Function

    Private Sub HandleExcelFileAccess(productCode As String, networkResult As NetworkCheckResult)
        Try
            If networkResult.NetworkType = "OA" Then
                Dim excelPath As String = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

                If File.Exists(excelPath) Then
                    ' ค้นหาข้อมูลในไฟล์ Excel
                    SearchAndDisplayExcelData(excelPath, productCode)
                Else
                    MessageBox.Show($"ไม่พบไฟล์ Excel ที่ต้องการ:{vbNewLine}{excelPath}",
                                  "ตรวจสอบไฟล์ Excel", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                End If

            ElseIf networkResult.NetworkType = "FAB" Then
                MessageBox.Show("เครือข่าย FAB ไม่สามารถเข้าถึงไฟล์ Excel ได้{vbNewLine}กรุณาเชื่อมต่อกับเครือข่าย OA",
                              "ตรวจสอบไฟล์ Excel", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเข้าถึงไฟล์ Excel: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SearchAndDisplayExcelData(excelPath As String, productCode As String)
        Try
            ' แสดงหน้าต่างสถานะการค้นหา
            Dim searchForm As New System.Windows.Forms.Form()
            searchForm.Text = "กำลังค้นหาข้อมูล"
            searchForm.Size = New Size(400, 120)
            searchForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            searchForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            searchForm.ControlBox = False
            searchForm.ShowInTaskbar = False

            Dim lblSearchStatus As New System.Windows.Forms.Label()
            lblSearchStatus.Text = $"กำลังค้นหา '{productCode}' ในไฟล์ Excel..."
            lblSearchStatus.Location = New System.Drawing.Point(20, 30)
            lblSearchStatus.AutoSize = True

            searchForm.Controls.Add(lblSearchStatus)
            searchForm.Show(Me)
            System.Windows.Forms.Application.DoEvents()

            ' ค้นหาข้อมูลในไฟล์ Excel
            Dim searchResult As ExcelUtility.ExcelSearchResult = ExcelUtility.SearchProductInExcel(excelPath, productCode)

            searchForm.Close()

            ' แสดงผลลัพธ์
            DisplayExcelSearchResult(searchResult)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการค้นหาข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub DisplayExcelSearchResult(result As ExcelUtility.ExcelSearchResult)
        Try
            If result.IsSuccess AndAlso result.HasMatches Then
                ' แสดงผลลัพธ์ที่พบ
                Dim message As New System.Text.StringBuilder()
                message.AppendLine("🎉 พบข้อมูลในไฟล์ Excel!")
                message.AppendLine()
                message.AppendLine($"รหัสผลิตภัณฑ์: {result.SearchedProductCode}")
                message.AppendLine($"จำนวนที่พบ: {result.MatchCount} รายการ")
                message.AppendLine()

                If result.FirstMatch IsNot Nothing Then
                    message.AppendLine("ข้อมูลแรกที่พบ:")
                    message.AppendLine($"• แถวที่: {result.FirstMatch.RowNumber}")
                    If Not String.IsNullOrEmpty(result.FirstMatch.Column4Value) Then
                        message.AppendLine($"• ข้อมูลหลัก: {result.FirstMatch.Column4Value}")
                    End If

                    ' ค้นหาไฟล์ในโฟลเดอร์ตามข้อมูลที่พบ
                    Dim fileSearchResult = SearchFilesInDirectory(result.FirstMatch.Column4Value)
                    If fileSearchResult.FilesFound.Count > 0 Then
                        message.AppendLine()
                        message.AppendLine($"🔍 พบไฟล์ที่เกี่ยวข้อง: {fileSearchResult.FilesFound.Count} ไฟล์")

                        ' แสดงไฟล์ที่พบ (จำกัดที่ 5 ไฟล์แรก)
                        Dim maxDisplay As Integer = Math.Min(5, fileSearchResult.FilesFound.Count)
                        For i As Integer = 0 To maxDisplay - 1
                            Dim fileInfo = fileSearchResult.FilesFound(i)
                            message.AppendLine($"  📁 {fileInfo.FileName}")
                            message.AppendLine($"     ตำแหน่ง: {fileInfo.RelativePath}")
                            message.AppendLine($"     ขนาด: {FormatFileSize(fileInfo.FileSize)}")
                            message.AppendLine($"     แก้ไขล่าสุด: {fileInfo.LastModified:yyyy-MM-dd HH:mm}")
                            message.AppendLine()
                        Next

                        If fileSearchResult.FilesFound.Count > 5 Then
                            message.AppendLine($"... และอีก {fileSearchResult.FilesFound.Count - 5} ไฟล์")
                            message.AppendLine()
                        End If

                        ' แสดงสถิติการค้นหา
                        message.AppendLine($"📊 สถิติการค้นหา:")
                        message.AppendLine($"• โฟลเดอร์ที่ค้นหาทั้งหมด: {fileSearchResult.DirectoriesSearched}")
                        message.AppendLine($"• เวลาที่ใช้: {fileSearchResult.SearchDuration.TotalSeconds:F2} วินาที")

                        If fileSearchResult.ErrorDirectories.Count > 0 Then
                            message.AppendLine($"• โฟลเดอร์ที่เข้าถึงไม่ได้: {fileSearchResult.ErrorDirectories.Count}")
                        End If
                    Else
                        message.AppendLine()
                        message.AppendLine($"❌ ไม่พบไฟล์ที่เกี่ยวข้องกับ '{result.FirstMatch.Column4Value}'")
                        If fileSearchResult.ErrorDirectories.Count > 0 Then
                            message.AppendLine($"⚠️ มีโฟลเดอร์ที่เข้าถึงไม่ได้ {fileSearchResult.ErrorDirectories.Count} โฟลเดอร์")
                        End If
                    End If
                End If

                MessageBox.Show(message.ToString(), "ผลการค้นหา Excel",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' เสนอให้เปิดไฟล์ Excel
                If MessageBox.Show("ต้องการเปิดไฟล์ Excel เพื่อดูข้อมูลเพิ่มเติมหรือไม่?",
                              "เปิดไฟล์ Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                    OpenFileWithErrorHandling(result.ExcelFilePath)
                End If

                ' ถ้าพบไฟล์ ให้เสนอให้เปิดโฟลเดอร์หรือเปิดไฟล์โดยตรง
                Dim fileSearchResult2 = SearchFilesInDirectory(result.FirstMatch.Column4Value)
                If fileSearchResult2.FilesFound.Count > 0 Then
                    ' ถ้ามีไฟล์เดียว เสนอให้เปิดไฟล์หรือโฟลเดอร์
                    If fileSearchResult2.FilesFound.Count = 1 Then
                        Dim options As String() = {"เปิดไฟล์", "เปิดโฟลเดอร์", "ยกเลิก"}
                        Dim result2 = MessageBox.Show($"พบ 1 ไฟล์: {fileSearchResult2.FilesFound(0).FileName}{vbNewLine}ต้องการดำเนินการอย่างไร?",
                                    "เปิดไฟล์หรือโฟลเดอร์", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question)

                        If result2 = DialogResult.Yes Then ' เปิดไฟล์
                            OpenFileWithErrorHandling(fileSearchResult2.FilesFound(0).FullPath)
                        ElseIf result2 = DialogResult.No Then ' เปิดโฟลเดอร์
                            Dim fileDir = Path.GetDirectoryName(fileSearchResult2.FilesFound(0).FullPath)
                            OpenFileWithErrorHandling(fileDir)
                        End If
                    Else
                        ' มีหลายไฟล์ ให้แสดงรายการให้เลือก
                        Dim fileListForm As New Form()
                        fileListForm.Text = "เลือกไฟล์ที่ต้องการเปิด"
                        fileListForm.Size = New Size(600, 400)
                        fileListForm.StartPosition = FormStartPosition.CenterParent
                        fileListForm.MinimizeBox = False
                        fileListForm.MaximizeBox = False
                        fileListForm.FormBorderStyle = FormBorderStyle.FixedDialog

                        ' สร้าง ListView สำหรับแสดงรายการไฟล์
                        Dim listView As New ListView()
                        listView.View = View.Details
                        listView.FullRowSelect = True
                        listView.GridLines = True
                        listView.Dock = DockStyle.Fill
                        listView.Columns.Add("ชื่อไฟล์", 200)
                        listView.Columns.Add("ขนาด", 80)
                        listView.Columns.Add("แก้ไขล่าสุด", 120)
                        listView.Columns.Add("เส้นทาง", 350)

                        ' เพิ่มไฟล์ลงใน ListView 
                        For Each file In fileSearchResult2.FilesFound
                            Dim item As New ListViewItem(file.FileName)
                            item.SubItems.Add(FormatFileSize(file.FileSize))
                            item.SubItems.Add(file.LastModified.ToString("yyyy-MM-dd HH:mm"))
                            item.SubItems.Add(file.RelativePath)
                            item.Tag = file.FullPath
                            listView.Items.Add(item)
                        Next

                        ' ปรับขนาดคอลัมน์
                        listView.AutoResizeColumns(ColumnHeaderAutoResizeStyle.HeaderSize)

                        ' สร้าง Panel สำหรับใส่ปุ่ม
                        Dim buttonPanel As New Panel()
                        buttonPanel.Dock = DockStyle.Bottom
                        buttonPanel.Height = 50

                        ' สร้างปุ่มเปิดไฟล์
                        Dim btnOpen As New System.Windows.Forms.Button()
                        btnOpen.Text = "เปิดไฟล์"
                        btnOpen.Width = 100
                        btnOpen.Location = New System.Drawing.Point(10, 10)
                        btnOpen.Enabled = False

                        ' สร้างปุ่มเปิดโฟลเดอร์
                        Dim btnOpenFolder As New System.Windows.Forms.Button()
                        btnOpenFolder.Text = "เปิดโฟลเดอร์"
                        btnOpenFolder.Width = 100
                        btnOpenFolder.Location = New System.Drawing.Point(120, 10)
                        btnOpenFolder.Enabled = False

                        ' สร้างปุ่มยกเลิก
                        Dim btnCancel As New System.Windows.Forms.Button()
                        btnCancel.Text = "ยกเลิก"
                        btnCancel.Width = 100
                        btnCancel.Location = New System.Drawing.Point(230, 10)
                        btnCancel.DialogResult = DialogResult.Cancel

                        ' กำหนด Event เมื่อเลือกไฟล์
                        AddHandler listView.SelectedIndexChanged, Sub()
                                                                      btnOpen.Enabled = listView.SelectedItems.Count > 0
                                                                      btnOpenFolder.Enabled = listView.SelectedItems.Count > 0
                                                                  End Sub

                        ' กำหนด Event เมื่อดับเบิลคลิกที่ไฟล์
                        AddHandler listView.DoubleClick, Sub()
                                                             If listView.SelectedItems.Count > 0 Then
                                                                 OpenFileWithErrorHandling(listView.SelectedItems(0).Tag.ToString())
                                                                 fileListForm.Close()
                                                             End If
                                                         End Sub

                        ' กำหนด Event เมื่อคลิกปุ่มเปิดไฟล์
                        AddHandler btnOpen.Click, Sub()
                                                      If listView.SelectedItems.Count > 0 Then
                                                          OpenFileWithErrorHandling(listView.SelectedItems(0).Tag.ToString())
                                                          fileListForm.Close()
                                                      End If
                                                  End Sub

                        ' กำหนด Event เมื่อคลิกปุ่มเปิดโฟลเดอร์
                        AddHandler btnOpenFolder.Click, Sub()
                                                            If listView.SelectedItems.Count > 0 Then
                                                                Dim selectedFilePath = listView.SelectedItems(0).Tag.ToString()
                                                                Dim fileDir = Path.GetDirectoryName(selectedFilePath)
                                                                OpenFileWithErrorHandling(fileDir)
                                                                fileListForm.Close()
                                                            End If
                                                        End Sub

                        ' เพิ่ม Controls ลงในฟอร์ม
                        buttonPanel.Controls.Add(btnOpen)
                        buttonPanel.Controls.Add(btnOpenFolder)
                        buttonPanel.Controls.Add(btnCancel)
                        fileListForm.Controls.Add(listView)
                        fileListForm.Controls.Add(buttonPanel)

                        ' แสดงฟอร์ม
                        fileListForm.ShowDialog()
                    End If
                End If
            Else
                ' ไม่พบข้อมูล
                Dim message As String = $"❌ ไม่พบรหัสผลิตภัณฑ์ '{result.SearchedProductCode}' ในไฟล์ Excel"

                If result.HasError Then
                    message &= vbNewLine & vbNewLine & $"ข้อผิดพลาด: {result.ErrorMessage}"
                End If

                Dim dialogResult As DialogResult = MessageBox.Show(message & vbNewLine & vbNewLine & "ต้องการเปิดไฟล์ Excel เพื่อตรวจสอบด้วยตนเองหรือไม่?",
                                                          "ผลการค้นหา Excel", MessageBoxButtons.YesNo, MessageBoxIcon.Information)

                If dialogResult = DialogResult.Yes Then
                    OpenFileWithErrorHandling(result.ExcelFilePath)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงผลลัพธ์: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ค้นหาไฟล์ในโฟลเดอร์ตามชื่อที่กำหนด
    ''' </summary>
    Private Function SearchFilesInDirectory(fileName As String) As FileSearchResult
        Dim result As New FileSearchResult()
        Dim stopwatch As New Stopwatch()
        stopwatch.Start()

        Try
            If String.IsNullOrEmpty(fileName) Then
                Return result
            End If

            ' โฟลเดอร์หลักที่จะค้นหา
            Dim baseFolderPath As String = "\\fls951\OAFAB\OA2FAB\Film charecter check"

            ' ตรวจสอบว่าโฟลเดอร์หลักมีอยู่จริงหรือไม่
            If Not Directory.Exists(baseFolderPath) Then
                result.ErrorDirectories.Add($"โฟลเดอร์หลักไม่พบ: {baseFolderPath}")
                Return result
            End If

            ' สร้าง pattern สำหรับค้นหา
            Dim searchPatterns As String() = {
            fileName & "_*",           ' SN1C63Z083XU-01N_*
            fileName & ".*",           ' SN1C63Z083XU-01N.*
            "*" & fileName & "*",      ' *SN1C63Z083XU-01N*
            fileName                   ' SN1C63Z083XU-01N (ตรงทุกตัว)
        }

            ' ค้นหาในทุกโฟลเดอร์ย่อย
            SearchInDirectoryRecursive(baseFolderPath, searchPatterns, result)

        Catch ex As Exception
            result.ErrorDirectories.Add($"ข้อผิดพลาดทั่วไป: {ex.Message}")
        Finally
            stopwatch.Stop()
            result.SearchDuration = stopwatch.Elapsed
        End Try

        Return result
    End Function

    ''' <summary>
    ''' ค้นหาไฟล์แบบ recursive ในทุกโฟลเดอร์ย่อย
    ''' </summary>
    Private Sub SearchInDirectoryRecursive(directoryPath As String, searchPatterns As String(), result As FileSearchResult)
        Try
            result.DirectoriesSearched += 1

            ' ค้นหาไฟล์ในโฟลเดอร์ปัจจุบัน
            For Each pattern As String In searchPatterns
                Try
                    Dim files As String() = Directory.GetFiles(directoryPath, pattern, SearchOption.TopDirectoryOnly)

                    For Each filePath As String In files
                        Try
                            Dim fileInfo As New FileInfo(filePath)

                            ' สร้างข้อมูลไฟล์
                            Dim fileDetail As New FileDetail() With {
                            .FileName = Path.GetFileName(filePath),
                            .FullPath = filePath,
                            .RelativePath = GetRelativePath("\\fls951\OAFAB\OA2FAB\20250607 Pimploy S", filePath),
                            .FileSize = fileInfo.Length,
                            .LastModified = fileInfo.LastWriteTime
                        }

                            ' ตรวจสอบว่าไฟล์นี้ยังไม่ได้เพิ่มแล้ว (ป้องกันการซ้ำ)
                            If Not result.FilesFound.Any(Function(f) f.FullPath.Equals(filePath, StringComparison.OrdinalIgnoreCase)) Then
                                result.FilesFound.Add(fileDetail)
                            End If

                        Catch fileEx As Exception
                            ' ข้ามไฟล์ที่เข้าถึงไม่ได้
                            Continue For
                        End Try
                    Next

                Catch patternEx As Exception
                    ' ข้าม pattern ที่มีปัญหา
                    Continue For
                End Try
            Next

            ' ค้นหาในโฟลเดอร์ย่อย
            Try
                Dim subDirectories As String() = Directory.GetDirectories(directoryPath)

                For Each subDir As String In subDirectories
                    Try
                        SearchInDirectoryRecursive(subDir, searchPatterns, result)
                    Catch subDirEx As Exception
                        result.ErrorDirectories.Add($"ไม่สามารถเข้าถึงโฟลเดอร์: {subDir} - {subDirEx.Message}")
                    End Try
                Next

            Catch dirEx As Exception
                result.ErrorDirectories.Add($"ไม่สามารถดูโฟลเดอร์ย่อยใน: {directoryPath} - {dirEx.Message}")
            End Try

        Catch ex As Exception
            result.ErrorDirectories.Add($"ข้อผิดพลาดในโฟลเดอร์: {directoryPath} - {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' แปลง path เป็น relative path
    ''' </summary>
    Private Function GetRelativePath(basePath As String, fullPath As String) As String
        Try
            If fullPath.StartsWith(basePath, StringComparison.OrdinalIgnoreCase) Then
                Return fullPath.Substring(basePath.Length).TrimStart("\"c)
            End If
            Return fullPath
        Catch
            Return fullPath
        End Try
    End Function

    ''' <summary>
    ''' แปลงขนาดไฟล์เป็นรูปแบบที่อ่านง่าย
    ''' </summary>
    Private Function FormatFileSize(bytes As Long) As String
        Try
            Dim suffixes As String() = {"B", "KB", "MB", "GB", "TB"}
            Dim counter As Integer = 0
            Dim number As Decimal = bytes

            While number >= 1024 AndAlso counter < suffixes.Length - 1
                number /= 1024
                counter += 1
            End While

            Return $"{number:N1} {suffixes(counter)}"
        Catch
            Return $"{bytes} B"
        End Try
    End Function

    ''' <summary>
    ''' คลาสสำหรับเก็บผลลัพธ์การค้นหาไฟล์
    ''' </summary>
    Public Class FileSearchResult
        Public Property FilesFound As New List(Of FileDetail)()
        Public Property DirectoriesSearched As Integer = 0
        Public Property ErrorDirectories As New List(Of String)()
        Public Property SearchDuration As TimeSpan = TimeSpan.Zero
    End Class

    ''' <summary>
    ''' คลาสสำหรับเก็บรายละเอียดไฟล์
    ''' </summary>
    Public Class FileDetail
        Public Property FileName As String = ""
        Public Property FullPath As String = ""
        Public Property RelativePath As String = ""
        Public Property FileSize As Long = 0
        Public Property LastModified As DateTime = DateTime.MinValue
    End Class
#End Region

#Region "Data Operations"
    Private Sub ViewSelectedRecord()
        Try
            If dgvHistory.SelectedRows.Count = 0 Then
                Return
            End If

            Dim selectedRow As DataGridViewRow = dgvHistory.SelectedRows(0)
            If selectedRow.Tag Is Nothing Then
                MessageBox.Show("ไม่พบข้อมูลสำหรับรายการที่เลือก", "แจ้งเตือน", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If

            Dim record As ScanDataRecord = CType(selectedRow.Tag, ScanDataRecord)

            ' สร้างฟอร์มแสดงรายละเอียด
            Dim detailForm As New System.Windows.Forms.Form()
            detailForm.Text = "รายละเอียดการสแกน"
            detailForm.Size = New Size(600, 500)
            detailForm.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
            detailForm.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog
            detailForm.MaximizeBox = False
            detailForm.MinimizeBox = False

            ' สร้าง TextBox สำหรับแสดงข้อมูล
            Dim txtDetail As New System.Windows.Forms.TextBox()
            txtDetail.Multiline = True
            txtDetail.ReadOnly = True
            txtDetail.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
            txtDetail.Dock = System.Windows.Forms.DockStyle.Fill
            txtDetail.Font = New System.Drawing.Font("Segoe UI", 10)
            txtDetail.Text = $"รหัสผลิตภัณฑ์: {record.ProductCode}{Environment.NewLine}" &
                           $"รหัสอ้างอิง: {record.ReferenceCode}{Environment.NewLine}" &
                           $"วันที่/เวลาสแกน: {record.ScanDateTime}{Environment.NewLine}" &
                           $"จำนวน: {record.Quantity}{Environment.NewLine}" &
                           $"วันที่ผลิต: {record.DateCode}{Environment.NewLine}" &
                           $"สถานะ: {If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง")}{Environment.NewLine}" &
                           $"เครื่อง: {record.ComputerName}{Environment.NewLine}" &
                           $"ผู้ใช้: {record.UserName}{Environment.NewLine}{Environment.NewLine}" &
                           $"ข้อมูลต้นฉบับ:{Environment.NewLine}{record.OriginalData}"

            ' สร้างปุ่มปิด
            Dim btnClose As New System.Windows.Forms.Button()
            btnClose.Text = "ปิด"
            btnClose.DialogResult = System.Windows.Forms.DialogResult.OK
            btnClose.Dock = System.Windows.Forms.DockStyle.Bottom
            btnClose.Height = 40
            btnClose.BackColor = System.Drawing.Color.FromArgb(108, 117, 125)
            btnClose.ForeColor = System.Drawing.Color.White
            btnClose.FlatStyle = System.Windows.Forms.FlatStyle.Flat
            btnClose.Font = New System.Drawing.Font("Segoe UI", 9, System.Drawing.FontStyle.Bold)

            ' เพิ่ม Controls เข้าฟอร์ม
            detailForm.Controls.Add(txtDetail)
            detailForm.Controls.Add(btnClose)

            ' แสดงฟอร์ม
            detailForm.ShowDialog(Me)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด: {ex.Message}",
                          "ข้อผิดพลาด", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Console.WriteLine($"Error in ViewSelectedRecord: {ex.Message}")
        End Try
    End Sub

    Private Sub DeleteSelectedRecord()
        Try
            If dgvHistory.SelectedRows.Count = 0 Then
                Return
            End If

            Dim selectedRow As DataGridViewRow = dgvHistory.SelectedRows(0)
            If selectedRow.Tag Is Nothing Then
                MessageBox.Show("ไม่พบข้อมูลสำหรับรายการที่เลือก", "แจ้งเตือน", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If

            Dim record As ScanDataRecord = CType(selectedRow.Tag, ScanDataRecord)

            ' ยืนยันการลบ
            Dim result As System.Windows.Forms.DialogResult = MessageBox.Show($"คุณต้องการลบรายการนี้ใช่หรือไม่?{Environment.NewLine}{Environment.NewLine}" &
                                                      $"รหัสผลิตภัณฑ์: {record.ProductCode}{Environment.NewLine}" &
                                                      $"วันที่/เวลาสแกน: {record.ScanDateTime}",
                                                      "ยืนยันการลบ", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning)

            If result = System.Windows.Forms.DialogResult.Yes Then
                ' ลบข้อมูลจากฐานข้อมูล
                Dim success As Boolean = DatabaseManager.DeleteScanRecord(record.Id)

                If success Then
                    ' ลบออกจากรายการในหน้าจอ
                    dgvHistory.Rows.Remove(selectedRow)
                    scanHistory.Remove(record)
                    filteredHistory.Remove(record)

                    ' อัปเดตจำนวนรายการ
                    Dim totalCount As Integer = If(scanHistory?.Count, 0)
                    Dim filteredCount As Integer = If(filteredHistory?.Count, 0)

                    If filteredCount = totalCount Then
                        lblCount.Text = $"จำนวนรายการ: {totalCount}"
                    Else
                        lblCount.Text = $"จำนวนรายการ: {filteredCount} จาก {totalCount} รายการ"
                    End If

                    MessageBox.Show("ลบรายการสำเร็จ", "แจ้งเตือน", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)
                Else
                    MessageBox.Show("ไม่สามารถลบรายการได้", "ข้อผิดพลาด", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
                End If
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการลบรายการ: {ex.Message}",
                          "ข้อผิดพลาด", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Console.WriteLine($"Error in DeleteSelectedRecord: {ex.Message}")
        End Try
    End Sub

    Private Sub ExportToCSV()
        Try
            If filteredHistory Is Nothing OrElse filteredHistory.Count = 0 Then
                MessageBox.Show("ไม่มีข้อมูลที่จะส่งออก", "แจ้งเตือน", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If

            ' กำหนดค่าเริ่มต้นสำหรับไดอะล็อก
            saveFileDialog.FileName = $"ScanHistory_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            saveFileDialog.Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*"

            If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                ' แสดงแถบความคืบหน้า
                toolStripProgressBar.Visible = True
                toolStripProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee
                lblCount.Text = "กำลังส่งออกข้อมูล..."

                ' เริ่มเขียนไฟล์
                Using writer As New StreamWriter(saveFileDialog.FileName, False, System.Text.Encoding.UTF8)
                    ' เขียนหัวตาราง
                    writer.WriteLine("วันที่/เวลา,รหัสผลิตภัณฑ์,รหัสอ้างอิง,จำนวน,วันที่ผลิต,สถานะ,เครื่อง,ผู้ใช้,ข้อมูลต้นฉบับ")

                    ' เขียนข้อมูล
                    For Each record In filteredHistory
                        writer.WriteLine($"{record.ScanDateTime:yyyy-MM-dd HH:mm:ss}," &
                                      $"""{record.ProductCode}""," &
                                      $"""{record.ReferenceCode}""," &
                                      $"{record.Quantity}," &
                                      $"""{record.DateCode}""," &
                                      $"{If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง")}," &
                                      $"""{record.ComputerName}""," &
                                      $"""{record.UserName}""," &
                                      $"""{record.OriginalData.Replace("""", """""")}"" ")
                    Next
                End Using

                ' ซ่อนแถบความคืบหน้า
                toolStripProgressBar.Visible = False
                lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count} จาก {scanHistory.Count} รายการ"

                MessageBox.Show($"ส่งออกข้อมูลสำเร็จ{Environment.NewLine}ที่อยู่ไฟล์: {saveFileDialog.FileName}",
                              "สำเร็จ", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

                ' ถามว่าต้องการเปิดไฟล์หรือไม่
                If MessageBox.Show("ต้องการเปิดไฟล์ที่ส่งออกหรือไม่?", "เปิดไฟล์",
                                 System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    OpenFileWithErrorHandling(saveFileDialog.FileName)
                End If
            End If

        Catch ex As Exception
            toolStripProgressBar.Visible = False
            lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count} จาก {scanHistory.Count} รายการ"

            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Console.WriteLine($"Error in ExportToCSV: {ex.Message}")
        End Try
    End Sub

    Private Sub ExportToExcel()
        Try
            If filteredHistory Is Nothing OrElse filteredHistory.Count = 0 Then
                MessageBox.Show("ไม่มีข้อมูลที่จะส่งออก", "แจ้งเตือน", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Warning)
                Return
            End If

            ' กำหนดค่าเริ่มต้นสำหรับไดอะล็อก
            saveFileDialog.FileName = $"ScanHistory_{DateTime.Now:yyyyMMdd_HHmmss}.xlsx"
            saveFileDialog.Filter = "Excel Files (*.xlsx)|*.xlsx|All Files (*.*)|*.*"

            If saveFileDialog.ShowDialog() = System.Windows.Forms.DialogResult.OK Then
                ' แสดงแถบความคืบหน้า
                toolStripProgressBar.Visible = True
                toolStripProgressBar.Style = System.Windows.Forms.ProgressBarStyle.Marquee
                lblCount.Text = "กำลังส่งออกข้อมูล..."

                ' สร้าง Excel Application
                Dim excelApp As New Microsoft.Office.Interop.Excel.Application()
                excelApp.Visible = False

                ' สร้าง Workbook
                Dim workbook As Microsoft.Office.Interop.Excel.Workbook = excelApp.Workbooks.Add()
                Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = CType(workbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

                ' ตั้งชื่อ Sheet
                worksheet.Name = "ประวัติการสแกน"

                ' กำหนดหัวตาราง
                worksheet.Cells(1, 1) = "ลำดับ"
                worksheet.Cells(1, 2) = "วันที่/เวลา"
                worksheet.Cells(1, 3) = "รหัสผลิตภัณฑ์"
                worksheet.Cells(1, 4) = "รหัสอ้างอิง"
                worksheet.Cells(1, 5) = "จำนวน"
                worksheet.Cells(1, 6) = "วันที่ผลิต"
                worksheet.Cells(1, 7) = "สถานะ"
                worksheet.Cells(1, 8) = "เครื่อง"
                worksheet.Cells(1, 9) = "ผู้ใช้"
                worksheet.Cells(1, 10) = "ข้อมูลต้นฉบับ"

                ' จัดรูปแบบหัวตาราง
                Dim headerRange As Microsoft.Office.Interop.Excel.Range = worksheet.Range("A1:J1")
                headerRange.Font.Bold = True
                headerRange.Interior.Color = RGB(52, 152, 219)
                headerRange.Font.Color = RGB(255, 255, 255)

                ' เพิ่มข้อมูล
                For i As Integer = 0 To filteredHistory.Count - 1
                    Dim record As ScanDataRecord = filteredHistory(i)
                    Dim row As Integer = i + 2

                    worksheet.Cells(row, 1) = i + 1
                    worksheet.Cells(row, 2) = record.ScanDateTime
                    worksheet.Cells(row, 3) = record.ProductCode
                    worksheet.Cells(row, 4) = record.ReferenceCode
                    worksheet.Cells(row, 5) = record.Quantity
                    worksheet.Cells(row, 6) = record.DateCode
                    worksheet.Cells(row, 7) = If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง")
                    worksheet.Cells(row, 8) = record.ComputerName
                    worksheet.Cells(row, 9) = record.UserName
                    worksheet.Cells(row, 10) = record.OriginalData

                    ' กำหนดสีตามสถานะ
                    If Not record.IsValid Then
                        worksheet.Cells(row, 7).Interior.Color = RGB(231, 76, 60)
                        worksheet.Cells(row, 7).Font.Color = RGB(255, 255, 255)
                    End If
                Next

                ' ปรับขนาดคอลัมน์ให้พอดีกับข้อมูล
                worksheet.Columns("A:J").AutoFit()

                ' บันทึกไฟล์
                workbook.SaveAs(saveFileDialog.FileName)

                ' ปิด Excel
                workbook.Close(False)
                excelApp.Quit()

                ' คืนค่าหน่วยความจำ
                ReleaseObject(worksheet)
                ReleaseObject(workbook)
                ReleaseObject(excelApp)

                ' ซ่อนแถบความคืบหน้า
                toolStripProgressBar.Visible = False
                lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count} จาก {scanHistory.Count} รายการ"

                MessageBox.Show($"ส่งออกข้อมูลสำเร็จ{Environment.NewLine}ที่อยู่ไฟล์: {saveFileDialog.FileName}",
                              "สำเร็จ", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Information)

                ' ถามว่าต้องการเปิดไฟล์หรือไม่
                If MessageBox.Show("ต้องการเปิดไฟล์ที่ส่งออกหรือไม่?", "เปิดไฟล์",
                                 System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question) = System.Windows.Forms.DialogResult.Yes Then
                    OpenFileWithErrorHandling(saveFileDialog.FileName)
                End If
            End If

        Catch ex As Exception
            toolStripProgressBar.Visible = False
            lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count} จาก {scanHistory.Count} รายการ"

            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error)
            Console.WriteLine($"Error in ExportToExcel: {ex.Message}")
        End Try
    End Sub

    Private Sub ReleaseObject(ByVal obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    ''' <summary>
    ''' เปิดไฟล์หรือโฟลเดอร์ด้วยโปรแกรมที่เหมาะสมพร้อมจัดการข้อผิดพลาด
    ''' </summary>
    Private Sub OpenFileWithErrorHandling(filePath As String)
        Try
            ' ตรวจสอบว่าเป็นไฟล์หรือโฟลเดอร์
            If System.IO.File.Exists(filePath) Then
                ' วิธีที่ 1: ใช้ ProcessStartInfo เพื่อเปิดไฟล์อย่างปลอดภัย
                Dim startInfo As New System.Diagnostics.ProcessStartInfo()
                startInfo.FileName = filePath
                startInfo.UseShellExecute = True
                System.Diagnostics.Process.Start(startInfo)
            ElseIf System.IO.Directory.Exists(filePath) Then
                ' เปิดโฟลเดอร์โดยตรง
                Dim startInfo As New System.Diagnostics.ProcessStartInfo()
                startInfo.FileName = "explorer.exe"
                startInfo.Arguments = """" & filePath & """"
                System.Diagnostics.Process.Start(startInfo)
            Else
                MessageBox.Show($"ไม่พบไฟล์หรือโฟลเดอร์ที่ระบุ:{vbNewLine}{filePath}",
                              "ไม่พบไฟล์", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"ไม่สามารถเปิดไฟล์หรือโฟลเดอร์ได้:{vbNewLine}{ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)

            ' วิธีที่ 2: พยายามเปิดโฟลเดอร์และเลือกไฟล์ (สำหรับกรณีที่เป็นไฟล์เท่านั้น)
            Try
                If System.IO.File.Exists(filePath) Then
                    Dim explorerPath As String = System.IO.Path.GetDirectoryName(filePath)
                    If System.IO.Directory.Exists(explorerPath) Then
                        System.Diagnostics.Process.Start("explorer.exe", "/select," & filePath)
                    End If
                ElseIf System.IO.Directory.Exists(System.IO.Path.GetDirectoryName(filePath)) Then
                    ' ถ้าไฟล์ไม่มี แต่โฟลเดอร์พ่อแม่มี ให้เปิดโฟลเดอร์พ่อแม่
                    System.Diagnostics.Process.Start("explorer.exe", """" & System.IO.Path.GetDirectoryName(filePath) & """")
                End If
            Catch ex2 As Exception
                MessageBox.Show($"ไม่สามารถเปิดโฟลเดอร์ได้:{vbNewLine}{ex2.Message}",
                              "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            End Try
        End Try
    End Sub
#End Region

#Region "Support Classes"
    ' คลาสสำหรับผลลัพธ์การตรวจสอบเครือข่าย
    Public Class NetworkCheckResult
        Public Property IsConnected As Boolean = False
        Public Property NetworkType As String = ""
        Public Property ErrorMessage As String = ""
    End Class
#End Region

End Class
