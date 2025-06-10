Public Class frmHistory

#Region "Variables"
    Private scanHistory As List(Of ScanDataRecord)
    Private filteredHistory As List(Of ScanDataRecord)
    Private isLoading As Boolean = False
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            InitializeForm()
            SetupDataGridView()
            LoadScanHistory()
            'LoadScanHistory()
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดฟอร์ม: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHistory_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            RefreshData()
        Catch ex As Exception
            Console.WriteLine($"Error in frmHistory_Shown: {ex.Message}")
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub
#End Region

#Region "Initialization"
    Private Sub InitializeForm()
        Try
            ' ตั้งค่าเริ่มต้นสำหรับ ComboBox
            If cmbStatus.Items.Count > 0 Then
                cmbStatus.SelectedIndex = 0
            End If

            ' ตั้งค่าวันที่เริ่มต้น
            dtpFromDate.Value = DateTime.Now.AddDays(-7)
            dtpToDate.Value = DateTime.Now

            ' ตั้งค่าเริ่มต้นสำหรับสถานะปุ่ม
            btnViewDetail.Enabled = False
            btnDelete.Enabled = False

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเริ่มต้น: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetupDataGridView()
        Try
            ' เคลียร์คอลัมน์เดิม
            dgvHistory.Columns.Clear()

            ' ตั้งค่าพื้นฐานของ DataGridView
            dgvHistory.AutoGenerateColumns = False
            dgvHistory.AllowUserToAddRows = False
            dgvHistory.AllowUserToDeleteRows = False
            dgvHistory.ReadOnly = True
            dgvHistory.SelectionMode = DataGridViewSelectionMode.FullRowSelect
            dgvHistory.MultiSelect = False

            ' สร้างคอลัมน์วันที่/เวลา
            Dim colDateTime As New DataGridViewTextBoxColumn() With {
                .Name = "ScanDateTime",
                .HeaderText = "วันที่/เวลา",
                .DataPropertyName = "ScanDateTime",
                .Width = 150,
                .DefaultCellStyle = New DataGridViewCellStyle() With {.Format = "dd/MM/yyyy HH:mm:ss"}
            }
            dgvHistory.Columns.Add(colDateTime)

            ' สร้างคอลัมน์รหัสผลิตภัณฑ์
            Dim colProductCode As New DataGridViewTextBoxColumn() With {
                .Name = "ProductCode",
                .HeaderText = "รหัสผลิตภัณฑ์",
                .DataPropertyName = "ProductCode",
                .Width = 180
            }
            dgvHistory.Columns.Add(colProductCode)

            ' สร้างคอลัมน์รหัสอ้างอิง
            Dim colReferenceCode As New DataGridViewTextBoxColumn() With {
                .Name = "ReferenceCode",
                .HeaderText = "รหัสอ้างอิง",
                .DataPropertyName = "ReferenceCode",
                .Width = 150
            }
            dgvHistory.Columns.Add(colReferenceCode)

            ' สร้างคอลัมน์จำนวน
            Dim colQuantity As New DataGridViewTextBoxColumn() With {
                .Name = "Quantity",
                .HeaderText = "จำนวน",
                .DataPropertyName = "Quantity",
                .Width = 80
            }
            dgvHistory.Columns.Add(colQuantity)

            ' สร้างคอลัมน์วันที่ผลิต
            Dim colDateCode As New DataGridViewTextBoxColumn() With {
                .Name = "DateCode",
                .HeaderText = "วันที่ผลิต",
                .DataPropertyName = "DateCode",
                .Width = 100
            }
            dgvHistory.Columns.Add(colDateCode)

            ' สร้างคอลัมน์สถานะ
            Dim colStatus As New DataGridViewTextBoxColumn() With {
                .Name = "Status",
                .HeaderText = "สถานะ",
                .Width = 100
            }
            dgvHistory.Columns.Add(colStatus)

            ' สร้างคอลัมน์เครื่อง
            Dim colComputerName As New DataGridViewTextBoxColumn() With {
                .Name = "ComputerName",
                .HeaderText = "เครื่อง",
                .DataPropertyName = "ComputerName",
                .Width = 100
            }
            dgvHistory.Columns.Add(colComputerName)

            ' สร้างคอลัมน์ผู้ใช้
            Dim colUserName As New DataGridViewTextBoxColumn() With {
                .Name = "UserName",
                .HeaderText = "ผู้ใช้",
                .DataPropertyName = "UserName",
                .Width = 100
            }
            dgvHistory.Columns.Add(colUserName)

            Console.WriteLine($"DataGridView setup completed with {dgvHistory.Columns.Count} columns")

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตั้งค่า DataGridView: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Data Management"
    Private Sub LoadScanHistory()
        Try
            isLoading = True
            toolStripProgressBar.Visible = True
            toolStripProgressBar.Style = ProgressBarStyle.Marquee

            Console.WriteLine("Loading scan history...")

            ' ตรวจสอบการเชื่อมต่อฐานข้อมูล
            If Not DatabaseManager.IsConnected Then
                Console.WriteLine("Database not connected, attempting to initialize...")
                If Not DatabaseManager.Initialize() Then
                    Console.WriteLine("Database initialization failed, creating test data...")
                    CreateTestData()
                    Return
                End If
            End If

            ' โหลดข้อมูลจากฐานข้อมูล
            scanHistory = DatabaseManager.GetScanHistory(1000)
            Console.WriteLine($"Loaded {scanHistory.Count} records from database")

            ' ถ้าไม่มีข้อมูลจากฐานข้อมูล ให้สร้างข้อมูลทดสอบ
            If scanHistory.Count = 0 Then
                CreateTestData()
            Else
                filteredHistory = New List(Of ScanDataRecord)(scanHistory)
                RefreshDataGridView()
                UpdateRecordCount()
            End If

        Catch ex As Exception
            Console.WriteLine($"Error loading data: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            CreateTestData()
        Finally
            isLoading = False
            toolStripProgressBar.Visible = False
        End Try
    End Sub

    Private Sub CreateTestData()
        Try
            Console.WriteLine("Creating test data...")
            scanHistory = New List(Of ScanDataRecord)()

            ' สร้างข้อมูลทดสอบ 15 รายการ
            For i As Integer = 1 To 15
                Dim testRecord As New ScanDataRecord() With {
                    .ScanDateTime = DateTime.Now.AddHours(-i),
                    .OriginalData = $"R00C-19160425501276{i}+Q000060+P20414-00770{i:D2}A000+D20250527+LPT0000000+V00C-191604+U0000000",
                    .ExtractedData = $"20414-00770{i:D2}A000",
                    .ProductCode = $"20414-00770{i:D2}A000",
                    .ReferenceCode = $"00C-19160425501276{i}",
                    .Quantity = "60",
                    .DateCode = "20250527",
                    .IsValid = (i Mod 3 <> 0),
                    .ValidationMessages = If(i Mod 3 = 0, "ข้อมูลไม่สมบูรณ์", ""),
                    .ComputerName = Environment.MachineName,
                    .UserName = Environment.UserName
                }
                scanHistory.Add(testRecord)
            Next

            filteredHistory = New List(Of ScanDataRecord)(scanHistory)
            RefreshDataGridView()
            UpdateRecordCount()

            Console.WriteLine($"Created {scanHistory.Count} test records")

        Catch ex As Exception
            Console.WriteLine($"Error creating test data: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการสร้างข้อมูลทดสอบ: {ex.Message}")
        End Try
    End Sub

    Private Sub RefreshDataGridView()
        Try
            If isLoading Then Return
            If filteredHistory Is Nothing Then
                Console.WriteLine("filteredHistory is Nothing")
                Return
            End If

            Console.WriteLine($"Refreshing DataGridView with {filteredHistory.Count} records")

            ' ใช้ BindingSource เพื่อป้องกันปัญหา threading
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = filteredHistory
            dgvHistory.DataSource = bindingSource

        Catch ex As Exception
            Console.WriteLine($"Error refreshing DataGridView: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการรีเฟรชข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshData()
        LoadScanHistory()
    End Sub

    Private Sub UpdateRecordCount()
        Try
            lblCount.Text = $"จำนวนรายการ: {If(filteredHistory?.Count, 0)} จาก {If(scanHistory?.Count, 0)} รายการทั้งหมด"
            Console.WriteLine(lblCount.Text)
        Catch ex As Exception
            Console.WriteLine($"Error updating record count: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub dgvHistory_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles dgvHistory.DataBindingComplete
        Try
            Console.WriteLine($"DataBindingComplete fired. Rows: {dgvHistory.Rows.Count}")

            If dgvHistory.Rows.Count = 0 Then
                Console.WriteLine("No rows in DataGridView after binding")
            Else
                Console.WriteLine("DataGridView has data, updating status column...")

                ' อัปเดตคอลัมน์สถานะ
                For Each row As DataGridViewRow In dgvHistory.Rows
                    If row.DataBoundItem IsNot Nothing Then
                        Dim record As ScanDataRecord = CType(row.DataBoundItem, ScanDataRecord)
                        If dgvHistory.Columns.Contains("Status") Then
                            row.Cells("Status").Value = If(record.IsValid, "✅ ถูกต้อง", "❌ ไม่ถูกต้อง")
                            row.Cells("Status").Style.ForeColor = If(record.IsValid, Color.Green, Color.Red)
                        End If
                    End If
                Next
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in DataBindingComplete: {ex.Message}")
        End Try
    End Sub

    Private Sub dgvHistory_SelectionChanged(sender As Object, e As EventArgs) Handles dgvHistory.SelectionChanged
        Try
            Dim hasSelection As Boolean = dgvHistory.SelectedRows.Count > 0
            btnViewDetail.Enabled = hasSelection
            btnDelete.Enabled = hasSelection

        Catch ex As Exception
            Console.WriteLine($"Error in selection changed: {ex.Message}")
        End Try
    End Sub

    Private Sub dgvHistory_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs) Handles dgvHistory.CellFormatting
        Try
            ' เปลี่ยนสีแถวตามสถานะ
            If e.RowIndex >= 0 AndAlso dgvHistory.Rows(e.RowIndex).DataBoundItem IsNot Nothing Then
                Dim record As ScanDataRecord = CType(dgvHistory.Rows(e.RowIndex).DataBoundItem, ScanDataRecord)
                If Not record.IsValid Then
                    dgvHistory.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235)
                End If
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in cell formatting: {ex.Message}")
        End Try
    End Sub

    Private Sub dgvHistory_DoubleClick(sender As Object, e As EventArgs) Handles dgvHistory.DoubleClick
        btnViewDetail_Click(sender, e)
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        RefreshData()
    End Sub

    Private Sub btnViewDetail_Click(sender As Object, e As EventArgs) Handles btnViewDetail.Click
        Try
            If dgvHistory.SelectedRows.Count = 0 Then Return

            Dim selectedRecord As ScanDataRecord = CType(dgvHistory.SelectedRows(0).DataBoundItem, ScanDataRecord)
            ShowDetailDialog(selectedRecord)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs) Handles btnDelete.Click
        Try
            If dgvHistory.SelectedRows.Count = 0 Then Return

            Dim result As DialogResult = MessageBox.Show(
                "คุณต้องการลบรายการนี้หรือไม่?",
                "ยืนยันการลบ",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                MessageBox.Show("ฟีเจอร์การลบจะถูกเพิ่มในเวอร์ชันถัดไป", "แจ้งเตือน",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการลบข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnExport_Click(sender As Object, e As EventArgs) Handles btnExport.Click
        Try
            ExportToCSV()
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs) Handles btnExportExcel.Click
        Try
            MessageBox.Show("ฟีเจอร์ส่งออก Excel จะถูกเพิ่มในเวอร์ชันถัดไป", "แจ้งเตือน",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาด: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Filter and Search"
    Private Sub ApplyFilters()
        Try
            If isLoading OrElse scanHistory Is Nothing Then Return

            filteredHistory = New List(Of ScanDataRecord)(scanHistory)

            ' กรองตามข้อความค้นหา
            If Not String.IsNullOrEmpty(txtSearch.Text) Then
                Dim searchText As String = txtSearch.Text.ToLower()
                filteredHistory = filteredHistory.Where(Function(x)
                                                            Return x.ProductCode.ToLower().Contains(searchText) OrElse
                                                                  x.ReferenceCode.ToLower().Contains(searchText) OrElse
                                                                  x.OriginalData.ToLower().Contains(searchText) OrElse
                                                                  x.ExtractedData.ToLower().Contains(searchText)
                                                        End Function).ToList()
            End If

            ' กรองตามสถานะ
            If cmbStatus.SelectedIndex > 0 Then
                Dim isValid As Boolean = (cmbStatus.SelectedIndex = 1)
                filteredHistory = filteredHistory.Where(Function(x) x.IsValid = isValid).ToList()
            End If

            ' กรองตามช่วงวันที่
            Dim fromDate As DateTime = dtpFromDate.Value.Date
            Dim toDate As DateTime = dtpToDate.Value.Date.AddDays(1).AddSeconds(-1)

            filteredHistory = filteredHistory.Where(Function(x)
                                                        Return x.ScanDateTime >= fromDate AndAlso x.ScanDateTime <= toDate
                                                    End Function).ToList()

            RefreshDataGridView()
            UpdateRecordCount()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการกรองข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

#Region "Utility Methods"
    Private Sub ShowDetailDialog(record As ScanDataRecord)
        Try
            Dim detailForm As New Form() With {
                .Text = "รายละเอียดการสแกน",
                .Size = New Size(600, 500),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .MaximizeBox = False,
                .MinimizeBox = False,
                .Font = New Font("Segoe UI", 9)
            }

            ' สร้าง TextBox แสดงรายละเอียด
            Dim txtDetail As New TextBox() With {
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .ReadOnly = True,
                .Dock = DockStyle.Fill,
                .Font = New Font("Consolas", 10),
                .Margin = New Padding(10)
            }

            ' สร้างข้อความรายละเอียด
            Dim details As New System.Text.StringBuilder()
            details.AppendLine("=== รายละเอียดการสแกน QR Code ===")
            details.AppendLine()
            details.AppendLine($"วันที่/เวลา: {record.ScanDateTime:dd/MM/yyyy HH:mm:ss}")
            details.AppendLine($"สถานะ: {If(record.IsValid, "✅ ถูกต้อง", "❌ ไม่ถูกต้อง")}")
            details.AppendLine()
            details.AppendLine("ข้อมูลที่ดึงออกมา:")
            details.AppendLine($"  รหัสผลิตภัณฑ์: {record.ProductCode}")
            details.AppendLine($"  รหัสอ้างอิง: {record.ReferenceCode}")
            details.AppendLine($"  จำนวน: {record.Quantity}")
            details.AppendLine($"  วันที่ผลิต: {record.DateCode}")
            details.AppendLine()
            details.AppendLine($"เครื่องที่สแกน: {record.ComputerName}")
            details.AppendLine($"ผู้ใช้: {record.UserName}")
            details.AppendLine()

            If Not String.IsNullOrEmpty(record.ValidationMessages) Then
                details.AppendLine("ข้อความเตือน:")
                details.AppendLine(record.ValidationMessages)
                details.AppendLine()
            End If

            details.AppendLine("ข้อมูลต้นฉบับ:")
            details.AppendLine(record.OriginalData)

            txtDetail.Text = details.ToString()

            ' เพิ่ม panel สำหรับปุ่ม
            Dim pnlDetailButtons As New Panel() With {
                .Height = 50,
                .Dock = DockStyle.Bottom
            }

            Dim btnDetailClose As New Button() With {
                .Text = "ปิด",
                .Size = New Size(75, 30),
                .Location = New Point(detailForm.Width - 95, 10),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                .DialogResult = DialogResult.OK
            }

            pnlDetailButtons.Controls.Add(btnDetailClose)
            detailForm.Controls.Add(pnlDetailButtons)
            detailForm.Controls.Add(txtDetail)

            detailForm.ShowDialog()
            detailForm.Dispose()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub ExportToCSV()
        Try
            If filteredHistory Is Nothing OrElse filteredHistory.Count = 0 Then
                MessageBox.Show("ไม่มีข้อมูลสำหรับส่งออก", "แจ้งเตือน",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            saveFileDialog.FileName = $"ScanHistory_{DateTime.Now:yyyyMMdd_HHmmss}.csv"

            If saveFileDialog.ShowDialog() = DialogResult.OK Then
                Using writer As New IO.StreamWriter(saveFileDialog.FileName, False, System.Text.Encoding.UTF8)
                    ' เขียน header
                    writer.WriteLine("วันที่/เวลา,รหัสผลิตภัณฑ์,รหัสอ้างอิง,จำนวน,วันที่ผลิต,สถานะ,เครื่อง,ผู้ใช้,ข้อมูลต้นฉบับ")

                    ' เขียนข้อมูล
                    For Each record As ScanDataRecord In filteredHistory
                        Dim line As String = $"""{record.ScanDateTime:dd/MM/yyyy HH:mm:ss}"",""{record.ProductCode}"",""{record.ReferenceCode}"",""{record.Quantity}"",""{record.DateCode}"",""{If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง")}"",""{record.ComputerName}"",""{record.UserName}"",""{record.OriginalData.Replace("""", """""")}"""
                        writer.WriteLine(line)
                    Next
                End Using

                MessageBox.Show($"ส่งออกข้อมูลเรียบร้อยแล้ว ({filteredHistory.Count} รายการ)" & vbNewLine & $"ไฟล์: {saveFileDialog.FileName}",
                              "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

End Class