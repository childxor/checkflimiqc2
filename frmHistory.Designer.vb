Public Class frmHistory

#Region "Variables"
    Private scanHistory As List(Of ScanDataRecord)
    Private filteredHistory As List(Of ScanDataRecord)
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()
        SetupDataGridView()
        LoadScanHistory()
    End Sub

    Private Sub frmHistory_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        RefreshData()
    End Sub
#End Region

#Region "Initialization"
    Private Sub InitializeForm()
        Try
            ' ตั้งค่าเริ่มต้นสำหรับ ComboBox
            cmbStatus.SelectedIndex = 0
            
            ' ตั้งค่าวันที่เริ่มต้น
            dtpFromDate.Value = DateTime.Now.AddDays(-7)
            dtpToDate.Value = DateTime.Now
            
            ' เชื่อม Event Handlers
            SetupEventHandlers()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเริ่มต้น: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub SetupEventHandlers()
        ' Event handlers สำหรับ Filter controls
        AddHandler txtSearch.TextChanged, AddressOf txtSearch_TextChanged
        AddHandler cmbStatus.SelectedIndexChanged, AddressOf cmbStatus_SelectedIndexChanged
        AddHandler dtpFromDate.ValueChanged, AddressOf DateFilter_Changed
        AddHandler dtpToDate.ValueChanged, AddressOf DateFilter_Changed
        
        ' Event handlers สำหรับ Button controls
        AddHandler btnRefresh.Click, AddressOf btnRefresh_Click
        AddHandler btnExport.Click, AddressOf btnExport_Click
        AddHandler btnViewDetail.Click, AddressOf btnViewDetail_Click
        AddHandler btnDelete.Click, AddressOf btnDelete_Click
        AddHandler btnExportExcel.Click, AddressOf btnExportExcel_Click
        
        ' Event handlers สำหรับ DataGridView
        AddHandler dgvHistory.SelectionChanged, AddressOf dgvHistory_SelectionChanged
        AddHandler dgvHistory.CellFormatting, AddressOf dgvHistory_CellFormatting
        AddHandler dgvHistory.DoubleClick, AddressOf dgvHistory_DoubleClick
        AddHandler dgvHistory.DataBindingComplete, AddressOf dgvHistory_DataBindingComplete
    End Sub

    Private Sub SetupDataGridView()
        Try
            dgvHistory.Columns.Clear()

            ' สร้างคอลัมน์
            Dim colDateTime As New DataGridViewTextBoxColumn() With {
                .Name = "ScanDateTime",
                .HeaderText = "วันที่/เวลา",
                .DataPropertyName = "ScanDateTime",
                .Width = 150,
                .DefaultCellStyle = New DataGridViewCellStyle() With {.Format = "dd/MM/yyyy HH:mm:ss"}
            }
            dgvHistory.Columns.Add(colDateTime)

            Dim colProductCode As New DataGridViewTextBoxColumn() With {
                .Name = "ProductCode",
                .HeaderText = "รหัสผลิตภัณฑ์",
                .DataPropertyName = "ProductCode",
                .Width = 180
            }
            dgvHistory.Columns.Add(colProductCode)

            Dim colReferenceCode As New DataGridViewTextBoxColumn() With {
                .Name = "ReferenceCode",
                .HeaderText = "รหัสอ้างอิง",
                .DataPropertyName = "ReferenceCode",
                .Width = 150
            }
            dgvHistory.Columns.Add(colReferenceCode)

            Dim colQuantity As New DataGridViewTextBoxColumn() With {
                .Name = "Quantity",
                .HeaderText = "จำนวน",
                .DataPropertyName = "Quantity",
                .Width = 80
            }
            dgvHistory.Columns.Add(colQuantity)

            Dim colDateCode As New DataGridViewTextBoxColumn() With {
                .Name = "DateCode",
                .HeaderText = "วันที่ผลิต",
                .DataPropertyName = "DateCode",
                .Width = 100
            }
            dgvHistory.Columns.Add(colDateCode)

            ' คอลัมน์สถานะ
            Dim colStatus As New DataGridViewTextBoxColumn() With {
                .Name = "Status",
                .HeaderText = "สถานะ",
                .Width = 100
            }
            dgvHistory.Columns.Add(colStatus)

            Dim colComputerName As New DataGridViewTextBoxColumn() With {
                .Name = "ComputerName",
                .HeaderText = "เครื่อง",
                .DataPropertyName = "ComputerName",
                .Width = 100
            }
            dgvHistory.Columns.Add(colComputerName)

            Dim colUserName As New DataGridViewTextBoxColumn() With {
                .Name = "UserName",
                .HeaderText = "ผู้ใช้",
                .DataPropertyName = "UserName",
                .Width = 100
            }
            dgvHistory.Columns.Add(colUserName)

            Console.WriteLine($"Created {scanHistory.Count} test records")

        Catch ex As Exception
            Console.WriteLine($"Error creating test data: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการสร้างข้อมูลทดสอบ: {ex.Message}")
        End Try
    End Sub

    Private Sub RefreshDataGridView()
        Try
            If filteredHistory Is Nothing Then
                Console.WriteLine("filteredHistory is Nothing")
                Return
            End If

            Console.WriteLine($"Refreshing DataGridView with {filteredHistory.Count} records")

            ' ใช้ BindingSource
            Dim bindingSource As New BindingSource()
            bindingSource.DataSource = filteredHistory
            dgvHistory.DataSource = bindingSource

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

            Console.WriteLine($"DataGridView refreshed. Row count: {dgvHistory.Rows.Count}")

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
    Private Sub dgvHistory_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs)
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

    Private Sub dgvHistory_SelectionChanged(sender As Object, e As EventArgs)
        Try
            Dim hasSelection As Boolean = dgvHistory.SelectedRows.Count > 0
            btnViewDetail.Enabled = hasSelection
            btnDelete.Enabled = hasSelection

        Catch ex As Exception
            Console.WriteLine($"Error in selection changed: {ex.Message}")
        End Try
    End Sub

    Private Sub dgvHistory_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
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

    Private Sub dgvHistory_DoubleClick(sender As Object, e As EventArgs)
        btnViewDetail_Click(sender, e)
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs)
        RefreshData()
    End Sub

    Private Sub btnViewDetail_Click(sender As Object, e As EventArgs)
        Try
            If dgvHistory.SelectedRows.Count = 0 Then Return

            Dim selectedRecord As ScanDataRecord = CType(dgvHistory.SelectedRows(0).DataBoundItem, ScanDataRecord)
            ShowDetailDialog(selectedRecord)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs)
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

    Private Sub btnExport_Click(sender As Object, e As EventArgs)
        Try
            ExportToCSV()
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnExportExcel_Click(sender As Object, e As EventArgs)
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
            If scanHistory Is Nothing Then Return

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

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs)
        ApplyFilters()
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As Object, e As EventArgs)
        ApplyFilters()
    End Sub

    Private Sub DateFilter_Changed(sender As Object, e As EventArgs)
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
                .MinimizeBox = False
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
            Dim saveDialog As New SaveFileDialog() With {
                .Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                .Title = "ส่งออกข้อมูลเป็น CSV",
                .FileName = $"ScanHistory_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Using writer As New IO.StreamWriter(saveDialog.FileName, False, System.Text.Encoding.UTF8)
                    ' เขียน header
                    writer.WriteLine("วันที่/เวลา,รหัสผลิตภัณฑ์,รหัสอ้างอิง,จำนวน,วันที่ผลิต,สถานะ,เครื่อง,ผู้ใช้,ข้อมูลต้นฉบับ")

                    ' เขียนข้อมูล
                    For Each record As ScanDataRecord In filteredHistory
                        Dim line As String = $"""{record.ScanDateTime:dd/MM/yyyy HH:mm:ss}"",""{record.ProductCode}"",""{record.ReferenceCode}"",""{record.Quantity}"",""{record.DateCode}"",""{If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง")}"",""{record.ComputerName}"",""{record.UserName}"",""{record.OriginalData.Replace("""", """""")}"""
                        writer.WriteLine(line)
                    Next
                End Using

                MessageBox.Show($"ส่งออกข้อมูลเรียบร้อยแล้ว" & vbNewLine & $"ไฟล์: {saveDialog.FileName}",
                              "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

End Class.WriteLine("DataGridView setup completed with " & dgvHistory.Columns.Count & " columns")

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตั้งค่า DataGridView: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Data Management"
    Private Sub LoadScanHistory()
        Try
            Console.WriteLine("Loading scan history...")

            ' ตรวจสอบการเชื่อมต่อฐานข้อมูล
            If Not DatabaseManager.IsConnected Then
                Console.WriteLine("Database not connected, attempting to initialize...")
                If Not DatabaseManager.Initialize() Then
                    MessageBox.Show("ไม่สามารถเชื่อมต่อฐานข้อมูลได้", "ข้อผิดพลาด")
                    ' สร้างข้อมูลทดสอบ
                    CreateTestData()
                    Return
                End If
            End If

            scanHistory = DatabaseManager.GetScanHistory(1000)
            Console.WriteLine($"Loaded {scanHistory.Count} records from database")

            ' ถ้าไม่มีข้อมูลจากฐานข้อมูล ให้สร้างข้อมูลทดสอบ
            If scanHistory.Count = 0 Then
                CreateTestData()
            End If

            filteredHistory = New List(Of ScanDataRecord)(scanHistory)
            RefreshDataGridView()
            UpdateRecordCount()

        Catch ex As Exception
            Console.WriteLine($"Error loading data: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ' สร้างข้อมูลทดสอบในกรณีมีปัญหา
            CreateTestData()
        End Try
    End Sub

    ''' <summary>
    ''' สร้างข้อมูลทดสอบเมื่อไม่มีข้อมูลจากฐานข้อมูล
    ''' </summary>
    Private Sub CreateTestData()
        Try
            Console.WriteLine("Creating test data...")
            scanHistory = New List(Of ScanDataRecord)()

            ' สร้างข้อมูลทดสอบ 10 รายการ
            For i As Integer = 1 To 10
                Dim testRecord As New ScanDataRecord() With {
                    .ScanDateTime = DateTime.Now.AddHours(-i),
                    .OriginalData = $"R00C-19160425501276{i}+Q000060+P20414-00770{i}A000+D20250527+LPT0000000+V00C-191604+U0000000",
                    .ExtractedData = $"20414-00770{i}A000",
                    .ProductCode = $"20414-00770{i}A000",
                    .ReferenceCode = $"00C-19160425501276{i}",
                    .Quantity = "60",
                    .DateCode = "20250527",
                    .IsValid = (i Mod 3 <> 0), ' สลับสถานะ
                    .ValidationMessages = If(i Mod 3 = 0, "ข้อมูลไม่สมบูรณ์", ""),
                    .ComputerName = Environment.MachineName,
                    .UserName = Environment.UserName
                }
                scanHistory.Add(testRecord)
            Next

            filteredHistory = New List(Of ScanDataRecord)(scanHistory)
            RefreshDataGridView()
            UpdateRecordCount()

            Console.WriteLine("Test data created successfully")
            