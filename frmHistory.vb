Imports System.Net.NetworkInformation
Imports System.Diagnostics
Imports System.IO

Public Class frmHistory

#Region "Variables"
    Private scanHistory As List(Of ScanDataRecord)
    Private filteredHistory As List(Of ScanDataRecord)
    Private isLoading As Boolean = False
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Console.WriteLine("frmHistory_Load started")

            InitializeForm()
            SetupDataGridView()
            RefreshData()
            ApplyFilters()
            'LoadScanHistory()

            Console.WriteLine("frmHistory_Load completed")

        Catch ex As Exception
            Console.WriteLine($"Error in frmHistory_Load: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดฟอร์ม: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHistory_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        Try
            ' ไม่ต้องเรียก RefreshData() ซ้ำที่นี่เพราะเรียกใน Load แล้ว
            Console.WriteLine("frmHistory_Shown completed")
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

            Console.WriteLine("InitializeForm completed")

        Catch ex As Exception
            Console.WriteLine($"Error in InitializeForm: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการเริ่มต้น: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

            dgvHistory.Columns.Add(New DataGridViewButtonColumn() With {
                .Name = "btnCheckExcel",
                .HeaderText = "ตรวจสอบไฟล์ Excel",
                .Text = "ตรวจสอบไฟล์ Excel",
                .UseColumnTextForButtonValue = True,
                .Width = 150
            })

            ' สร้างคอลัมน์วันที่/เวลา
            Dim colDateTime As New DataGridViewTextBoxColumn() With {
            .Name = "ScanDateTime",
            .HeaderText = "วันที่/เวลา",
            .DataPropertyName = "ScanDateTime",
            .Width = 150
        }
            colDateTime.DefaultCellStyle.Format = "dd/MM/yyyy HH:mm:ss"
            dgvHistory.Columns.Add(colDateTime)

            ' สร้างคอลัมน์รหัสผลิตภัณฑ์
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "ProductCode",
            .HeaderText = "รหัสผลิตภัณฑ์",
            .DataPropertyName = "ProductCode",
            .Width = 180
        })

            ' สร้างคอลัมน์รหัสอ้างอิง
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "ReferenceCode",
            .HeaderText = "รหัสอ้างอิง",
            .DataPropertyName = "ReferenceCode",
            .Width = 150
        })

            ' สร้างคอลัมน์จำนวน
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "Quantity",
            .HeaderText = "จำนวน",
            .DataPropertyName = "Quantity",
            .Width = 80
        })

            ' สร้างคอลัมน์วันที่ผลิต
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "DateCode",
            .HeaderText = "วันที่ผลิต",
            .DataPropertyName = "DateCode",
            .Width = 100
        })

            ' สร้างคอลัมน์สถานะ - ใช้ TextBox แทน Boolean เพื่อแสดงข้อความ
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "StatusDisplay",
            .HeaderText = "สถานะ",
            .Width = 100
        })

            ' สร้างคอลัมน์เครื่อง
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "ComputerName",
            .HeaderText = "เครื่อง",
            .DataPropertyName = "ComputerName",
            .Width = 100
        })

            ' สร้างคอลัมน์ผู้ใช้
            dgvHistory.Columns.Add(New DataGridViewTextBoxColumn() With {
            .Name = "UserName",
            .HeaderText = "ผู้ใช้",
            .DataPropertyName = "UserName",
            .Width = 100
        })

            Console.WriteLine($"SetupDataGridView completed with {dgvHistory.Columns.Count} columns")

        Catch ex As Exception
            Console.WriteLine($"Error in SetupDataGridView: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการตั้งค่า DataGridView: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnCheckExcel_Click(sender As Object, e As EventArgs)
        Try
            ' ดึงข้อมูลรหัสผลิตภัณฑ์จากแถวที่เลือก
            Dim productCode As String = ""
            If dgvHistory.SelectedRows.Count > 0 Then
                Dim selectedRecord As ScanDataRecord = CType(dgvHistory.SelectedRows(0).DataBoundItem, ScanDataRecord)
                productCode = selectedRecord.ProductCode
            End If

            Console.WriteLine("เริ่มตรวจสอบการเชื่อมต่อกับเซิร์ฟเวอร์")

            ' แสดงสถานะการทำงานให้ผู้ใช้ทราบ
            Dim statusForm As New Form With {
                .Text = "กำลังตรวจสอบการเชื่อมต่อ",
                .Size = New Size(300, 100),
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .StartPosition = FormStartPosition.CenterParent,
                .ControlBox = False
            }

            Dim lblStatus As New Label With {
                .Text = "กำลังตรวจสอบการเชื่อมต่อกับเซิร์ฟเวอร์...",
                .Location = New Point(20, 20),
                .AutoSize = True
            }

            statusForm.Controls.Add(lblStatus)

            ' เริ่มการตรวจสอบในเธรดแยก
            Dim pingSuccess As Boolean = False
            Dim networkType As String = ""
            Dim errorMessage As String = ""

            ' แสดงหน้าต่างสถานะ
            statusForm.Show(Me)
            Application.DoEvents()

            ' ตรวจสอบการเชื่อมต่อ
            Dim ping As New Ping()

            ' ทดสอบเครือข่าย FAB ก่อน
            Try
                Dim replyFab As PingReply = ping.Send("172.24.0.3", 2000)
                If replyFab.Status = IPStatus.Success Then
                    pingSuccess = True
                    networkType = "FAB"
                End If
            Catch ex As Exception
                Console.WriteLine($"ไม่สามารถเชื่อมต่อกับเครือข่าย FAB: {ex.Message}")
            End Try

            ' ถ้าไม่สำเร็จ ให้ลองเครือข่าย OA
            If Not pingSuccess Then
                Try
                    Dim replyOa As PingReply = ping.Send("10.24.179.2", 2000)
                    If replyOa.Status = IPStatus.Success Then
                        pingSuccess = True
                        networkType = "OA"
                    End If
                Catch ex As Exception
                    errorMessage = ex.Message
                    Console.WriteLine($"ไม่สามารถเชื่อมต่อกับเครือข่าย OA: {ex.Message}")
                End Try
            End If

            ' ปิดหน้าต่างสถานะ
            statusForm.Close()

            ' แสดงผลลัพธ์
            If pingSuccess Then
                MessageBox.Show($"เชื่อมต่อสำเร็จกับเครือข่าย {networkType}", "แจ้งเตือน",
                    MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' เปิดไฟล์ Excel ตามประเภทเครือข่าย
                Try
                    Dim excelPath As String = ""

                    If networkType = "OA" Then
                        ' กำหนด path สำหรับเครือข่าย OA
                        excelPath = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

                        ' ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        If IO.File.Exists(excelPath) Then
                            MessageBox.Show($"พบไฟล์ Excel ที่ต้องการ:{Environment.NewLine}{excelPath}",
                                          "ตรวจสอบไฟล์ Excel", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            ' เปิดไฟล์ Excel และค้นหาข้อมูล
                            If Not String.IsNullOrEmpty(productCode) Then
                                ' ถามผู้ใช้ว่าต้องการเปิดไฟล์ Excel หรือไม่
                                Dim result = MessageBox.Show(
                                    $"ต้องการเปิดไฟล์ Excel เพื่อค้นหารหัสผลิตภัณฑ์ '{productCode}' หรือไม่?{Environment.NewLine}" &
                                    $"คลิก Yes เพื่อเปิดไฟล์ และค้นหาด้วยตนเอง (ใช้ Ctrl+F)",
                                    "ยืนยันการเปิดไฟล์ Excel",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question)

                                If result = DialogResult.Yes Then
                                    ' เปิดไฟล์ Excel ด้วย Process.Start
                                    Process.Start(excelPath)

                                    ' แสดงข้อความแนะนำวิธีค้นหา
                                    MessageBox.Show(
                                        $"เมื่อไฟล์ Excel เปิดขึ้นมาแล้ว ให้กด Ctrl+F เพื่อค้นหารหัสผลิตภัณฑ์: {productCode}" &
                                        $"{Environment.NewLine}โดยมักจะอยู่ในคอลัมน์ C ของ Sheet1",
                                        "วิธีค้นหาข้อมูล",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information)
                                End If
                            Else
                                ' ถ้าไม่มีการเลือกแถว ให้เพียงเปิดไฟล์ Excel
                                Process.Start(excelPath)
                            End If
                        Else
                            MessageBox.Show($"ไม่พบไฟล์ Excel ที่ต้องการ:{Environment.NewLine}{excelPath}",
                                          "ตรวจสอบไฟล์ Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    ElseIf networkType = "FAB" Then
                        ' ถ้าเป็นเครือข่าย FAB ให้แจ้งว่าไม่สามารถเข้าถึงไฟล์ได้
                        MessageBox.Show("เครือข่าย FAB ไม่สามารถเข้าถึงไฟล์ Excel ได้ กรุณาเชื่อมต่อกับเครือข่าย OA",
                                      "ตรวจสอบไฟล์ Excel", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                Catch ex As Exception
                    MessageBox.Show($"เกิดข้อผิดพลาดในการเปิดไฟล์ Excel: {ex.Message}",
                                  "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Console.WriteLine($"Error opening Excel file: {ex.Message}")
                End Try
            Else
                MessageBox.Show("ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้" &
                    If(Not String.IsNullOrEmpty(errorMessage), vbCrLf & "สาเหตุ: " & errorMessage, ""),
                    "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตรวจสอบการเชื่อมต่อ: {ex.Message}",
                "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in btnCheckExcel_Click: {ex.Message}")
        End Try
    End Sub

    Private Sub dgvHistory_CellClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvHistory.CellClick
        Try
            ' ตรวจสอบว่าคลิกที่คอลัมน์ปุ่ม "ตรวจสอบไฟล์ Excel" หรือไม่
            If e.RowIndex >= 0 AndAlso e.ColumnIndex = 0 AndAlso dgvHistory.Columns(e.ColumnIndex).Name = "btnCheckExcel" Then
                ' เลือกแถวที่คลิก
                dgvHistory.Rows(e.RowIndex).Selected = True

                ' เรียกฟังก์ชันตรวจสอบการเชื่อมต่อและค้นหาข้อมูล
                btnCheckExcel_Click(sender, e)
            End If
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาด: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in dgvHistory_CellClick: {ex.Message}")
        End Try
    End Sub

    ' เพิ่มเมธอดสำหรับจัดการข้อผิดพลาดของ DataGridView
    Private Sub dgvHistory_DataError(sender As Object, e As DataGridViewDataErrorEventArgs) Handles dgvHistory.DataError
        Try
            Console.WriteLine($"DataGridView Error: {e.Exception.Message}")

            ' ป้องกันไม่ให้แสดง error dialog
            e.Cancel = True

            ' ถ้าเป็นปัญหาเรื่อง format ให้ตั้งค่าเป็นค่าว่าง
            If TypeOf e.Exception Is FormatException Then
                dgvHistory.Rows(e.RowIndex).Cells(e.ColumnIndex).Value = ""
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in DataError handler: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Data Management"
    Private Sub LoadScanHistory()
        Try
            Console.WriteLine("LoadScanHistory started")

            isLoading = True
            toolStripProgressBar.Visible = True
            toolStripProgressBar.Style = ProgressBarStyle.Marquee

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
            Console.WriteLine($"Loaded {If(scanHistory?.Count, 0)} records from database")

            ' ถ้าไม่มีข้อมูลจากฐานข้อมูล ให้สร้างข้อมูลทดสอบ
            If scanHistory Is Nothing OrElse scanHistory.Count = 0 Then
                Console.WriteLine("No data from database, creating test data...")
                CreateTestData()
            Else
                Console.WriteLine("ApplyFilters started")

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

                Console.WriteLine($"ApplyFilters: {filteredHistory.Count} records after filtering")

                RefreshDataGridView()
                UpdateRecordCount()
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in LoadScanHistory: {ex.Message}")
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
            Console.WriteLine("CreateTestData started")

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

            Console.WriteLine($"Created {scanHistory.Count} test records")

            ' กำหนดข้อมูลที่จะแสดง
            filteredHistory = New List(Of ScanDataRecord)(scanHistory)

            ' รีเฟรช DataGridView
            RefreshDataGridView()
            UpdateRecordCount()

        Catch ex As Exception
            Console.WriteLine($"Error in CreateTestData: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการสร้างข้อมูลทดสอบ: {ex.Message}")
        End Try
    End Sub

    Private Sub RefreshDataGridView()
        Try
            If isLoading Then
                Console.WriteLine("RefreshDataGridView: Still loading, skipping...")
                Return
            End If

            If filteredHistory Is Nothing Then
                Console.WriteLine("RefreshDataGridView: filteredHistory is Nothing")
                Return
            End If

            Console.WriteLine($"RefreshDataGridView: Binding {filteredHistory.Count} records to DataGridView")

            ' ล้าง DataSource เดิม
            dgvHistory.DataSource = Nothing

            ' ตั้งค่าข้อมูลใหม่
            If filteredHistory.Count > 0 Then
                dgvHistory.DataSource = filteredHistory
                Console.WriteLine($"DataGridView bound successfully. Rows count: {dgvHistory.Rows.Count}")
            Else
                Console.WriteLine("No filtered data to display")
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in RefreshDataGridView: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการรีเฟรชข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshData()
        LoadScanHistory()
    End Sub

    Private Sub UpdateRecordCount()
        Try
            Dim filteredCount As Integer = If(filteredHistory?.Count, 0)
            Dim totalCount As Integer = If(scanHistory?.Count, 0)

            lblCount.Text = $"จำนวนรายการ: {filteredCount} จาก {totalCount} รายการทั้งหมด"
            Console.WriteLine($"UpdateRecordCount: {lblCount.Text}")

        Catch ex As Exception
            Console.WriteLine($"Error in UpdateRecordCount: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Event Handlers"
    Private Sub dgvHistory_DataBindingComplete(sender As Object, e As DataGridViewBindingCompleteEventArgs) Handles dgvHistory.DataBindingComplete
        Try
            Console.WriteLine($"DataBindingComplete: Rows={dgvHistory.Rows.Count}, Columns={dgvHistory.Columns.Count}")

            ' อัปเดตคอลัมน์สถานะ
            For Each row As DataGridViewRow In dgvHistory.Rows
                If row.DataBoundItem IsNot Nothing Then
                    Dim record As ScanDataRecord = CType(row.DataBoundItem, ScanDataRecord)

                    ' อัปเดตคอลัมน์สถานะ
                    If dgvHistory.Columns.Contains("IsValid") Then
                        Dim statusText As String = If(record.IsValid, "✅ ถูกต้อง", "❌ ไม่ถูกต้อง")
                        row.Cells("IsValid").Value = statusText
                        row.Cells("IsValid").Style.ForeColor = If(record.IsValid, Color.Green, Color.Red)
                    End If

                    ' เปลี่ยนสีแถวตามสถานะ
                    If Not record.IsValid Then
                        row.DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235)
                    End If
                End If
            Next

            Console.WriteLine("DataBindingComplete: Status column updated")

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
            Console.WriteLine($"Error in SelectionChanged: {ex.Message}")
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

            Console.WriteLine("ApplyFilters started")

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

            Console.WriteLine($"ApplyFilters: {filteredHistory.Count} records after filtering")

            RefreshDataGridView()
            UpdateRecordCount()

        Catch ex As Exception
            Console.WriteLine($"Error in ApplyFilters: {ex.Message}")
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