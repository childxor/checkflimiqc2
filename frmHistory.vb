Public Class frmHistory

#Region "Variables"
    Private scanHistory As List(Of ScanDataRecord)
    Private filteredHistory As List(Of ScanDataRecord)
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        InitializeForm()
        LoadScanHistory()
        SetupDataGridView()
    End Sub

    Private Sub frmHistory_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        RefreshData()
    End Sub
#End Region

#Region "Initialization"
    Private Sub InitializeForm()
        Try
            Me.Text = "ประวัติการสแกน QR Code"
            Me.Size = New Size(1200, 700)
            Me.StartPosition = FormStartPosition.CenterParent
            Me.WindowState = FormWindowState.Normal

            ' สร้าง controls
            CreateControls()
            SetupLayout()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการเริ่มต้น: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub CreateControls()
        ' Panel หลัก
        Dim pnlMain As New Panel() With {
            .Name = "pnlMain",
            .Dock = DockStyle.Fill,
            .Padding = New Padding(10)
        }
        Me.Controls.Add(pnlMain)

        ' Panel สำหรับการค้นหาและกรอง
        Dim pnlFilter As New Panel() With {
            .Name = "pnlFilter",
            .Height = 80,
            .Dock = DockStyle.Top,
            .BackColor = Color.FromArgb(248, 249, 250)
        }
        pnlMain.Controls.Add(pnlFilter)

        ' กล่องค้นหา
        Dim lblSearch As New Label() With {
            .Text = "ค้นหา:",
            .Location = New Point(10, 15),
            .Size = New Size(50, 23),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        pnlFilter.Controls.Add(lblSearch)

        Dim txtSearch As New TextBox() With {
            .Name = "txtSearch",
            .Location = New Point(70, 12),
            .Size = New Size(300, 23),
            .PlaceholderText = "ค้นหาด้วยรหัสผลิตภัณฑ์, รหัสอ้างอิง หรือข้อมูล..."
        }
        AddHandler txtSearch.TextChanged, AddressOf txtSearch_TextChanged
        pnlFilter.Controls.Add(txtSearch)

        ' ComboBox สำหรับกรองสถานะ
        Dim lblStatus As New Label() With {
            .Text = "สถานะ:",
            .Location = New Point(390, 15),
            .Size = New Size(50, 23),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        pnlFilter.Controls.Add(lblStatus)

        Dim cmbStatus As New ComboBox() With {
            .Name = "cmbStatus",
            .Location = New Point(450, 12),
            .Size = New Size(120, 23),
            .DropDownStyle = ComboBoxStyle.DropDownList
        }
        cmbStatus.Items.AddRange({"ทั้งหมด", "ถูกต้อง", "ไม่ถูกต้อง"})
        cmbStatus.SelectedIndex = 0
        AddHandler cmbStatus.SelectedIndexChanged, AddressOf cmbStatus_SelectedIndexChanged
        pnlFilter.Controls.Add(cmbStatus)

        ' DateTimePicker สำหรับกรองวันที่
        Dim lblFromDate As New Label() With {
            .Text = "จาก:",
            .Location = New Point(590, 15),
            .Size = New Size(30, 23),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        pnlFilter.Controls.Add(lblFromDate)

        Dim dtpFromDate As New DateTimePicker() With {
            .Name = "dtpFromDate",
            .Location = New Point(630, 12),
            .Size = New Size(120, 23),
            .Value = DateTime.Now.AddDays(-7)
        }
        AddHandler dtpFromDate.ValueChanged, AddressOf DateFilter_Changed
        pnlFilter.Controls.Add(dtpFromDate)

        Dim lblToDate As New Label() With {
            .Text = "ถึง:",
            .Location = New Point(760, 15),
            .Size = New Size(30, 23),
            .TextAlign = ContentAlignment.MiddleLeft
        }
        pnlFilter.Controls.Add(lblToDate)

        Dim dtpToDate As New DateTimePicker() With {
            .Name = "dtpToDate",
            .Location = New Point(800, 12),
            .Size = New Size(120, 23),
            .Value = DateTime.Now
        }
        AddHandler dtpToDate.ValueChanged, AddressOf DateFilter_Changed
        pnlFilter.Controls.Add(dtpToDate)

        ' ปุ่มรีเฟรช
        Dim btnRefresh As New Button() With {
            .Name = "btnRefresh",
            .Text = "🔄 รีเฟรช",
            .Location = New Point(940, 10),
            .Size = New Size(80, 27),
            .UseVisualStyleBackColor = True
        }
        AddHandler btnRefresh.Click, AddressOf btnRefresh_Click
        pnlFilter.Controls.Add(btnRefresh)

        ' ปุ่มส่งออก
        Dim btnExport As New Button() With {
            .Name = "btnExport",
            .Text = "📤 ส่งออก",
            .Location = New Point(1030, 10),
            .Size = New Size(80, 27),
            .UseVisualStyleBackColor = True
        }
        AddHandler btnExport.Click, AddressOf btnExport_Click
        pnlFilter.Controls.Add(btnExport)

        ' Label สำหรับแสดงจำนวนรายการ
        Dim lblCount As New Label() With {
            .Name = "lblCount",
            .Text = "จำนวนรายการ: 0",
            .Location = New Point(10, 45),
            .Size = New Size(200, 20),
            .ForeColor = Color.Gray
        }
        pnlFilter.Controls.Add(lblCount)

        ' DataGridView สำหรับแสดงข้อมูล
        Dim dgvHistory As New DataGridView() With {
            .Name = "dgvHistory",
            .Dock = DockStyle.Fill,
            .AutoGenerateColumns = False,
            .ReadOnly = True,
            .AllowUserToAddRows = False,
            .AllowUserToDeleteRows = False,
            .SelectionMode = DataGridViewSelectionMode.FullRowSelect,
            .MultiSelect = False,
            .RowHeadersVisible = False,
            .AlternatingRowsDefaultCellStyle = New DataGridViewCellStyle() With {.BackColor = Color.FromArgb(248, 249, 250)}
        }
        pnlMain.Controls.Add(dgvHistory)

        ' Panel สำหรับปุ่มด้านล่าง
        Dim pnlButtons As New Panel() With {
            .Name = "pnlButtons",
            .Height = 50,
            .Dock = DockStyle.Bottom,
            .BackColor = SystemColors.Control
        }
        pnlMain.Controls.Add(pnlButtons)

        ' ปุ่มดูรายละเอียด
        Dim btnViewDetail As New Button() With {
            .Name = "btnViewDetail",
            .Text = "ดูรายละเอียด",
            .Location = New Point(10, 10),
            .Size = New Size(100, 30),
            .Enabled = False
        }
        AddHandler btnViewDetail.Click, AddressOf btnViewDetail_Click
        pnlButtons.Controls.Add(btnViewDetail)

        ' ปุ่มลบรายการ
        Dim btnDelete As New Button() With {
            .Name = "btnDelete",
            .Text = "ลบรายการ",
            .Location = New Point(120, 10),
            .Size = New Size(100, 30),
            .Enabled = False,
            .BackColor = Color.FromArgb(231, 76, 60),
            .ForeColor = Color.White,
            .FlatStyle = FlatStyle.Flat
        }
        AddHandler btnDelete.Click, AddressOf btnDelete_Click
        pnlButtons.Controls.Add(btnDelete)

        ' ปุ่มปิด
        Dim btnClose As New Button() With {
            .Name = "btnClose",
            .Text = "ปิด",
            .Location = New Point(Me.Width - 120, 10),
            .Size = New Size(80, 30),
            .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
            .DialogResult = DialogResult.OK
        }
        pnlButtons.Controls.Add(btnClose)
    End Sub

    Private Sub SetupLayout()
        ' จัดการ layout responsive
        AddHandler Me.Resize, AddressOf frmHistory_Resize
    End Sub

    Private Sub SetupDataGridView()
        Try
            Dim dgv As DataGridView = CType(Me.Controls.Find("dgvHistory", True).FirstOrDefault(), DataGridView)
            If dgv Is Nothing Then Return

            dgv.Columns.Clear()

            ' สร้างคอลัมน์
            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "ScanDateTime",
                .HeaderText = "วันที่/เวลา",
                .DataPropertyName = "ScanDateTime",
                .Width = 120,
                .DefaultCellStyle = New DataGridViewCellStyle() With {.Format = "dd/MM/yyyy HH:mm:ss"}
            })

            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "ProductCode",
                .HeaderText = "รหัสผลิตภัณฑ์",
                .DataPropertyName = "ProductCode",
                .Width = 150
            })

            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "ReferenceCode",
                .HeaderText = "รหัสอ้างอิง",
                .DataPropertyName = "ReferenceCode",
                .Width = 120
            })

            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "Quantity",
                .HeaderText = "จำนวน",
                .DataPropertyName = "Quantity",
                .Width = 80
            })

            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "DateCode",
                .HeaderText = "วันที่ผลิต",
                .DataPropertyName = "DateCode",
                .Width = 100
            })

            ' คอลัมน์สถานะแบบกราฟิก
            Dim statusColumn As New DataGridViewImageColumn() With {
                .Name = "StatusIcon",
                .HeaderText = "สถานะ",
                .Width = 60
            }
            dgv.Columns.Add(statusColumn)

            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "ComputerName",
                .HeaderText = "เครื่อง",
                .DataPropertyName = "ComputerName",
                .Width = 100
            })

            dgv.Columns.Add(New DataGridViewTextBoxColumn() With {
                .Name = "UserName",
                .HeaderText = "ผู้ใช้",
                .DataPropertyName = "UserName",
                .Width = 100
            })

            ' เพิ่ม event handlers
            AddHandler dgv.SelectionChanged, AddressOf dgvHistory_SelectionChanged
            AddHandler dgv.CellFormatting, AddressOf dgvHistory_CellFormatting
            AddHandler dgv.DoubleClick, AddressOf dgvHistory_DoubleClick

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตั้งค่า DataGridView: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

#Region "Data Management"
    Private Sub LoadScanHistory()
        Try
            scanHistory = DatabaseManager.GetScanHistory(1000)
            filteredHistory = New List(Of ScanDataRecord)(scanHistory)
            RefreshDataGridView()
            UpdateRecordCount()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshDataGridView()
        Try
            Dim dgv As DataGridView = CType(Me.Controls.Find("dgvHistory", True).FirstOrDefault(), DataGridView)
            If dgv Is Nothing Then Return

            dgv.DataSource = Nothing
            dgv.DataSource = filteredHistory

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการรีเฟรชข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub RefreshData()
        LoadScanHistory()
    End Sub

    Private Sub UpdateRecordCount()
        Try
            Dim lblCount As Label = CType(Me.Controls.Find("lblCount", True).FirstOrDefault(), Label)
            If lblCount IsNot Nothing Then
                lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count} จาก {scanHistory.Count} รายการทั้งหมด"
            End If
        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub
#End Region

#Region "Filter and Search"
    Private Sub ApplyFilters()
        Try
            If scanHistory Is Nothing Then Return

            filteredHistory = New List(Of ScanDataRecord)(scanHistory)

            ' กรองตามข้อความค้นหา
            Dim txtSearch As TextBox = CType(Me.Controls.Find("txtSearch", True).FirstOrDefault(), TextBox)
            If txtSearch IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtSearch.Text) Then
                Dim searchText As String = txtSearch.Text.ToLower()
                filteredHistory = filteredHistory.Where(Function(x)
                                                           Return x.ProductCode.ToLower().Contains(searchText) OrElse
                                                                  x.ReferenceCode.ToLower().Contains(searchText) OrElse
                                                                  x.OriginalData.ToLower().Contains(searchText) OrElse
                                                                  x.ExtractedData.ToLower().Contains(searchText)
                                                       End Function).ToList()
            End If

            ' กรองตามสถานะ
            Dim cmbStatus As ComboBox = CType(Me.Controls.Find("cmbStatus", True).FirstOrDefault(), ComboBox)
            If cmbStatus IsNot Nothing AndAlso cmbStatus.SelectedIndex > 0 Then
                Dim isValid As Boolean = (cmbStatus.SelectedIndex = 1)
                filteredHistory = filteredHistory.Where(Function(x) x.IsValid = isValid).ToList()
            End If

            ' กรองตามช่วงวันที่
            Dim dtpFromDate As DateTimePicker = CType(Me.Controls.Find("dtpFromDate", True).FirstOrDefault(), DateTimePicker)
            Dim dtpToDate As DateTimePicker = CType(Me.Controls.Find("dtpToDate", True).FirstOrDefault(), DateTimePicker)
            
            If dtpFromDate IsNot Nothing AndAlso dtpToDate IsNot Nothing Then
                Dim fromDate As DateTime = dtpFromDate.Value.Date
                Dim toDate As DateTime = dtpToDate.Value.Date.AddDays(1).AddSeconds(-1)
                
                filteredHistory = filteredHistory.Where(Function(x)
                                                           Return x.ScanDateTime >= fromDate AndAlso x.ScanDateTime <= toDate
                                                       End Function).ToList()
            End If

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

#Region "Event Handlers"
    Private Sub dgvHistory_SelectionChanged(sender As Object, e As EventArgs)
        Try
            Dim dgv As DataGridView = CType(sender, DataGridView)
            Dim hasSelection As Boolean = dgv.SelectedRows.Count > 0

            Dim btnViewDetail As Button = CType(Me.Controls.Find("btnViewDetail", True).FirstOrDefault(), Button)
            Dim btnDelete As Button = CType(Me.Controls.Find("btnDelete", True).FirstOrDefault(), Button)

            If btnViewDetail IsNot Nothing Then btnViewDetail.Enabled = hasSelection
            If btnDelete IsNot Nothing Then btnDelete.Enabled = hasSelection

        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub

    Private Sub dgvHistory_CellFormatting(sender As Object, e As DataGridViewCellFormattingEventArgs)
        Try
            Dim dgv As DataGridView = CType(sender, DataGridView)
            
            If e.ColumnIndex = dgv.Columns("StatusIcon").Index AndAlso e.RowIndex >= 0 Then
                Dim record As ScanDataRecord = CType(dgv.Rows(e.RowIndex).DataBoundItem, ScanDataRecord)
                
                ' สร้างไอคอนสถานะ
                Dim icon As Bitmap = CreateStatusIcon(record.IsValid)
                e.Value = icon
            End If

            ' เปลี่ยนสีแถวตามสถานะ
            If e.RowIndex >= 0 Then
                Dim record As ScanDataRecord = CType(dgv.Rows(e.RowIndex).DataBoundItem, ScanDataRecord)
                If Not record.IsValid Then
                    dgv.Rows(e.RowIndex).DefaultCellStyle.BackColor = Color.FromArgb(255, 235, 235)
                End If
            End If

        Catch
            ' ไม่ต้องทำอะไร
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
            Dim dgv As DataGridView = CType(Me.Controls.Find("dgvHistory", True).FirstOrDefault(), DataGridView)
            If dgv Is Nothing OrElse dgv.SelectedRows.Count = 0 Then Return

            Dim selectedRecord As ScanDataRecord = CType(dgv.SelectedRows(0).DataBoundItem, ScanDataRecord)
            ShowDetailDialog(selectedRecord)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnDelete_Click(sender As Object, e As EventArgs)
        Try
            Dim dgv As DataGridView = CType(Me.Controls.Find("dgvHistory", True).FirstOrDefault(), DataGridView)
            If dgv Is Nothing OrElse dgv.SelectedRows.Count = 0 Then Return

            Dim result As DialogResult = MessageBox.Show(
                "คุณต้องการลบรายการนี้หรือไม่?",
                "ยืนยันการลบ",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                ' TODO: เพิ่มฟังก์ชันลบข้อมูลในฐานข้อมูล
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

    Private Sub frmHistory_Resize(sender As Object, e As EventArgs)
        Try
            ' ปรับขนาดปุ่มปิดให้อยู่มุมขวา
            Dim btnClose As Button = CType(Me.Controls.Find("btnClose", True).FirstOrDefault(), Button)
            If btnClose IsNot Nothing Then
                btnClose.Location = New Point(Me.Width - 120, btnClose.Location.Y)
            End If
        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub
#End Region

#Region "Utility Methods"
    Private Function CreateStatusIcon(isValid As Boolean) As Bitmap
        Try
            Dim icon As New Bitmap(16, 16)
            Using g As Graphics = Graphics.FromImage(icon)
                g.SmoothingMode = Drawing2D.SmoothingMode.AntiAlias
                
                If isValid Then
                    ' สีเขียว สำหรับสถานะถูกต้อง
                    Using brush As New SolidBrush(Color.FromArgb(46, 125, 50))
                        g.FillEllipse(brush, 2, 2, 12, 12)
                    End Using
                    
                    ' เครื่องหมายถูก
                    Using pen As New Pen(Color.White, 2)
                        g.DrawLines(pen, {New Point(5, 8), New Point(7, 10), New Point(11, 6)})
                    End Using
                Else
                    ' สีแดง สำหรับสถานะไม่ถูกต้อง
                    Using brush As New SolidBrush(Color.FromArgb(211, 47, 47))
                        g.FillEllipse(brush, 2, 2, 12, 12)
                    End Using
                    
                    ' เครื่องหมาย X
                    Using pen As New Pen(Color.White, 2)
                        g.DrawLine(pen, 5, 5, 11, 11)
                        g.DrawLine(pen, 11, 5, 5, 11)
                    End Using
                End If
            End Using
            
            Return icon
            
        Catch
            ' Return simple colored square if graphics creation fails
            Dim fallback As New Bitmap(16, 16)
            Using g As Graphics = Graphics.FromImage(fallback)
                Dim color As Color = If(isValid, Color.Green, Color.Red)
                g.Clear(color)
            End Using
            Return fallback
        End Try
    End Function

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
            Dim pnlButtons As New Panel() With {
                .Height = 50,
                .Dock = DockStyle.Bottom
            }

            Dim btnClose As New Button() With {
                .Text = "ปิด",
                .Size = New Size(75, 30),
                .Location = New Point(detailForm.Width - 95, 10),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                .DialogResult = DialogResult.OK
            }

            pnlButtons.Controls.Add(btnClose)
            detailForm.Controls.Add(pnlButtons)
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

End Class