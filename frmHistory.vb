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
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Console.WriteLine("frmHistory_Load started")

            InitializeForm()
            SetupDataGridView()
            'RefreshData()
            'ApplyFilters()
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

            'Dim lblStatus As New Label With {
            '    .Text = "กำลังตรวจสอบการเชื่อมต่อกับเซิร์ฟเวอร์...",
            '    .Location = New Point(20, 20),
            '    .AutoSize = True
            '}

            statusForm.Controls.Add(lblStatus)

            ' เริ่มการตรวจสอบในเธรดแยก
            Dim pingSuccess As Boolean = False
            Dim networkType As String = ""
            Dim errorMessage As String = ""

            ' แสดงหน้าต่างสถานะ
            statusForm.Show(Me)
            'Application.DoEvents()

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
                                'SearchExcelFile(excelPath, productCode)
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
End Class