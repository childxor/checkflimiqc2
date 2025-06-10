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

    ' เปลี่ยนฟังก์ชัน btnCheckExcel_Click ในไฟล์ frmHistory.vb
    Private Sub btnCheckExcel_Click(sender As Object, e As EventArgs)
        Try
            ' ดึงข้อมูลรหัสผลิตภัณฑ์จากแถวที่เลือก
            Dim productCode As String = ""
            Dim selectedRecord As ScanDataRecord = Nothing

            If dgvHistory.SelectedRows.Count > 0 Then
                selectedRecord = CType(dgvHistory.SelectedRows(0).DataBoundItem, ScanDataRecord)
                productCode = selectedRecord.ProductCode
            End If

            If String.IsNullOrEmpty(productCode) Then
                MessageBox.Show("กรุณาเลือกรายการที่มีรหัสผลิตภัณฑ์", "แจ้งเตือน",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            Console.WriteLine("เริ่มตรวจสอบการเชื่อมต่อกับเซิร์ฟเวอร์")

            ' แสดงสถานะการทำงานให้ผู้ใช้ทราบ
            Dim statusForm As New Form With {
                .Text = "กำลังตรวจสอบการเชื่อมต่อ",
                .Size = New Size(400, 120),
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .StartPosition = FormStartPosition.CenterParent,
                .ControlBox = False
            }

            Dim lblStatus As New Label With {
                .Text = "กำลังตรวจสอบการเชื่อมต่อกับเซิร์ฟเวอร์...",
                .Location = New Point(20, 20),
                .Size = New Size(360, 20),
                .TextAlign = ContentAlignment.MiddleCenter
            }

            Dim progressBar As New ProgressBar With {
                .Location = New Point(20, 50),
                .Size = New Size(360, 23),
                .Style = ProgressBarStyle.Marquee
            }

            statusForm.Controls.Add(lblStatus)
            statusForm.Controls.Add(progressBar)

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
            lblStatus.Text = "กำลังทดสอบเครือข่าย FAB..."
            Application.DoEvents()

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
                lblStatus.Text = "กำลังทดสอบเครือข่าย OA..."
                Application.DoEvents()

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
                ' เปิดไฟล์ Excel ตามประเภทเครือข่าย
                Try
                    Dim excelPath As String = ""

                    If networkType = "OA" Then
                        ' กำหนด path สำหรับเครือข่าย OA
                        excelPath = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

                        ' ตรวจสอบว่าไฟล์มีอยู่จริงหรือไม่
                        If IO.File.Exists(excelPath) Then
                            MessageBox.Show($"เชื่อมต่อสำเร็จกับเครือข่าย {networkType}" & vbNewLine &
                                          $"พบไฟล์ Excel: {IO.Path.GetFileName(excelPath)}" & vbNewLine &
                                          $"กำลังค้นหารหัสผลิตภัณฑ์: {productCode}",
                                          "เชื่อมต่อสำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)

                            ' ค้นหาข้อมูลในไฟล์ Excel
                            Dim searchResult As ExcelSearchResult = SearchExcelFile(excelPath, productCode)

                            If searchResult.HasError Then
                                MessageBox.Show(searchResult.ErrorMessage, "ข้อผิดพลาด",
                                              MessageBoxButtons.OK, MessageBoxIcon.Error)
                            ElseIf searchResult.HasMatches Then
                                ' แสดงผลการค้นหา
                                ShowExcelSearchResults(searchResult, selectedRecord)

                                ' ถามผู้ใช้ว่าต้องการเปิดไฟล์ Excel หรือไม่
                                Dim result = MessageBox.Show(
                                    $"พบข้อมูลในไฟล์ Excel แล้ว!" & vbNewLine &
                                    $"ต้องการเปิดไฟล์ Excel เพื่อดูข้อมูลเพิ่มเติมหรือไม่?",
                                    "เปิดไฟล์ Excel",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question)

                                If result = DialogResult.Yes Then
                                    Process.Start(excelPath)
                                End If
                            Else
                                MessageBox.Show(searchResult.SummaryMessage, "ไม่พบข้อมูล",
                                              MessageBoxButtons.OK, MessageBoxIcon.Information)

                                ' ถามผู้ใช้ว่าต้องการเปิดไฟล์ Excel เพื่อตรวจสอบด้วยตนเองหรือไม่
                                Dim result = MessageBox.Show(
                                    $"ไม่พบรหัสผลิตภัณฑ์ '{productCode}' ในฐานข้อมูล" & vbNewLine &
                                    $"ต้องการเปิดไฟล์ Excel เพื่อค้นหาด้วยตนเองหรือไม่?",
                                    "ค้นหาด้วยตนเอง",
                                    MessageBoxButtons.YesNo,
                                    MessageBoxIcon.Question)

                                If result = DialogResult.Yes Then
                                    Process.Start(excelPath)
                                    MessageBox.Show(
                                        $"เมื่อไฟล์ Excel เปิดขึ้นมาแล้ว ให้กด Ctrl+F เพื่อค้นหารหัสผลิตภัณฑ์: {productCode}" &
                                        $"{Environment.NewLine}โดยมักจะอยู่ในคอลัมน์ C ของ Sheet1",
                                        "วิธีค้นหาข้อมูล",
                                        MessageBoxButtons.OK,
                                        MessageBoxIcon.Information)
                                End If
                            End If
                        Else
                            MessageBox.Show($"เชื่อมต่อกับเครือข่าย {networkType} สำเร็จ" & vbNewLine &
                                          $"แต่ไม่พบไฟล์ Excel ที่: {excelPath}",
                                          "ไม่พบไฟล์", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        End If
                    ElseIf networkType = "FAB" Then
                        ' ถ้าเป็นเครือข่าย FAB ให้แจ้งว่าไม่สามารถเข้าถึงไฟล์ได้
                        MessageBox.Show($"เชื่อมต่อกับเครือข่าย {networkType} สำเร็จ" & vbNewLine &
                                      "แต่เครือข่าย FAB ไม่สามารถเข้าถึงไฟล์ Excel ได้" & vbNewLine &
                                      "กรุณาเชื่อมต่อกับเครือข่าย OA",
                                      "ไม่สามารถเข้าถึงไฟล์", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                    End If
                Catch ex As Exception
                    MessageBox.Show($"เกิดข้อผิดพลาดในการเปิดไฟล์ Excel: {ex.Message}",
                                  "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    Console.WriteLine($"Error opening Excel file: {ex.Message}")
                End Try
            Else
                MessageBox.Show("ไม่สามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้" &
                    If(Not String.IsNullOrEmpty(errorMessage), vbCrLf & "สาเหตุ: " & errorMessage, ""),
                    "ไม่สามารถเชื่อมต่อ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตรวจสอบการเชื่อมต่อ: {ex.Message}",
                "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in btnCheckExcel_Click: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' แสดงผลการค้นหาใน Excel
    ''' </summary>
    Private Sub ShowExcelSearchResults(searchResult As ExcelSearchResult, selectedRecord As ScanDataRecord)
        Try
            ' สร้างฟอร์มแสดงผลการค้นหา
            Dim resultForm As New Form() With {
                .Text = "ผลการค้นหาใน Excel Database",
                .Size = New Size(700, 500),
                .StartPosition = FormStartPosition.CenterParent,
                .FormBorderStyle = FormBorderStyle.Sizable,
                .MinimumSize = New Size(600, 400),
                .Font = New Font("Segoe UI", 9)
            }

            ' สร้าง header panel
            Dim headerPanel As New Panel() With {
                .Height = 80,
                .Dock = DockStyle.Top,
                .BackColor = Color.FromArgb(41, 128, 185)
            }

            Dim lblHeader As New Label() With {
                .Text = "🔍 ผลการค้นหาใน Excel Database",
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 12, FontStyle.Bold),
                .Location = New Point(15, 15),
                .AutoSize = True
            }

            Dim lblSubHeader As New Label() With {
                .Text = $"รหัสผลิตภัณฑ์: {searchResult.SearchedProductCode}",
                .ForeColor = Color.White,
                .Font = New Font("Segoe UI", 10),
                .Location = New Point(15, 40),
                .AutoSize = True
            }

            headerPanel.Controls.Add(lblHeader)
            headerPanel.Controls.Add(lblSubHeader)

            ' สร้าง main panel
            Dim mainPanel As New Panel() With {
                .Dock = DockStyle.Fill,
                .Padding = New Padding(15)
            }

            ' สร้าง TextBox แสดงผลลัพธ์
            Dim txtResult As New TextBox() With {
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .ReadOnly = True,
                .Dock = DockStyle.Fill,
                .Font = New Font("Consolas", 10),
                .BackColor = Color.White,
                .BorderStyle = BorderStyle.FixedSingle
            }

            ' สร้างข้อความแสดงผล
            Dim resultText As New System.Text.StringBuilder()

            resultText.AppendLine("=== ผลการค้นหาใน Excel Database ===")
            resultText.AppendLine()
            resultText.AppendLine($"ไฟล์: {IO.Path.GetFileName(searchResult.ExcelFilePath)}")
            resultText.AppendLine($"รหัสผลิตภัณฑ์ที่ค้นหา: {searchResult.SearchedProductCode}")
            resultText.AppendLine($"จำนวนที่พบ: {searchResult.MatchCount} รายการ")
            resultText.AppendLine()

            If searchResult.HasMatches Then
                resultText.AppendLine("📋 รายละเอียดที่พบ:")
                resultText.AppendLine(New String("-", 60))

                For i As Integer = 0 To searchResult.Matches.Count - 1
                    Dim match As ExcelMatchResult = searchResult.Matches(i)

                    resultText.AppendLine($"รายการที่ {i + 1}:")
                    resultText.AppendLine($"  แถวที่: {match.RowNumber}")
                    resultText.AppendLine($"  รหัสผลิตภัณฑ์: {match.ProductCode}")
                    resultText.AppendLine($"  ข้อมูลคอลัมน์ที่ 4: {match.Column4Value}")

                    If Not String.IsNullOrEmpty(match.Column1Value) Then
                        resultText.AppendLine($"  Item: {match.Column1Value}")
                    End If

                    If Not String.IsNullOrEmpty(match.Column2Value) Then
                        resultText.AppendLine($"  LITEON FG PN: {match.Column2Value}")
                    End If

                    If Not String.IsNullOrEmpty(match.Column5Value) Then
                        resultText.AppendLine($"  LEGEND: {match.Column5Value}")
                    End If

                    If Not String.IsNullOrEmpty(match.Column6Value) Then
                        resultText.AppendLine($"  LAYOUT: {match.Column6Value}")
                    End If

                    resultText.AppendLine()
                Next

                ' แสดงข้อมูลการสแกนเปรียบเทียบ
                resultText.AppendLine()
                resultText.AppendLine("📊 ข้อมูลการสแกนปัจจุบัน:")
                resultText.AppendLine(New String("-", 60))
                resultText.AppendLine($"  เวลาสแกน: {selectedRecord.ScanDateTime:dd/MM/yyyy HH:mm:ss}")
                resultText.AppendLine($"  รหัสผลิตภัณฑ์: {selectedRecord.ProductCode}")
                resultText.AppendLine($"  รหัสอ้างอิง: {selectedRecord.ReferenceCode}")
                resultText.AppendLine($"  จำนวน: {selectedRecord.Quantity}")
                resultText.AppendLine($"  วันที่ผลิต: {selectedRecord.DateCode}")
                resultText.AppendLine($"  สถานะ: {If(selectedRecord.IsValid, "✅ ถูกต้อง", "❌ ไม่ถูกต้อง")}")

                ' เพิ่มข้อมูลเปรียบเทียบ
                resultText.AppendLine()
                resultText.AppendLine("🔍 การเปรียบเทียบ:")
                resultText.AppendLine(New String("-", 60))

                Dim firstMatch As ExcelMatchResult = searchResult.FirstMatch
                If Not String.IsNullOrEmpty(firstMatch.Column4Value) Then
                    resultText.AppendLine($"✅ พบข้อมูลในฐานข้อมูล Excel")
                    resultText.AppendLine($"   ข้อมูลคอลัมน์ที่ 4: {firstMatch.Column4Value}")

                    If Not String.IsNullOrEmpty(firstMatch.Column5Value) Then
                        resultText.AppendLine($"   LEGEND: {firstMatch.Column5Value}")
                    End If

                    If searchResult.MatchCount > 1 Then
                        resultText.AppendLine($"⚠️  มีข้อมูลซ้ำกัน {searchResult.MatchCount} รายการ")
                    End If
                Else
                    resultText.AppendLine($"❌ ไม่พบข้อมูลในฐานข้อมูล Excel")
                End If
            End If

            txtResult.Text = resultText.ToString()

            ' สร้าง button panel
            Dim buttonPanel As New Panel() With {
                .Height = 60,
                .Dock = DockStyle.Bottom,
                .Padding = New Padding(15, 10, 15, 10)
            }

            Dim btnOpenExcel As New Button() With {
                .Text = "📊 เปิด Excel",
                .Size = New Size(120, 35),
                .Location = New Point(15, 12),
                .BackColor = Color.FromArgb(46, 125, 50),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold)
            }
            btnOpenExcel.FlatAppearance.BorderSize = 0

            Dim btnCopyResult As New Button() With {
                .Text = "📋 คัดลอก",
                .Size = New Size(100, 35),
                .Location = New Point(145, 12),
                .BackColor = Color.FromArgb(52, 152, 219),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold)
            }
            btnCopyResult.FlatAppearance.BorderSize = 0

            Dim btnClose As New Button() With {
                .Text = "❌ ปิด",
                .Size = New Size(80, 35),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                .Location = New Point(resultForm.Width - 110, 12),
                .BackColor = Color.FromArgb(108, 117, 125),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold),
                .DialogResult = DialogResult.OK
            }
            btnClose.FlatAppearance.BorderSize = 0

            ' Event handlers สำหรับปุ่ม
            AddHandler btnOpenExcel.Click, Sub()
                                               Try
                                                   Process.Start(searchResult.ExcelFilePath)
                                               Catch ex As Exception
                                                   MessageBox.Show($"ไม่สามารถเปิดไฟล์ Excel ได้: {ex.Message}",
                                                                 "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                               End Try
                                           End Sub

            AddHandler btnCopyResult.Click, Sub()
                                                Try
                                                    Clipboard.SetText(txtResult.Text)
                                                    MessageBox.Show("คัดลอกข้อมูลเรียบร้อยแล้ว", "สำเร็จ",
                                                                  MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                Catch ex As Exception
                                                    MessageBox.Show($"ไม่สามารถคัดลอกข้อมูลได้: {ex.Message}",
                                                                  "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                End Try
                                            End Sub

            ' จัดการเมื่อ form ถูก resize
            AddHandler resultForm.Resize, Sub()
                                              btnClose.Location = New Point(resultForm.Width - 110, 12)
                                          End Sub

            buttonPanel.Controls.Add(btnOpenExcel)
            buttonPanel.Controls.Add(btnCopyResult)
            buttonPanel.Controls.Add(btnClose)

            mainPanel.Controls.Add(txtResult)

            resultForm.Controls.Add(mainPanel)
            resultForm.Controls.Add(buttonPanel)
            resultForm.Controls.Add(headerPanel)

            ' แสดง form
            resultForm.ShowDialog()
            resultForm.Dispose()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงผลการค้นหา: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' เพิ่มเมธอดเหล่านี้ในคลาส frmHistory

    ''' <summary>
    ''' สร้างข้อความสรุปสำหรับแสดงใน MessageBox แบบสั้น
    ''' </summary>
    Private Function CreateQuickSummary(searchResult As ExcelSearchResult) As String
        If searchResult.HasMatches Then
            Dim firstMatch = searchResult.FirstMatch
            Return $"🎯 พบข้อมูลแล้ว!" & vbNewLine &
                   $"รหัสผลิตภัณฑ์: {firstMatch.ProductCode}" & vbNewLine &
                   $"ข้อมูลคอลัมน์ที่ 4: {firstMatch.Column4Value}" & vbNewLine &
                   $"แถวที่: {firstMatch.RowNumber}" &
                   If(searchResult.MatchCount > 1, $" (และอีก {searchResult.MatchCount - 1} รายการ)", "")
        Else
            Return $"❌ ไม่พบข้อมูล" & vbNewLine &
                   $"รหัสผลิตภัณฑ์: {searchResult.SearchedProductCode}" & vbNewLine &
                   "กรุณาตรวจสอบรหัสผลิตภัณฑ์หรือค้นหาด้วยตนเองใน Excel"
        End If
    End Function

    ''' <summary>
    ''' ทดสอบการค้นหา Excel ด้วยข้อมูลจำลอง
    ''' </summary>
    Private Sub TestExcelSearchFunction()
        Try
            ' สร้างข้อมูลทดสอบ
            Dim testResult As New ExcelSearchResult() With {
                .IsSuccess = True,
                .SearchedProductCode = "20414-095200A002",
                .ExcelFilePath = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx",
                .MatchCount = 1
            }

            Dim testMatch As New ExcelMatchResult() With {
                .RowNumber = 2,
                .ProductCode = "20414-095200A002",
                .Column4Value = "SN1B63L101XU-01N",
                .Column1Value = "SG-C1010-XUA",
                .Column2Value = "LITEON FG PN painting keycaps part no.",
                .Column5Value = "US",
                .Column6Value = "SN1B63B42",
                .IsExactMatch = True
            }

            testResult.Matches = New List(Of ExcelMatchResult) From {testMatch}
            testResult.FirstMatch = testMatch
            testResult.SummaryMessage = CreateQuickSummary(testResult)

            ' สร้างข้อมูล ScanDataRecord จำลอง
            Dim testScanRecord As New ScanDataRecord() With {
                .ScanDateTime = DateTime.Now,
                .ProductCode = "20414-095200A002",
                .ReferenceCode = "00C-191604255012766",
                .Quantity = "60",
                .DateCode = "20250527",
                .IsValid = True,
                .ComputerName = Environment.MachineName,
                .UserName = Environment.UserName
            }

            ' แสดงผลการทดสอบ
            ShowExcelSearchResults(testResult, testScanRecord)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการทดสอบ: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ปุ่มทดสอบการค้นหา Excel (สามารถเพิ่มใน form designer หรือเรียกจากเมนู Debug)
    ''' </summary>
    Private Sub btnTestExcelSearch_Click(sender As Object, e As EventArgs)
        TestExcelSearchFunction()
    End Sub

    ''' <summary>
    ''' เมธอดสำหรับแปลงข้อมูล Excel เป็น DataTable (ใช้เมื่อต้องการแสดงผลในรูปแบบตาราง)
    ''' </summary>
    Private Function ConvertExcelMatchesToDataTable(matches As List(Of ExcelMatchResult)) As DataTable
        Try
            Dim dt As New DataTable()

            ' สร้างคอลัมน์
            dt.Columns.Add("แถว", GetType(Integer))
            dt.Columns.Add("Item", GetType(String))
            dt.Columns.Add("LITEON FG PN", GetType(String))
            dt.Columns.Add("รหัสผลิตภัณฑ์", GetType(String))
            dt.Columns.Add("ข้อมูลคอลัมน์ที่ 4", GetType(String))
            dt.Columns.Add("LEGEND", GetType(String))
            dt.Columns.Add("LAYOUT", GetType(String))

            ' เพิ่มข้อมูล
            For Each match In matches
                Dim row As DataRow = dt.NewRow()
                row("แถว") = match.RowNumber
                row("Item") = match.Column1Value
                row("LITEON FG PN") = match.Column2Value
                row("รหัสผลิตภัณฑ์") = match.ProductCode
                row("ข้อมูลคอลัมน์ที่ 4") = match.Column4Value
                row("LEGEND") = match.Column5Value
                row("LAYOUT") = match.Column6Value
                dt.Rows.Add(row)
            Next

            Return dt

        Catch ex As Exception
            Console.WriteLine($"Error creating DataTable: {ex.Message}")
            Return New DataTable()
        End Try
    End Function

    ''' <summary>
    ''' สร้างรายงาน Excel Search เป็นไฟล์ CSV
    ''' </summary>
    Private Sub ExportExcelSearchResults(searchResult As ExcelSearchResult, scanRecord As ScanDataRecord)
        Try
            If Not searchResult.HasMatches Then
                MessageBox.Show("ไม่มีข้อมูลที่พบสำหรับส่งออก", "แจ้งเตือน",
                              MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim saveDialog As New SaveFileDialog() With {
                .Filter = "CSV Files (*.csv)|*.csv|All Files (*.*)|*.*",
                .DefaultExt = "csv",
                .FileName = $"ExcelSearchResult_{searchResult.SearchedProductCode}_{DateTime.Now:yyyyMMdd_HHmmss}.csv"
            }

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Using writer As New StreamWriter(saveDialog.FileName, False, System.Text.Encoding.UTF8)
                    ' Header
                    writer.WriteLine("รหัสผลิตภัณฑ์ที่ค้นหา,แถวที่พบ,Item,LITEON FG PN,รหัสผลิตภัณฑ์,ข้อมูลคอลัมน์ที่ 4,LEGEND,LAYOUT,วันที่ค้นหา")

                    ' Data
                    For Each match In searchResult.Matches
                        Dim line As String = $"""{searchResult.SearchedProductCode}"",""{match.RowNumber}"",""{match.Column1Value}"",""{match.Column2Value}"",""{match.ProductCode}"",""{match.Column4Value}"",""{match.Column5Value}"",""{match.Column6Value}"",""{DateTime.Now:yyyy-MM-dd HH:mm:ss}"""
                        writer.WriteLine(line)
                    Next
                End Using

                MessageBox.Show($"ส่งออกผลการค้นหาเรียบร้อยแล้ว" & vbNewLine & $"ไฟล์: {saveDialog.FileName}",
                              "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ' เพิ่มโค้ดนี้ในไฟล์ frmHistory.vb

#Region "Debug and Test Features"
    ''' <summary>
    ''' เพิ่ม Context Menu สำหรับ DataGridView (เรียกใน InitializeForm)
    ''' </summary>
    Private Sub SetupContextMenu()
        Try
            Dim contextMenu As New ContextMenuStrip()

            ' รายการเมนู
            Dim menuViewDetail As New ToolStripMenuItem("ดูรายละเอียด", Nothing, AddressOf btnViewDetail_Click)
            Dim menuSeparator1 As New ToolStripSeparator()
            Dim menuCheckExcel As New ToolStripMenuItem("🔍 ตรวจสอบไฟล์ Excel", Nothing, AddressOf btnCheckExcel_Click)
            Dim menuSeparator2 As New ToolStripSeparator()
            Dim menuTestExcel As New ToolStripMenuItem("🧪 ทดสอบ Excel Search", Nothing, AddressOf TestExcelSearchWithSampleData)
            Dim menuSystemStatus As New ToolStripMenuItem("📊 สถานะระบบ", Nothing, AddressOf ShowSystemStatus)

            ' เพิ่มรายการเมนู
            contextMenu.Items.Add(menuViewDetail)
            contextMenu.Items.Add(menuSeparator1)
            contextMenu.Items.Add(menuCheckExcel)
            contextMenu.Items.Add(menuSeparator2)
            contextMenu.Items.Add(menuTestExcel)
            contextMenu.Items.Add(menuSystemStatus)

            ' กำหนด Context Menu ให้กับ DataGridView
            dgvHistory.ContextMenuStrip = contextMenu

        Catch ex As Exception
            Console.WriteLine($"Error setting up context menu: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' เพิ่มปุ่มทดสอบในแถบปุ่ม (เรียกใน InitializeForm เมื่อต้องการ)
    ''' </summary>
    Private Sub AddTestButtons()
        Try
            ' ปุ่มทดสอบ Excel
            Dim btnTestExcel As New Button() With {
                .Text = "🧪 ทดสอบ",
                .Size = New Size(80, 35),
                .BackColor = Color.FromArgb(156, 39, 176),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold),
                .Location = New Point(485, 15)
            }
            btnTestExcel.FlatAppearance.BorderSize = 0
            AddHandler btnTestExcel.Click, AddressOf TestExcelSearchWithSampleData

            ' ปุ่มสถานะระบบ
            Dim btnStatus As New Button() With {
                .Text = "📊 สถานะ",
                .Size = New Size(80, 35),
                .BackColor = Color.FromArgb(96, 125, 139),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 9, FontStyle.Bold),
                .Location = New Point(575, 15)
            }
            btnStatus.FlatAppearance.BorderSize = 0
            AddHandler btnStatus.Click, AddressOf ShowSystemStatus

            ' เพิ่มปุ่มไปยัง panel (ถ้าต้องการ)
            If Application.OpenForms("frmHistory") IsNot Nothing Then
                ' เพิ่มเฉพาะในโหมด Debug
#If DEBUG Then
                pnlButtons.Controls.Add(btnTestExcel)
                pnlButtons.Controls.Add(btnStatus)
#End If
            End If

        Catch ex As Exception
            Console.WriteLine($"Error adding test buttons: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Keyboard shortcuts สำหรับฟีเจอร์ทดสอบ
    ''' </summary>
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Try
            Select Case keyData
                Case Keys.F5
                    ' F5 = Refresh
                    RefreshData()
                    Return True

                Case Keys.F9
                    ' F9 = Test Excel Search
                    TestExcelSearchWithSampleData()
                    Return True

                Case Keys.F10
                    ' F10 = System Status
                    ShowSystemStatus()
                    Return True

                Case Keys.F12
                    ' F12 = Check Excel for selected item
                    If dgvHistory.SelectedRows.Count > 0 Then
                        btnCheckExcel_Click(Nothing, Nothing)
                    End If
                    Return True

                Case Keys.Control Or Keys.T
                    ' Ctrl+T = Test mode
                    ToggleTestMode()
                    Return True
            End Select

        Catch ex As Exception
            Console.WriteLine($"Error in ProcessCmdKey: {ex.Message}")
        End Try

        Return MyBase.ProcessCmdKey(msg, keyData)
    End Function

    ''' <summary>
    ''' สลับโหมดทดสอบ
    ''' </summary>
    Private Sub ToggleTestMode()
        Try
            Static testModeEnabled As Boolean = False
            testModeEnabled = Not testModeEnabled

            If testModeEnabled Then
                Me.Text += " [TEST MODE]"
                Me.BackColor = Color.FromArgb(255, 245, 238) ' สีส้มอ่อน
                MessageBox.Show("เปิดโหมดทดสอบ" & vbNewLine &
                              "F5 = Refresh, F9 = Test Excel, F10 = Status, F12 = Check Excel" & vbNewLine &
                              "Ctrl+T = ปิดโหมดทดสอบ",
                              "โหมดทดสอบ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                Me.Text = Me.Text.Replace(" [TEST MODE]", "")
                Me.BackColor = Color.White
                MessageBox.Show("ปิดโหมดทดสอบแล้ว", "โหมดปกติ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาด: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' แสดงหน้าต่างช่วยเหลือ
    ''' </summary>
    Private Sub ShowHelpDialog()
        Try
            Dim helpText As String = "=== คู่มือการใช้งาน ===" & vbNewLine & vbNewLine &
                                   "🔍 การค้นหาใน Excel:" & vbNewLine &
                                   "1. เลือกรายการในตาราง" & vbNewLine &
                                   "2. คลิกปุ่ม 'ตรวจสอบไฟล์ Excel' หรือกด F12" & vbNewLine &
                                   "3. ระบบจะค้นหารหัสผลิตภัณฑ์ใน Excel Database" & vbNewLine & vbNewLine &
                                   "⌨️ คีย์ลัด:" & vbNewLine &
                                   "F5 = รีเฟรชข้อมูล" & vbNewLine &
                                   "F9 = ทดสอบ Excel Search" & vbNewLine &
                                   "F10 = ตรวจสอบสถานะระบบ" & vbNewLine &
                                   "F12 = ตรวจสอบ Excel สำหรับรายการที่เลือก" & vbNewLine &
                                   "Ctrl+T = เปิด/ปิดโหมดทดสอบ" & vbNewLine & vbNewLine &
                                   "🌐 ข้อกำหนดเครือข่าย:" & vbNewLine &
                                   "- ต้องเชื่อมต่อเครือข่าย OA เพื่อเข้าถึงไฟล์ Excel" & vbNewLine &
                                   "- ต้องติดตั้ง Microsoft Office Excel" & vbNewLine & vbNewLine &
                                   "📂 ไฟล์ Excel:" & vbNewLine &
                                   "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx" & vbNewLine &
                                   "คอลัมน์ที่ 3 = รหัสผลิตภัณฑ์" & vbNewLine &
                                   "คอลัมน์ที่ 4 = ข้อมูลที่ต้องการ"

            MessageBox.Show(helpText, "คู่มือการใช้งาน", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาด: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' อัปเดต InitializeForm ให้เรียกใช้ Context Menu
    ''' </summary>
    Private Sub InitializeFormWithContextMenu()
        Try
            ' เรียกฟังก์ชันเดิม
            InitializeForm()

            ' เพิ่ม Context Menu
            SetupContextMenu()

            ' เพิ่มปุ่มทดสอบ (เฉพาะในโหมด Debug)
#If DEBUG Then
            AddTestButtons()
#End If

            ' เพิ่ม Help button
            AddHelpButton()

        Catch ex As Exception
            Console.WriteLine($"Error in InitializeFormWithContextMenu: {ex.Message}")
            ' ถ้าเกิดข้อผิดพลาด ให้เรียกฟังก์ชันเดิม
            InitializeForm()
        End Try
    End Sub

    ''' <summary>
    ''' เพิ่มปุ่ม Help
    ''' </summary>
    Private Sub AddHelpButton()
        Try
            Dim btnHelp As New Button() With {
                .Text = "❓",
                .Size = New Size(35, 35),
                .BackColor = Color.FromArgb(52, 152, 219),
                .ForeColor = Color.White,
                .FlatStyle = FlatStyle.Flat,
                .Font = New Font("Segoe UI", 12, FontStyle.Bold),
                .Anchor = AnchorStyles.Top Or AnchorStyles.Right,
                .Location = New Point(Me.Width - 50, 15)
            }
            btnHelp.FlatAppearance.BorderSize = 0
            AddHandler btnHelp.Click, AddressOf ShowHelpDialog

            ' เพิ่มปุ่มไปยังพาเนลปุ่ม
            pnlButtons.Controls.Add(btnHelp)

            ' จัดการเมื่อฟอร์มถูก resize
            AddHandler Me.Resize, Sub()
                                      btnHelp.Location = New Point(Me.Width - 60, 15)
                                  End Sub

        Catch ex As Exception
            Console.WriteLine($"Error adding help button: {ex.Message}")
        End Try
    End Sub
#End Region

#Region "Excel Search Integration"
    ''' <summary>
    ''' เพิ่มคอลัมน์สำหรับแสดงสถานะการค้นหา Excel
    ''' </summary>
    Private Sub AddExcelStatusColumn()
        Try
            ' เพิ่มคอลัมน์สถานะ Excel หลังจากคอลัมน์อื่นๆ
            Dim colExcelStatus As New DataGridViewTextBoxColumn() With {
                .Name = "ExcelStatus",
                .HeaderText = "สถานะ Excel",
                .Width = 100,
                .ReadOnly = True
            }

            dgvHistory.Columns.Add(colExcelStatus)

        Catch ex As Exception
            Console.WriteLine($"Error adding Excel status column: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' อัปเดตสถานะ Excel สำหรับแถวที่เลือก
    ''' </summary>
    Private Sub UpdateExcelStatusForRow(rowIndex As Integer, status As String)
        Try
            If rowIndex >= 0 AndAlso rowIndex < dgvHistory.Rows.Count Then
                If dgvHistory.Columns.Contains("ExcelStatus") Then
                    dgvHistory.Rows(rowIndex).Cells("ExcelStatus").Value = status

                    ' เปลี่ยนสีตามสถานะ
                    Select Case status.ToLower()
                        Case "พบข้อมูล", "found"
                            dgvHistory.Rows(rowIndex).Cells("ExcelStatus").Style.ForeColor = Color.Green
                        Case "ไม่พบ", "not found"
                            dgvHistory.Rows(rowIndex).Cells("ExcelStatus").Style.ForeColor = Color.Red
                        Case "กำลังค้นหา", "searching"
                            dgvHistory.Rows(rowIndex).Cells("ExcelStatus").Style.ForeColor = Color.Orange
                        Case Else
                            dgvHistory.Rows(rowIndex).Cells("ExcelStatus").Style.ForeColor = Color.Gray
                    End Select
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error updating Excel status: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ค้นหา Excel สำหรับรายการทั้งหมดที่แสดงอยู่ (Batch Search)
    ''' </summary>
    Private Sub BatchSearchExcel()
        Try
            If filteredHistory Is Nothing OrElse filteredHistory.Count = 0 Then
                MessageBox.Show("ไม่มีข้อมูลสำหรับค้นหา", "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            Dim result As DialogResult = MessageBox.Show(
                $"ต้องการค้นหาข้อมูลใน Excel สำหรับทุกรายการ ({filteredHistory.Count} รายการ) หรือไม่?" & vbNewLine &
                "การค้นหาจำนวนมากอาจใช้เวลานาน",
                "ยืนยันการค้นหาแบบกลุ่ม",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question)

            If result = DialogResult.Yes Then
                PerformBatchExcelSearch()
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาด: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ดำเนินการค้นหา Excel แบบกลุ่ม
    ''' </summary>
    Private Sub PerformBatchExcelSearch()
        Try
            Dim excelPath As String = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

            ' ตรวจสอบการเชื่อมต่อก่อน
            If Not File.Exists(excelPath) Then
                MessageBox.Show("ไม่สามารถเข้าถึงไฟล์ Excel ได้", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            ' แสดง Progress Form
            Dim progressForm As New Form() With {
                .Text = "กำลังค้นหาใน Excel",
                .Size = New Size(400, 150),
                .FormBorderStyle = FormBorderStyle.FixedDialog,
                .StartPosition = FormStartPosition.CenterParent,
                .ControlBox = False
            }

            Dim lblProgress As New Label() With {
                .Text = "เริ่มต้นการค้นหา...",
                .Location = New Point(20, 20),
                .Size = New Size(360, 20),
                .TextAlign = ContentAlignment.MiddleCenter
            }

            Dim progressBar As New ProgressBar() With {
                .Location = New Point(20, 50),
                .Size = New Size(360, 23),
                .Maximum = filteredHistory.Count
            }

            Dim btnCancel As New Button() With {
                .Text = "ยกเลิก",
                .Location = New Point(160, 85),
                .Size = New Size(80, 30),
                .DialogResult = DialogResult.Cancel
            }

            progressForm.Controls.Add(lblProgress)
            progressForm.Controls.Add(progressBar)
            progressForm.Controls.Add(btnCancel)

            ' แสดง Progress Form
            progressForm.Show(Me)

            Dim foundCount As Integer = 0
            Dim notFoundCount As Integer = 0

            For i As Integer = 0 To filteredHistory.Count - 1
                Application.DoEvents()

                ' ตรวจสอบการยกเลิก
                If progressForm.DialogResult = DialogResult.Cancel Then
                    Exit For
                End If

                Dim record As ScanDataRecord = filteredHistory(i)
                lblProgress.Text = $"กำลังค้นหา: {record.ProductCode} ({i + 1}/{filteredHistory.Count})"
                progressBar.Value = i + 1

                ' ค้นหาใน Excel
                Dim searchResult As ExcelSearchResult = SearchExcelFile(excelPath, record.ProductCode)

                ' อัปเดตสถานะ
                If searchResult.HasMatches Then
                    UpdateExcelStatusForRow(i, "พบข้อมูล")
                    foundCount += 1
                Else
                    UpdateExcelStatusForRow(i, "ไม่พบ")
                    notFoundCount += 1
                End If

                ' หน่วงเวลาเล็กน้อยเพื่อป้องกัน overload
                System.Threading.Thread.Sleep(100)
            Next

            progressForm.Close()

            ' แสดงผลสรุป
            MessageBox.Show($"ค้นหาเสร็จสิ้น" & vbNewLine &
                          $"พบข้อมูล: {foundCount} รายการ" & vbNewLine &
                          $"ไม่พบข้อมูล: {notFoundCount} รายการ",
                          "ผลการค้นหา", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการค้นหาแบบกลุ่ม: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub
#End Region

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

    ' SearchExcelFile
    Public Function SearchExcelFile(excelPath As String, productCode As String) As ExcelSearchResult
        Dim result As New ExcelSearchResult()
        result.SearchedProductCode = productCode
        result.ExcelFilePath = excelPath

        Dim excelApp As Microsoft.Office.Interop.Excel.Application = Nothing
        Dim workbook As Microsoft.Office.Interop.Excel.Workbook = Nothing
        Dim worksheet As Microsoft.Office.Interop.Excel.Worksheet = Nothing

        Try
            Console.WriteLine($"กำลังค้นหา '{productCode}' ในไฟล์ Excel: {excelPath}")

            ' ตรวจสอบว่าไฟล์มีอยู่จริง
            If Not File.Exists(excelPath) Then
                result.ErrorMessage = $"ไม่พบไฟล์ Excel: {excelPath}"
                result.IsSuccess = False
                Return result
            End If

            ' เริ่มต้น Excel Application
            excelApp = New Microsoft.Office.Interop.Excel.Application()
            excelApp.Visible = False
            excelApp.DisplayAlerts = False

            ' เปิดไฟล์ Excel
            workbook = excelApp.Workbooks.Open(excelPath, ReadOnly:=True)

            ' ดึง Sheet1 (หรือ sheet แรก)
            worksheet = CType(workbook.Worksheets(1), Microsoft.Office.Interop.Excel.Worksheet)

            ' หา range ที่มีข้อมูล
            Dim usedRange As Microsoft.Office.Interop.Excel.Range = worksheet.UsedRange
            Dim rowCount As Integer = usedRange.Rows.Count
            Dim colCount As Integer = usedRange.Columns.Count

            Console.WriteLine($"พบข้อมูลใน Excel: {rowCount} แถว, {colCount} คอลัมน์")

            ' ตรวจสอบว่ามีคอลัมน์ที่ 3 และ 4
            If colCount < 4 Then
                result.ErrorMessage = "ไฟล์ Excel ไม่มีคอลัมน์ที่ 4 (ต้องการอย่างน้อย 4 คอลัมน์)"
                result.IsSuccess = False
                Return result
            End If

            ' ค้นหาในคอลัมน์ที่ 3 (รหัสผลิตภัณฑ์)
            Dim foundRow As Integer = -1
            Dim searchResults As New List(Of ExcelMatchResult)()

            For row As Integer = 1 To rowCount
                Try
                    ' อ่านค่าจากคอลัมน์ที่ 3
                    Dim cellValue As Object = CType(worksheet.Cells(row, 3), Microsoft.Office.Interop.Excel.Range).Value

                    If cellValue IsNot Nothing Then
                        Dim cellText As String = cellValue.ToString().Trim()

                        ' ตรวจสอบการแมทช์
                        If cellText.Equals(productCode, StringComparison.OrdinalIgnoreCase) Then
                            ' พบข้อมูลที่ตรงกัน
                            foundRow = row

                            ' อ่านข้อมูลจากคอลัมน์ที่ 4
                            Dim column4Value As Object = CType(worksheet.Cells(row, 4), Microsoft.Office.Interop.Excel.Range).Value
                            Dim column4Text As String = If(column4Value?.ToString(), "")

                            ' สร้างผลการค้นหา
                            Dim matchResult As New ExcelMatchResult() With {
                                .RowNumber = row,
                                .ProductCode = cellText,
                                .Column4Value = column4Text,
                                .IsExactMatch = True
                            }

                            ' อ่านข้อมูลจากคอลัมน์อื่นๆ เพื่อแสดงข้อมูลเพิ่มเติม
                            Try
                                If colCount >= 1 Then
                                    Dim col1Value As Object = CType(worksheet.Cells(row, 1), Microsoft.Office.Interop.Excel.Range).Value
                                    matchResult.Column1Value = If(col1Value?.ToString(), "")
                                End If

                                If colCount >= 2 Then
                                    Dim col2Value As Object = CType(worksheet.Cells(row, 2), Microsoft.Office.Interop.Excel.Range).Value
                                    matchResult.Column2Value = If(col2Value?.ToString(), "")
                                End If

                                If colCount >= 5 Then
                                    Dim col5Value As Object = CType(worksheet.Cells(row, 5), Microsoft.Office.Interop.Excel.Range).Value
                                    matchResult.Column5Value = If(col5Value?.ToString(), "")
                                End If

                                If colCount >= 6 Then
                                    Dim col6Value As Object = CType(worksheet.Cells(row, 6), Microsoft.Office.Interop.Excel.Range).Value
                                    matchResult.Column6Value = If(col6Value?.ToString(), "")
                                End If
                            Catch ex As Exception
                                Console.WriteLine($"เกิดข้อผิดพลาดในการอ่านคอลัมน์เพิ่มเติม: {ex.Message}")
                            End Try

                            searchResults.Add(matchResult)

                            Console.WriteLine($"พบข้อมูลที่แถว {row}: {cellText} -> {column4Text}")
                        End If
                    End If

                Catch ex As Exception
                    ' ข้าม error ในการอ่านแถวนี้
                    Console.WriteLine($"ข้าม error ในแถว {row}: {ex.Message}")
                    Continue For
                End Try
            Next

            ' ตั้งค่าผลลัพธ์
            If searchResults.Count > 0 Then
                result.IsSuccess = True
                result.MatchCount = searchResults.Count
                result.Matches = searchResults
                result.FirstMatch = searchResults(0)

                ' สร้างข้อความสรุป
                If searchResults.Count = 1 Then
                    result.SummaryMessage = $"พบรหัสผลิตภัณฑ์ '{productCode}' ที่แถว {searchResults(0).RowNumber}" & vbNewLine &
                                          $"ข้อมูลคอลัมน์ที่ 4: {searchResults(0).Column4Value}"
                Else
                    result.SummaryMessage = $"พบรหัสผลิตภัณฑ์ '{productCode}' จำนวน {searchResults.Count} แถว:" & vbNewLine
                    For Each match In searchResults
                        result.SummaryMessage += $"- แถว {match.RowNumber}: {match.Column4Value}" & vbNewLine
                    Next
                End If
            Else
                result.IsSuccess = False
                result.MatchCount = 0
                result.SummaryMessage = $"ไม่พบรหัสผลิตภัณฑ์ '{productCode}' ในไฟล์ Excel"
            End If

        Catch ex As Exception
            result.IsSuccess = False
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหาไฟล์ Excel: {ex.Message}"
            Console.WriteLine($"Error in SearchExcelFile: {ex.Message}")

        Finally
            ' ปิดและเคลียร์ Excel objects
            Try
                If worksheet IsNot Nothing Then
                    Marshal.ReleaseComObject(worksheet)
                    worksheet = Nothing
                End If

                If workbook IsNot Nothing Then
                    workbook.Close(False)
                    Marshal.ReleaseComObject(workbook)
                    workbook = Nothing
                End If

                If excelApp IsNot Nothing Then
                    excelApp.Quit()
                    Marshal.ReleaseComObject(excelApp)
                    excelApp = Nothing
                End If

                ' บังคับ Garbage Collection เพื่อเคลียร์ COM objects
                GC.Collect()
                GC.WaitForPendingFinalizers()

            Catch ex As Exception
                Console.WriteLine($"Error closing Excel: {ex.Message}")
            End Try
        End Try

        Return result
    End Function

    ''' <summary>
    ''' คลาสสำหรับเก็บผลการค้นหาใน Excel
    ''' </summary>
    Public Class ExcelSearchResult
        Public Property IsSuccess As Boolean = False
        Public Property SearchedProductCode As String = ""
        Public Property ExcelFilePath As String = ""
        Public Property MatchCount As Integer = 0
        Public Property Matches As List(Of ExcelMatchResult)
        Public Property FirstMatch As ExcelMatchResult
        Public Property SummaryMessage As String = ""
        Public Property ErrorMessage As String = ""

        Public ReadOnly Property HasMatches As Boolean
            Get
                Return IsSuccess AndAlso MatchCount > 0
            End Get
        End Property

        Public ReadOnly Property HasError As Boolean
            Get
                Return Not String.IsNullOrEmpty(ErrorMessage)
            End Get
        End Property
    End Class

    ''' <summary>
    ''' คลาสสำหรับเก็บข้อมูลแต่ละแถวที่พบ
    ''' </summary>
    Public Class ExcelMatchResult
        Public Property RowNumber As Integer
        Public Property ProductCode As String = ""
        Public Property Column4Value As String = ""
        Public Property IsExactMatch As Boolean = False

        ' คอลัมน์เพิ่มเติมสำหรับแสดงข้อมูลครบถ้วน
        Public Property Column1Value As String = ""  ' Item
        Public Property Column2Value As String = ""  ' LITEON FG PN
        Public Property Column5Value As String = ""  ' LEGEND
        Public Property Column6Value As String = ""  ' LAYOUT

        Public ReadOnly Property FullRowData As String
            Get
                Return $"แถว {RowNumber}: {Column1Value} | {Column2Value} | {ProductCode} | {Column4Value} | {Column5Value} | {Column6Value}"
            End Get
        End Property
    End Class

    ''' <summary>
    ''' ฟังก์ชันสำหรับทดสอบการค้นหา Excel
    ''' </summary>
    Public Sub TestExcelSearch()
        Try
            ' ทดสอบการค้นหา
            Dim testProductCode As String = "20414-095200A002"
            Dim testExcelPath As String = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

            Console.WriteLine($"ทดสอบการค้นหา: {testProductCode}")

            Dim result As ExcelSearchResult = SearchExcelFile(testExcelPath, testProductCode)

            If result.HasError Then
                Console.WriteLine($"เกิดข้อผิดพลาด: {result.ErrorMessage}")
                MessageBox.Show(result.ErrorMessage, "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            ElseIf result.HasMatches Then
                Console.WriteLine($"ผลการค้นหา: {result.SummaryMessage}")
                MessageBox.Show(result.SummaryMessage, "ผลการค้นหา", MessageBoxButtons.OK, MessageBoxIcon.Information)

                ' แสดงรายละเอียดเพิ่มเติม
                For Each match In result.Matches
                    Console.WriteLine(match.FullRowData)
                Next
            Else
                Console.WriteLine("ไม่พบข้อมูล")
                MessageBox.Show(result.SummaryMessage, "ไม่พบข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in TestExcelSearch: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการทดสอบ: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
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