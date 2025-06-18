Imports System.Net.NetworkInformation
Imports System.Diagnostics
Imports System.IO
Imports System.Runtime.InteropServices
Imports Excel = Microsoft.Office.Interop.Excel
Imports System.Text
Imports System.Data 
Imports System.Threading
Imports System.ComponentModel
Imports System.Windows.Forms
Imports System.Text.RegularExpressions

Public Class frmHistory

#Region "Variables"
    Private scanHistory As List(Of ScanDataRecord)
    Private filteredHistory As List(Of ScanDataRecord)
    Private isLoading As Boolean = False
    Private backgroundWorker As System.ComponentModel.BackgroundWorker
    Private dataCache As ExcelDataCache
    Private excelFilePath As String = "\\fls951\OAFAB\OA2FAB\Film charecter check\Database.xlsx"
#End Region

#Region "Form Events"
    Private Sub frmHistory_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            Console.WriteLine("frmHistory_Load started")

            InitializeForm()
            SetupDataGridView()
            SetupBackgroundWorker()
            CheckDatabaseConnection()

            ' เริ่มต้น Excel Data Cache
            InitializeExcelCache()

            LoadScanHistory()
            
            ' อัปเดตชื่อโปรแกรมด้วยเวอร์ชันจาก Assembly
            UpdateFormTitleWithVersion()

            Console.WriteLine("frmHistory_Load completed")

        Catch ex As Exception
            Console.WriteLine($"Error in frmHistory_Load: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการโหลดฟอร์ม: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub btnExcelStats_Click(sender As Object, e As EventArgs)
        ShowExcelCacheStats()
    End Sub

    Private Sub txtSearch_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSearch.KeyDown
        Try
            ' ถ้ากด Enter ให้ค้นหา
            If e.KeyCode = Keys.Enter Then
                Dim searchText = txtSearch.Text.Trim()
                If Not String.IsNullOrEmpty(searchText) Then
                    PerformExcelSearch(searchText)
                End If
            End If
        Catch ex As Exception
            Console.WriteLine($"Error in txtSearch_KeyDown: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ฟังก์ชันสำหรับค้นหา Excel (ใช้จาก Event หรือปุ่ม)
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    Private Sub PerformExcelSearch(productCode As String)
        Try
            ' แสดงสถานะค้นหา
            ShowExcelLoadingStatus($"กำลังค้นหา '{productCode}'...")

            ' ค้นหาใน Cache
            Dim result = SearchInExcelCache(productCode)

            ' แสดงผลลัพธ์
            If result.IsSuccess Then
                ShowExcelLoadingStatus($"พบ '{productCode}': {result.MatchCount} รายการ")

                ' แสดงผลลัพธ์ใน MessageBox หรือ Form อื่น
                MessageBox.Show(result.SummaryMessage, "ผลการค้นหา",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                ShowExcelLoadingStatus($"ไม่พบ '{productCode}'")

                MessageBox.Show(result.SummaryMessage, "ผลการค้นหา",
                          MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            ShowExcelLoadingStatus($"ค้นหาผิดพลาด: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ฟังก์ชันสำหรับตรวจสอบสถานะ Excel Cache
    ''' </summary>
    ''' <returns>ข้อความสถานะ</returns>
    Private Function GetExcelCacheStatus() As String
        Try
            If dataCache Is Nothing Then
                Return "Excel Cache ยังไม่ได้เริ่มต้น"
            End If

            If dataCache.IsLoading Then
                Return "กำลังโหลดข้อมูล Excel..."
            End If

            If dataCache.IsLoaded Then
                Dim age = DateTime.Now - dataCache.LoadedTime
                Return $"ข้อมูล Excel: {dataCache.RowCount} แถว (อายุ {age.TotalMinutes:F0} นาที)"
            Else
                Return "ข้อมูล Excel ยังไม่ได้โหลด"
            End If

        Catch ex As Exception
            Return $"ตรวจสอบสถานะไม่ได้: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' อัพเดทสถานะ Excel Cache ใน Timer (ถ้าต้องการ)
    ''' </summary>
    Private Sub UpdateExcelCacheStatus()
        Try
            Dim status = GetExcelCacheStatus()

            ' อัพเดทสถานะใน UI
            If lblCount IsNot Nothing Then
                lblCount.Text = status
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in UpdateExcelCacheStatus: {ex.Message}")
        End Try
    End Sub

    ' ==============================
    ' Helper Methods สำหรับการจัดการ Excel Search
    ' ==============================

    ''' <summary>
    ''' ตรวจสอบว่า Excel Cache พร้อมใช้งานหรือไม่
    ''' </summary>
    ''' <returns>True ถ้าพร้อมใช้งาน</returns>
    Private Function IsExcelCacheReady() As Boolean
        Return dataCache IsNot Nothing AndAlso dataCache.IsLoaded
    End Function

    ''' <summary>
    ''' รอให้ Excel Cache โหลดเสร็จ (ใช้ใน case ที่ต้องการรอ)
    ''' </summary>
    ''' <param name="maxWaitSeconds">เวลารอสูงสุด (วินาที)</param>
    ''' <returns>True ถ้าโหลดเสร็จภายในเวลาที่กำหนด</returns>
    Private Function WaitForExcelCache(Optional maxWaitSeconds As Integer = 30) As Boolean
        Try
            Dim startTime = DateTime.Now

            While Not IsExcelCacheReady() AndAlso (DateTime.Now - startTime).TotalSeconds < maxWaitSeconds
                Application.DoEvents()
                Threading.Thread.Sleep(100)
            End While

            Return IsExcelCacheReady()

        Catch ex As Exception
            Console.WriteLine($"Error in WaitForExcelCache: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ดึงข้อมูล Excel สำหรับ Product Code โดยตรง
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>ข้อมูลแถวที่พบ หรือ Nothing ถ้าไม่พบ</returns>
    Private Function GetExcelDataForProduct(productCode As String) As ExcelRowData
        Try
            If Not IsExcelCacheReady() Then
                Return Nothing
            End If

            Dim result = dataCache.SearchInMemory(productCode)
            If result.IsSuccess AndAlso result.HasMatches Then
                ' แปลง ExcelMatchResult เป็น ExcelRowData
                Dim match = result.FirstMatch
                Return New ExcelRowData(match.RowNumber, match.ProductCode) With {
                .Column1Value = match.Column1Value,
                .Column2Value = match.Column2Value,
                .Column4Value = match.Column4Value,
                .Column5Value = match.Column5Value,
                .Column6Value = match.Column6Value
            }
            End If

            Return Nothing

        Catch ex As Exception
            Console.WriteLine($"Error in GetExcelDataForProduct: {ex.Message}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' ได้รายการ Product Code ทั้งหมดใน Cache (สำหรับ AutoComplete หรือ Dropdown)
    ''' </summary>
    ''' <returns>รายการ Product Code</returns>
    Private Function GetAllProductCodes() As List(Of String)
        Try
            If Not IsExcelCacheReady() Then
                Return New List(Of String)()
            End If

            Return dataCache.ExcelData.Where(Function(row) Not String.IsNullOrWhiteSpace(row.ProductCode)) _
                                 .Select(Function(row) row.ProductCode) _
                                 .Distinct() _
                                 .OrderBy(Function(code) code) _
                                 .ToList()

        Catch ex As Exception
            Console.WriteLine($"Error in GetAllProductCodes: {ex.Message}")
            Return New List(Of String)()
        End Try
    End Function

#End Region

    ' ==============================
    ' Timer สำหรับอัพเดทสถานะ (ถ้าต้องการ)
    ' ==============================

    ''' <summary>
    ''' Timer สำหรับอัพเดทสถานะ Excel Cache
    ''' </summary>
    Private WithEvents statusTimer As System.Windows.Forms.Timer

    ''' <summary>
    ''' เริ่มต้น Timer (เรียกใน Form Load ถ้าต้องการ)
    ''' </summary>
    Private Sub InitializeStatusTimer()
        Try
            statusTimer = New System.Windows.Forms.Timer()
            statusTimer.Interval = 5000 ' อัพเดททุก 5 วินาที
            statusTimer.Enabled = True
        Catch ex As Exception
            Console.WriteLine($"Error in InitializeStatusTimer: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' Event Handler สำหรับ Timer
    ''' </summary>
    Private Sub statusTimer_Tick(sender As Object, e As EventArgs) Handles statusTimer.Tick
        Try
            UpdateExcelCacheStatus()
        Catch ex As Exception
            Console.WriteLine($"Error in statusTimer_Tick: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ปุ่มสำหรับบังคับรีเฟรช Excel Cache (ถ้าต้องการเพิ่มปุ่ม)
    ''' </summary>
    Private Sub btnForceRefreshExcel_Click(sender As Object, e As EventArgs)
        Try
            If dataCache IsNot Nothing Then
                ShowExcelLoadingStatus("กำลังบังคับรีเฟรชข้อมูล Excel...")
                EnableExcelSearchControls(False)

                Task.Run(Sub() RefreshExcelDataAsync())
            Else
                MessageBox.Show("Excel Cache ยังไม่ได้เริ่มต้น", "ข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการรีเฟรช Excel: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    Private Sub frmHistory_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Try
            Console.WriteLine("กำลังปิด frmHistory...")

            ' เคลียร์ Excel Cache เมื่อปิด Form
            If dataCache IsNot Nothing Then
                Console.WriteLine("กำลังเคลียร์ Excel Cache...")
                dataCache.ClearData()
            End If

            ' ปิด Background Worker ถ้ามี
            If backgroundWorker IsNot Nothing AndAlso backgroundWorker.IsBusy Then
                backgroundWorker.CancelAsync()
            End If

            Console.WriteLine("ปิด frmHistory เรียบร้อยแล้ว")

        Catch ex As Exception
            Console.WriteLine($"Error in frmHistory_FormClosed: {ex.Message}")
        End Try
    End Sub

    Private Sub InitializeExcelCache()
        Try
            Console.WriteLine("กำลังเริ่มต้น Excel Cache...")

            ' เริ่มต้น Cache
            dataCache = ExcelDataCache.Instance

            ' แสดงสถานะพร้อมใช้งาน (ไม่โหลดข้อมูลทันที)
            ShowExcelLoadingStatus("พร้อมใช้งาน - ข้อมูล Excel จะโหลดเมื่อต้องการใช้งาน")
            EnableExcelSearchControls(True)

            Console.WriteLine("Excel Cache เริ่มต้นเรียบร้อย (Lazy Loading Mode)")

        Catch ex As Exception
            Console.WriteLine($"Error in InitializeExcelCache: {ex.Message}")
            ShowExcelLoadingStatus($"ไม่สามารถเริ่มต้น Excel Cache ได้: {ex.Message}")
            EnableExcelSearchControls(False)
        End Try
    End Sub

    ''' <summary>
    ''' ตรวจสอบและโหลดข้อมูล Excel เฉพาะเมื่อต้องการใช้งาน (Lazy Loading)
    ''' </summary>
    ''' <returns>True ถ้าข้อมูลพร้อมใช้งาน</returns>
    Private Async Function EnsureExcelDataLoadedAsync() As Task(Of Boolean)
        Try
            ' ตรวจสอบว่าข้อมูลโหลดแล้วหรือยัง
            If dataCache.IsLoaded AndAlso
               dataCache.ExcelFilePath.Equals(excelFilePath, StringComparison.OrdinalIgnoreCase) Then
                Console.WriteLine("ข้อมูล Excel โหลดแล้ว ไม่ต้องโหลดใหม่")
                Return True
            End If

            Console.WriteLine("เริ่มโหลดข้อมูล Excel แบบ Lazy Loading")

            ' แสดง Progress Bar
            ShowProgressBar()
            ShowExcelLoadingStatus("กำลังโหลดข้อมูล Excel...")

            ' ปิดการใช้งานปุ่มต่างๆ ระหว่างโหลด
            EnableExcelSearchControls(False)

            ' สร้าง Progress Handler
            Dim progressHandler As New Progress(Of Object)(
                Sub(progress)
                    Try
                        Me.Invoke(Sub()
                                      UpdateLoadingProgress(progress)
                                  End Sub)
                    Catch ex As Exception
                        Console.WriteLine($"Progress update error: {ex.Message}")
                    End Try
                End Sub)

            ' โหลดข้อมูลใน Background Thread
            Dim result = Await Task.Run(Function() As LoadResult
                                            Try
                                                Return dataCache.LoadExcelDataWithProgress(excelFilePath, progressHandler)
                                            Catch ex As Exception
                                                Console.WriteLine($"Error in background loading: {ex.Message}")
                                                Return New LoadResult() With {
                                                    .IsSuccess = False,
                                                    .ErrorMessage = ex.Message
                                                }
                                            End Try
                                        End Function)

            ' ซ่อน Progress Bar
            HideProgressBar()

            ' ตรวจสอบผลลัพธ์
            If result.IsSuccess Then
                ShowSuccessNotification($"โหลดข้อมูล Excel สำเร็จ! ({result.ProcessedRows:N0} แถว)")
                EnableExcelSearchControls(True)

                Console.WriteLine($"โหลดข้อมูล Excel สำเร็จ: {result.ProcessedRows:N0} แถว")
                Return True
            Else
                ShowExcelLoadingStatus($"❌ โหลดข้อมูลไม่สำเร็จ: {result.ErrorMessage}")
                EnableExcelSearchControls(False)
                Console.WriteLine($"โหลดข้อมูล Excel ไม่สำเร็จ: {result.ErrorMessage}")
                Return False
            End If

        Catch ex As Exception
            HideProgressBar()
            ShowExcelLoadingStatus($"❌ เกิดข้อผิดพลาด: {ex.Message}")
            EnableExcelSearchControls(False)
            Console.WriteLine($"Error in EnsureExcelDataLoadedAsync: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' โหลดข้อมูล Excel พร้อมแสดง Progress (สำหรับการรีเฟรช)
    ''' </summary>
    Private Sub LoadExcelDataAsyncWithProgress()
        Try
            Console.WriteLine("เริ่มโหลดข้อมูล Excel พร้อม Progress...")

            ' สร้าง Progress Handler
            Dim progressHandler As New Progress(Of Object)(
                Sub(progress)
                    Try
                        Me.Invoke(Sub()
                                      UpdateLoadingProgress(progress)
                                  End Sub)
                    Catch ex As Exception
                        Console.WriteLine($"Progress update error: {ex.Message}")
                    End Try
                End Sub)

            Dim result = dataCache.LoadExcelDataWithProgress(excelFilePath, progressHandler)

            ' อัพเดท UI ใน Main Thread
            Me.Invoke(Sub()
                          Try
                              HideProgressBar()

                              If result.IsSuccess Then
                                  ShowExcelLoadingStatus($"โหลดข้อมูล Excel สำเร็จ: {dataCache.RowCount:N0} แถว (ใช้เวลา {result.LoadTimeSeconds:F1} วินาที)")
                                  EnableExcelSearchControls(True)

                                  ' แสดงข้อมูลสถิติ
                                  Console.WriteLine(dataCache.GetMemoryStats())

                                  ' แสดงการแจ้งเตือนสั้นๆ
                                  ShowSuccessNotification($"โหลดข้อมูล {dataCache.RowCount:N0} แถว สำเร็จ")
                              Else
                                  ShowExcelLoadingStatus($"ไม่สามารถโหลด Excel ได้: {result.Message}")
                                  EnableExcelSearchControls(False)

                                  ' แสดง MessageBox เฉพาะข้อผิดพลาดสำคัญ
                                  If result.ErrorMessage.Contains("ไม่พบไฟล์") OrElse result.ErrorMessage.Contains("กำลังถูกใช้งาน") Then
                                      MessageBox.Show(result.Message, "ข้อมูล Excel ไม่พร้อม",
                                      MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                  End If
                              End If
                          Catch uiEx As Exception
                              Console.WriteLine($"Error updating UI after Excel load: {uiEx.Message}")
                          End Try
                      End Sub)

        Catch ex As Exception
            Console.WriteLine($"Error in LoadExcelDataAsyncWithProgress: {ex.Message}")
            Me.Invoke(Sub()
                          HideProgressBar()
                          ShowExcelLoadingStatus($"เกิดข้อผิดพลาด: {ex.Message}")
                          EnableExcelSearchControls(False)
                      End Sub)
        End Try
    End Sub

    ''' <summary>
    ''' อัพเดท Progress Bar และข้อความ
    ''' </summary>
    Private Sub UpdateLoadingProgress(progress As Object)
        Try
            ' ตรวจสอบและแปลง progress object
            Dim progressMessage As String = ""
            Dim processedRows As Integer = 0
            Dim totalRows As Integer = 0

            ' ใช้ Anonymous Type หรือ Dynamic Object
            Try
                Dim progressObj = CType(progress, Object)
                Dim progressType = progressObj.GetType()

                ' ดึงค่าจาก properties
                Dim messageProperty = progressType.GetProperty("Message")
                If messageProperty IsNot Nothing Then
                    Dim messageValue = messageProperty.GetValue(progressObj)
                    progressMessage = If(messageValue IsNot Nothing, messageValue.ToString(), "")
                End If

                Dim processedProperty = progressType.GetProperty("ProcessedRows")
                If processedProperty IsNot Nothing Then
                    Dim processedValue = processedProperty.GetValue(progressObj)
                    If processedValue IsNot Nothing Then
                        Integer.TryParse(processedValue.ToString(), processedRows)
                    End If
                End If

                Dim totalProperty = progressType.GetProperty("TotalRows")
                If totalProperty IsNot Nothing Then
                    Dim totalValue = totalProperty.GetValue(progressObj)
                    If totalValue IsNot Nothing Then
                        Integer.TryParse(totalValue.ToString(), totalRows)
                    End If
                End If

            Catch ex As Exception
                Console.WriteLine($"Error parsing progress object: {ex.Message}")
                progressMessage = "กำลังโหลดข้อมูล..."
            End Try

            ' อัพเดท Progress Bar
            If toolStripProgressBar IsNot Nothing Then
                If totalRows > 0 Then
                    toolStripProgressBar.Style = ProgressBarStyle.Continuous
                    toolStripProgressBar.Maximum = totalRows
                    toolStripProgressBar.Value = Math.Min(processedRows, totalRows)
                End If
            End If

            ' อัพเดทข้อความสถานะ
            Dim statusMessage = progressMessage
            If totalRows > 0 Then
                Dim percentage = (processedRows / totalRows * 100)
                statusMessage = $"{progressMessage} ({processedRows:N0}/{totalRows:N0} - {percentage:F1}%)"
            End If

            ShowExcelLoadingStatus(statusMessage)

            ' อัพเดท Application
            Application.DoEvents()

        Catch ex As Exception
            Console.WriteLine($"Error in UpdateLoadingProgress: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' แสดง Progress Bar
    ''' </summary>
    Private Sub ShowProgressBar()
        Try
            If toolStripProgressBar IsNot Nothing Then
                toolStripProgressBar.Visible = True
                toolStripProgressBar.Style = ProgressBarStyle.Marquee
                toolStripProgressBar.Value = 0
            End If
        Catch ex As Exception
            Console.WriteLine($"Error in ShowProgressBar: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ซ่อน Progress Bar
    ''' </summary>
    Private Sub HideProgressBar()
        Try
            If toolStripProgressBar IsNot Nothing Then
                toolStripProgressBar.Visible = False
                toolStripProgressBar.Value = 0
            End If
        Catch ex As Exception
            Console.WriteLine($"Error in HideProgressBar: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' แสดงการแจ้งเตือนสำเร็จแบบสั้น
    ''' </summary>
    Private Sub ShowSuccessNotification(message As String)
        Try
            ' สร้าง Timer เพื่อซ่อนข้อความหลังจาก 3 วินาที
            Dim hideTimer As New System.Windows.Forms.Timer()
            hideTimer.Interval = 3000

            AddHandler hideTimer.Tick, Sub(sender, e)
                                           hideTimer.Stop()
                                           hideTimer.Dispose()
                                           ShowExcelLoadingStatus("พร้อมใช้งาน")
                                       End Sub

            ShowExcelLoadingStatus($"✅ {message}")
            hideTimer.Start()

        Catch ex As Exception
            Console.WriteLine($"Error in ShowSuccessNotification: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' เปิด/ปิดการใช้งาน Controls ที่เกี่ยวข้องกับการค้นหา Excel
    ''' </summary>
    Private Sub EnableExcelSearchControls(enabled As Boolean)
        Try
            ' เปิด/ปิด txtSearch
            If txtSearch IsNot Nothing Then
                txtSearch.Enabled = enabled
                If enabled Then
                    txtSearch.BackColor = Color.White
                    txtSearch.PlaceholderText = "พิมพ์รหัสผลิตภัณฑ์และกด Enter..."
                Else
                    txtSearch.BackColor = Color.LightGray
                    txtSearch.PlaceholderText = "รอการโหลดข้อมูล Excel..."
                End If
            End If

            ' เปิด/ปิดปุ่มที่เกี่ยวข้อง (ถ้ามี)
            ' btnExcelSearch?.Enabled = enabled

        Catch ex As Exception
            Console.WriteLine($"Error in EnableExcelSearchControls: {ex.Message}")
        End Try
    End Sub

    Private Sub ShowExcelLoadingStatus(message As String)
        Try
            ' แสดงใน StatusStrip ถ้ามี
            If lblCount IsNot Nothing Then
                lblCount.Text = message
            End If

            ' แสดงใน Label สถานะถ้ามี
            If lblStatus IsNot Nothing Then
                lblStatus.Text = message
            End If

            ' แสดงใน Title Bar
            If Not String.IsNullOrEmpty(message) Then
                Me.Text = $"History - {message}"
            End If

            Console.WriteLine($"Excel Status: {message}")

            ' อัพเดท UI
            Application.DoEvents()

        Catch ex As Exception
            Console.WriteLine($"Error in ShowExcelLoadingStatus: {ex.Message}")
        End Try
    End Sub

    Private Sub btnClose_Click(sender As Object, e As EventArgs) Handles btnClose.Click
        Me.Close()
    End Sub

    Private Sub btnRefresh_Click(sender As Object, e As EventArgs) Handles btnRefresh.Click
        Try
            ' รีเฟรชข้อมูล History
            LoadScanHistory()

            ' รีเฟรชข้อมูล Excel Cache
            If dataCache IsNot Nothing AndAlso dataCache.IsLoaded Then
                ShowExcelLoadingStatus("กำลังรีเฟรชข้อมูล Excel...")
                EnableExcelSearchControls(False)

                Task.Run(Sub() RefreshExcelDataAsync())
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in btnRefresh_Click: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการรีเฟรช: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
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

    Private Sub btnSettings_Click(sender As Object, e As EventArgs) Handles btnSettings.Click
        ' เปิดหน้าจอตั้งค่าฐานข้อมูล
        Dim settingsForm As New frmSettings()
        If settingsForm.ShowDialog() = DialogResult.OK Then
            ' ตรวจสอบว่ามีการเปลี่ยนแปลงการตั้งค่าหรือไม่
            If settingsForm.HasUnsavedChanges Then
                ' เริ่มต้นการใช้งานฐานข้อมูลใหม่ตามการตั้งค่าที่เปลี่ยนแปลง
                AccessDatabaseManager.Initialize()

                ' อัปเดตชื่อหน้าต่างเพื่อแสดงพาธฐานข้อมูลใหม่
                Dim dbPath As String = AccessDatabaseManager.ConnectionString
                Me.Text = $"ประวัติการสแกน QR Code - {dbPath}"

                ' ตรวจสอบการเชื่อมต่อและโหลดข้อมูลใหม่
                CheckDatabaseConnection()
                LoadScanHistory()
            End If
        End If
    End Sub

    Private Sub dgvHistory_SelectionChanged(sender As Object, e As EventArgs) Handles dgvHistory.SelectionChanged
        UpdateButtonStates()
    End Sub

    Private Sub dgvHistory_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dgvHistory.CellContentClick
        ' จัดการการคลิกปุ่มในเซลล์
        If e.RowIndex >= 0 Then
            If e.ColumnIndex = dgvHistory.Columns("btnCheckExcel").Index Then
                ' คลิกปุ่มตรวจสอบ Excel
                CheckExcelFile(e.RowIndex)
            ElseIf e.ColumnIndex = dgvHistory.Columns("btnCreateMission").Index Then
                ' คลิกปุ่ม Mission
                HandleMissionButton(e.RowIndex)
            End If
        End If
    End Sub

    Private Sub txtSearch_TextChanged(sender As Object, e As EventArgs) Handles txtSearch.TextChanged
        ApplyFilters()
    End Sub

    Private Sub cmbStatus_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbStatus.SelectedIndexChanged
        ApplyFilters()
    End Sub

    Private Sub cmbMissionStatus_SelectedIndexChanged(sender As Object, e As EventArgs)
        ApplyFilters()
    End Sub

    Private Sub dtpFromDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpFromDate.ValueChanged
        ApplyFilters()
    End Sub

    Private Sub dtpToDate_ValueChanged(sender As Object, e As EventArgs) Handles dtpToDate.ValueChanged
        ApplyFilters()
    End Sub

    Public Shared Function Success(data As List(Of ExcelRowData), message As String) As LoadResult
        Dim result As New LoadResult(True, message) With {
            .Data = data,
            .ValidRows = If(data IsNot Nothing, data.Count, 0),
            .ProcessedRows = If(data IsNot Nothing, data.Count, 0)
        }
        result.StopTiming()
        Return result
    End Function


#Region "Initialization"
    Private Sub InitializeForm()
        Try
            Console.WriteLine("InitializeForm started")

            ' แสดงพาธฐานข้อมูลที่ใช้งาน
            Dim dbPath As String = AccessDatabaseManager.ConnectionString
            Me.Text = $"ประวัติการสแกน QR Code - {dbPath}"

            ' ตั้งค่าเริ่มต้นสำหรับ ComboBox สถานะความถูกต้อง
            cmbStatus.Items.Clear()
            cmbStatus.Items.AddRange(New String() {"ทั้งหมด", "ถูกต้อง", "ไม่ถูกต้อง"})
            cmbStatus.SelectedIndex = 0

            ' เพิ่ม ComboBox สำหรับกรองสถานะ Mission
            If Not pnlFilter.Controls.ContainsKey("cmbMissionStatus") Then
                Dim lblMissionStatus As New System.Windows.Forms.Label()
                lblMissionStatus.Name = "lblMissionStatus"
                lblMissionStatus.Text = "สถานะ Mission:"
                lblMissionStatus.AutoSize = True
                lblMissionStatus.Location = New System.Drawing.Point(655, 30)
                grpFilter.Controls.Add(lblMissionStatus)

                Dim cmbMissionStatus As New ComboBox()
                cmbMissionStatus.Name = "cmbMissionStatus"
                cmbMissionStatus.DropDownStyle = ComboBoxStyle.DropDownList
                cmbMissionStatus.Items.AddRange(New String() {"ทั้งหมด", "ไม่มี", "รอดำเนินการ", "สำเร็จ"})
                cmbMissionStatus.SelectedIndex = 0
                cmbMissionStatus.Location = New System.Drawing.Point(655, 48)
                cmbMissionStatus.Size = New Size(120, 23)
                AddHandler cmbMissionStatus.SelectedIndexChanged, AddressOf cmbMissionStatus_SelectedIndexChanged
                grpFilter.Controls.Add(cmbMissionStatus)
            End If


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

    Private Sub btnCreateAllMissions_Click(sender As Object, e As EventArgs) Handles btnCreateAllMissions.Click
        CreateAllMissions()
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
            btnCreateMission.HeaderText = "Mission"
            btnCreateMission.Text = "🚀 สร้าง"
            btnCreateMission.UseColumnTextForButtonValue = False
            btnCreateMission.Width = 100
            'สีปุ่ม
            btnCreateMission.DefaultCellStyle.ForeColor = Color.Blue
            dgvHistory.Columns.Add(btnCreateMission)

            ' สร้างคอลัมน์สถานะ Mission
            Dim colMissionStatus As New DataGridViewTextBoxColumn()
            colMissionStatus.Name = "MissionStatus"
            colMissionStatus.HeaderText = "สถานะ Mission"
            colMissionStatus.Width = 120
            dgvHistory.Columns.Add(colMissionStatus)

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

    ''' <summary>
    ''' ตรวจสอบการเชื่อมต่อฐานข้อมูล Access
    ''' </summary>
    Private Sub CheckDatabaseConnection()
        Try
            ' ดึงการตั้งค่าฐานข้อมูลจาก Settings
            Dim settings As New frmSettings()
            Dim dbPath As String = settings.GetAccessDatabasePath()

            ' เริ่มต้นการใช้งานฐานข้อมูล
            If Not AccessDatabaseManager.Initialize() Then
                MessageBox.Show($"ไม่สามารถเชื่อมต่อกับฐานข้อมูล: {dbPath}" & vbNewLine &
                              "กรุณาตรวจสอบการตั้งค่าฐานข้อมูลและสิทธิ์การเข้าถึง",
                              "ข้อผิดพลาดการเชื่อมต่อ", MessageBoxButtons.OK, MessageBoxIcon.Warning)

                ' แสดงสถานะในแถบสถานะ
                lblCount.Text = "ไม่สามารถเชื่อมต่อกับฐานข้อมูล"
                lblCount.ForeColor = Color.Red
            Else
                ' เชื่อมต่อสำเร็จ
                lblCount.Text = "เชื่อมต่อฐานข้อมูลสำเร็จ"
                lblCount.ForeColor = Color.Green
            End If
        Catch ex As Exception
            Console.WriteLine($"Error checking database connection: {ex.Message}")
            lblCount.Text = $"ข้อผิดพลาดการเชื่อมต่อ: {ex.Message}"
            lblCount.ForeColor = Color.Red
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

            ' ดึงค่า MaxRecords จาก Settings
            Dim settings As New frmSettings()
            Dim maxRecords As Integer = CInt(settings.GetSetting("maxrecords"))

            If backgroundWorker IsNot Nothing AndAlso Not backgroundWorker.IsBusy Then
                backgroundWorker.RunWorkerAsync(maxRecords)
            Else
                ' ถ้า background worker ไม่พร้อม ให้โหลดแบบ synchronous
                LoadDataSynchronous(maxRecords)
            End If

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
            Dim maxRecords As Integer = CInt(e.Argument)
            Dim data As List(Of ScanDataRecord) = AccessDatabaseManager.GetScanHistory(maxRecords)

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

    Private Sub LoadDataSynchronous(maxRecords As Integer)
        Try
            Console.WriteLine("Loading data synchronously")
            scanHistory = AccessDatabaseManager.GetScanHistory(maxRecords)
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
            If scanHistory Is Nothing Then Return

            ' ดึงค่าการกรอง
            Dim searchText As String = txtSearch.Text.Trim().ToLower()
            Dim statusFilter As String = cmbStatus.SelectedItem.ToString()

            ' ดึงค่าการกรองสถานะ Mission
            Dim missionStatusFilter As String = "ทั้งหมด"
            Dim cmbMissionStatus As ComboBox = TryCast(grpFilter.Controls("cmbMissionStatus"), ComboBox)

            If cmbMissionStatus IsNot Nothing AndAlso cmbMissionStatus.SelectedItem IsNot Nothing Then
                missionStatusFilter = cmbMissionStatus.SelectedItem.ToString()
            End If

            Dim fromDate As DateTime = dtpFromDate.Value.Date
            Dim toDate As DateTime = dtpToDate.Value.Date.AddDays(1).AddSeconds(-1) ' ถึงสิ้นวัน

            ' กรองข้อมูล
            filteredHistory = scanHistory.Where(Function(record)
                                                    ' กรองตามวันที่
                                                    Dim isInDateRange As Boolean = record.ScanDateTime >= fromDate AndAlso record.ScanDateTime <= toDate

                                                    ' กรองตามสถานะความถูกต้อง
                                                    Dim matchesStatus As Boolean = statusFilter = "ทั้งหมด" OrElse
                                          (statusFilter = "ถูกต้อง" AndAlso record.IsValid) OrElse
                                          (statusFilter = "ไม่ถูกต้อง" AndAlso Not record.IsValid)

                                                    ' กรองตามสถานะ Mission - แก้ไขส่วนนี้
                                                    Dim matchesMissionStatus As Boolean
                                                    If missionStatusFilter = "ทั้งหมด" Then
                                                        ' เมื่อเลือก "ทั้งหมด" ให้แสดงเฉพาะ "รอดำเนินการ" และ "ไม่มี"
                                                        matchesMissionStatus = record.MissionStatus = "รอดำเนินการ" OrElse
                                                                         record.MissionStatus = "ไม่มี" OrElse
                                                                         String.IsNullOrEmpty(record.MissionStatus)
                                                    Else
                                                        ' เมื่อเลือกสถานะเฉพาะ ให้แสดงตามที่เลือก
                                                        matchesMissionStatus = record.MissionStatus = missionStatusFilter
                                                    End If

                                                    ' กรองตามข้อความค้นหา
                                                    Dim matchesSearch As Boolean = String.IsNullOrEmpty(searchText) OrElse
                                         record.ProductCode.ToLower().Contains(searchText) OrElse
                                         record.ReferenceCode.ToLower().Contains(searchText) OrElse
                                         record.DateCode.ToLower().Contains(searchText)

                                                    ' ต้องตรงกับทุกเงื่อนไข
                                                    Return isInDateRange AndAlso matchesStatus AndAlso matchesMissionStatus AndAlso matchesSearch
                                                End Function).ToList()

            ' แสดงผลข้อมูลที่กรอง
            DisplayData()

        Catch ex As Exception
            Console.WriteLine($"Error in ApplyFilters: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการกรองข้อมูล: {ex.Message}",
                      "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

                        ' แสดงสถานะ Mission และกำหนดปุ่มตามสถานะ
                        .Cells("MissionStatus").Value = record.MissionStatus

                        ' กำหนดสีของข้อความสถานะ Mission
                        Select Case record.MissionStatus
                            Case "รอดำเนินการ"
                                .Cells("MissionStatus").Style.ForeColor = Color.Orange
                            Case "สำเร็จ"
                                .Cells("MissionStatus").Style.ForeColor = Color.Green
                        End Select

                        ' กำหนดปุ่ม Mission ตามสถานะ
                        Select Case record.MissionStatus
                            Case "ไม่มี"
                                If record.IsValid Then
                                    ' ตรวจสอบว่าข้อมูลมีใน Excel และหาไฟล์เจอแบบ 1:1 หรือไม่
                                    Dim canCreateMission As Boolean = CheckCanCreateMission(record.ProductCode)
                                    If canCreateMission Then
                                        .Cells("btnCreateMission").Value = "🚀 สร้าง"
                                        .Cells("btnCreateMission").Style.ForeColor = Color.Blue
                                    ElseIf Not dataCache.IsLoaded Then
                                        .Cells("btnCreateMission").Value = "📊 โหลดข้อมูล"
                                        .Cells("btnCreateMission").Style.ForeColor = Color.DarkBlue
                                    Else
                                        .Cells("btnCreateMission").Value = "⚠️ ไม่พร้อม"
                                        .Cells("btnCreateMission").Style.ForeColor = Color.Orange
                                    End If
                                Else
                                    .Cells("btnCreateMission").Value = "⛔ ไม่สามารถสร้าง"
                                    .Cells("btnCreateMission").Style.ForeColor = Color.Gray
                                End If
                            Case "รอดำเนินการ"
                                .Cells("btnCreateMission").Value = "📋 ตรวจสอบ"
                                .Cells("btnCreateMission").Style.ForeColor = Color.Orange
                            Case "สำเร็จ"
                                .Cells("btnCreateMission").Value = "✅ สำเร็จ"
                                .Cells("btnCreateMission").Style.ForeColor = Color.Green
                        End Select

                        ' เก็บข้อมูลในแท็กของแถว
                        .Tag = record
                    End With
                Next

                ' อัปเดตจำนวนรายการที่แสดง
                lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count}"
            Else
                lblCount.Text = "ไม่พบข้อมูล"
            End If

            ' อัปเดตสถานะปุ่ม
            UpdateButtonStates()

        Catch ex As Exception
            Console.WriteLine($"Error in DisplayData: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
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

    ''' <summary>
    ''' ตรวจสอบว่าสามารถสร้าง Mission ได้หรือไม่ (แบบ Lazy Loading)
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>True ถ้าสามารถสร้าง Mission ได้</returns>
    Private Async Function CheckCanCreateMissionAsync(productCode As String) As Task(Of Boolean)
        Try
            If String.IsNullOrEmpty(productCode) Then
                Return False
            End If

            ' ตรวจสอบการเชื่อมต่อเครือข่าย
            Dim networkResult As NetworkCheckResult = CheckNetworkConnection()
            If Not networkResult.IsConnected OrElse networkResult.NetworkType <> "OA" Then
                Return False
            End If

            ' ตรวจสอบว่าไฟล์ Excel มีอยู่หรือไม่
            If Not File.Exists(excelFilePath) Then
                Return False
            End If

            ' ตรวจสอบและโหลดข้อมูล Excel ถ้าจำเป็น
            Dim dataLoaded = Await EnsureExcelDataLoadedAsync()
            If Not dataLoaded Then
                Return False
            End If

            ' ค้นหาข้อมูลใน Cache
            Dim searchResult = dataCache.SearchInMemory(productCode)
            If Not searchResult.IsSuccess OrElse Not searchResult.HasMatches Then
                Return False
            End If

            ' ตรวจสอบว่าหาไฟล์เจอแบบ 1:1 หรือไม่
            If searchResult.FirstMatch IsNot Nothing AndAlso Not String.IsNullOrEmpty(searchResult.FirstMatch.Column4Value) Then
                Dim fileSearchResult = SearchFilesInDirectory(searchResult.FirstMatch.Column4Value)
                ' ต้องเจอไฟล์พอดี 1 ไฟล์
                If fileSearchResult.FilesFound.Count = 1 Then
                    Return True
                End If
            End If

            Return False

        Catch ex As Exception
            Console.WriteLine($"Error in CheckCanCreateMissionAsync: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ตรวจสอบว่าสามารถสร้าง Mission ได้หรือไม่ (เวอร์ชัน Synchronous สำหรับ UI)
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>True ถ้าสามารถสร้าง Mission ได้</returns>
    Private Function CheckCanCreateMission(productCode As String) As Boolean
        Try
            If String.IsNullOrEmpty(productCode) Then
                Return False
            End If

            ' ตรวจสอบการเชื่อมต่อเครือข่าย
            Dim networkResult As NetworkCheckResult = CheckNetworkConnection()
            If Not networkResult.IsConnected OrElse networkResult.NetworkType <> "OA" Then
                Return False
            End If

            ' ตรวจสอบว่าไฟล์ Excel มีอยู่หรือไม่
            If Not File.Exists(excelFilePath) Then
                Return False
            End If

            ' ถ้าข้อมูลโหลดแล้ว ใช้ Cache
            If dataCache.IsLoaded Then
                Dim searchResult = dataCache.SearchInMemory(productCode)
                If Not searchResult.IsSuccess OrElse Not searchResult.HasMatches Then
                    Return False
                End If

                ' ตรวจสอบว่าหาไฟล์เจอแบบ 1:1 หรือไม่
                If searchResult.FirstMatch IsNot Nothing AndAlso Not String.IsNullOrEmpty(searchResult.FirstMatch.Column4Value) Then
                    Dim fileSearchResult = SearchFilesInDirectory(searchResult.FirstMatch.Column4Value)
                    If fileSearchResult.FilesFound.Count = 1 Then
                        Return True
                    End If
                End If
            Else
                ' ถ้าข้อมูลยังไม่โหลด ให้ส่งคืน False และจะโหลดเมื่อผู้ใช้กดปุ่ม
                Return False
            End If

            Return False

        Catch ex As Exception
            Console.WriteLine($"Error in CheckCanCreateMission: {ex.Message}")
            Return False
        End Try
    End Function
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

    ''' <summary>
    ''' จัดการการคลิกปุ่ม Mission
    ''' </summary>
    Private Async Sub HandleMissionButton(rowIndex As Integer)
        Try
            If rowIndex < 0 OrElse rowIndex >= dgvHistory.Rows.Count Then
                Return
            End If

            Dim record As ScanDataRecord = CType(dgvHistory.Rows(rowIndex).Tag, ScanDataRecord)
            If record Is Nothing Then Return

            Select Case record.MissionStatus
                Case "ไม่มี"
                    ' ตรวจสอบความถูกต้องของข้อมูล
                    If Not record.IsValid Then
                        MessageBox.Show("ไม่สามารถสร้าง Mission ได้เนื่องจากข้อมูลไม่ถูกต้อง",
                                       "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If

                    ' ตรวจสอบว่าสามารถสร้าง Mission ได้หรือไม่
                    Dim canCreateResult As MissionCreationCheck = Await CheckMissionCreationRequirementsAsync(record.ProductCode)
                    If Not canCreateResult.CanCreate Then
                        MessageBox.Show($"ไม่สามารถสร้าง Mission ได้{vbCrLf}{vbCrLf}เหตุผล: {canCreateResult.Reason}",
                                       "ไม่สามารถสร้าง Mission", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                        Return
                    End If

                    ' สร้าง Mission ใหม่ (ส่งข้อมูลเพิ่มเติมด้วย)
                    If CreateNewMission(record, canCreateResult) Then
                        ' อัปเดตสถานะเป็น "รอดำเนินการ"
                        record.MissionStatus = "รอดำเนินการ"
                        UpdateMissionStatus(record)

                        ' อัปเดตการแสดงผลในตาราง
                        dgvHistory.Rows(rowIndex).Cells("MissionStatus").Value = "รอดำเนินการ"
                        dgvHistory.Rows(rowIndex).Cells("btnCreateMission").Value = "📋 ตรวจสอบ"
                        dgvHistory.Rows(rowIndex).Cells("btnCreateMission").Style.ForeColor = Color.Orange
                    End If

                Case "รอดำเนินการ"
                    CheckMissionStatus(record, rowIndex)

                Case "สำเร็จ"
                    ShowCompletedMissionDetails(record)
            End Select

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการจัดการ Mission: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in HandleMissionButton: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ตรวจสอบข้อกำหนดในการสร้าง Mission อย่างละเอียด (แบบ Lazy Loading)
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>ผลการตรวจสอบพร้อมเหตุผล</returns>
    Private Async Function CheckMissionCreationRequirementsAsync(productCode As String) As Task(Of MissionCreationCheck)
        Dim result As New MissionCreationCheck()

        Try
            ' ตรวจสอบรหัสผลิตภัณฑ์
            If String.IsNullOrEmpty(productCode) Then
                result.CanCreate = False
                result.Reason = "ไม่พบรหัสผลิตภัณฑ์"
                Return result
            End If

            ' ตรวจสอบการเชื่อมต่อเครือข่าย
            Dim networkResult As NetworkCheckResult = CheckNetworkConnection()
            If Not networkResult.IsConnected Then
                result.CanCreate = False
                result.Reason = $"ไม่สามารถเชื่อมต่อเครือข่ายได้{vbCrLf}{networkResult.ErrorMessage}"
                Return result
            End If

            If networkResult.NetworkType <> "OA" Then
                result.CanCreate = False
                result.Reason = "ต้องเชื่อมต่อกับเครือข่าย OA เท่านั้น"
                Return result
            End If

            ' ตรวจสอบไฟล์ Excel
            If Not File.Exists(excelFilePath) Then
                result.CanCreate = False
                result.Reason = "ไม่พบไฟล์ Excel Database"
                Return result
            End If

            ' ตรวจสอบและโหลดข้อมูล Excel ถ้าจำเป็น
            Dim dataLoaded = Await EnsureExcelDataLoadedAsync()
            If Not dataLoaded Then
                result.CanCreate = False
                result.Reason = "ไม่สามารถโหลดข้อมูล Excel ได้"
                Return result
            End If

            ' ค้นหาข้อมูลใน Cache
            Dim searchResult = dataCache.SearchInMemory(productCode)
            If Not searchResult.IsSuccess Then
                result.CanCreate = False
                result.Reason = $"เกิดข้อผิดพลาดในการค้นหา Excel{vbCrLf}{searchResult.ErrorMessage}"
                Return result
            End If

            If Not searchResult.HasMatches Then
                result.CanCreate = False
                result.Reason = $"ไม่พบรหัสผลิตภัณฑ์ '{productCode}' ในไฟล์ Excel"
                Return result
            End If

            ' ตรวจสอบข้อมูลใน Excel
            If searchResult.FirstMatch Is Nothing OrElse String.IsNullOrEmpty(searchResult.FirstMatch.Column4Value) Then
                result.CanCreate = False
                result.Reason = "ข้อมูลใน Excel ไม่สมบูรณ์"
                Return result
            End If

            ' ค้นหาไฟล์ตามข้อมูลใน Excel
            Dim fileSearchResult = SearchFilesInDirectory(searchResult.FirstMatch.Column4Value)
            If fileSearchResult.FilesFound.Count = 0 Then
                result.CanCreate = False
                result.Reason = $"ไม่พบไฟล์ที่เกี่ยวข้องกับ '{searchResult.FirstMatch.Column4Value}'"
                Return result
            End If

            If fileSearchResult.FilesFound.Count > 1 Then
                result.CanCreate = False
                result.Reason = $"พบไฟล์ที่เกี่ยวข้อง {fileSearchResult.FilesFound.Count} ไฟล์{vbCrLf}ต้องมีไฟล์เดียวเท่านั้น"
                Return result
            End If

            ' ผ่านการตรวจสอบทั้งหมด
            result.CanCreate = True
            result.ExcelMatch = searchResult.FirstMatch
            result.FoundFile = fileSearchResult.FilesFound(0)
            result.Reason = "พร้อมสร้าง Mission"

            Return result

        Catch ex As Exception
            result.CanCreate = False
            result.Reason = $"เกิดข้อผิดพลาด: {ex.Message}"
            Console.WriteLine($"Error in CheckMissionCreationRequirements: {ex.Message}")
            Return result
        End Try
    End Function

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
                Dim excelPath As String = "\\10.24.179.2\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

                If File.Exists(excelPath) Then
                    ' ค้นหาข้อมูลในไฟล์ Excel โดยใช้ฟังก์ชัน SearchProductInExcel จาก ExcelUtility
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

            ' ค้นหาข้อมูลในไฟล์ Excel โดยใช้ฟังก์ชัน SearchProductInExcel จาก ExcelUtility
            Dim searchResult As ExcelUtility.ExcelSearchResult = ExcelUtility.SearchProductInExcel(excelPath, productCode)

            searchForm.Close()

            ' แสดงผลลัพธ์
            DisplayExcelSearchResult(searchResult)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการค้นหาข้อมูล: {ex.Message}",
                          "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub



    Private Function SearchInExcelCache(productCode As String) As ExcelUtility.ExcelSearchResult
        Try
            If dataCache Is Nothing OrElse Not dataCache.IsLoaded Then
                Return New ExcelUtility.ExcelSearchResult() With {
                .SearchedProductCode = productCode,
                .ExcelFilePath = excelFilePath,
                .IsSuccess = False,
                .ErrorMessage = "ข้อมูล Excel ยังไม่ได้โหลด",
                .SummaryMessage = "❌ ข้อมูล Excel ยังไม่พร้อม กรุณารอการโหลดให้เสร็จสิ้น"
            }
            End If

            If String.IsNullOrWhiteSpace(productCode) Then
                Return New ExcelUtility.ExcelSearchResult() With {
                .SearchedProductCode = "",
                .ExcelFilePath = excelFilePath,
                .IsSuccess = False,
                .ErrorMessage = "ไม่มีรหัสผลิตภัณฑ์ที่ต้องการค้นหา",
                .SummaryMessage = "❌ กรุณาใส่รหัสผลิตภัณฑ์ที่ต้องการค้นหา"
            }
            End If

            ' ค้นหาใน Cache (เร็วมาก!)
            Dim startTime = DateTime.Now
            Dim result = dataCache.SearchInMemory(productCode)
            Dim elapsedTime = DateTime.Now - startTime

            Console.WriteLine($"ค้นหาเสร็จสิ้นใน {elapsedTime.TotalMilliseconds:F2} มิลลิวินาที")
            Return result

        Catch ex As Exception
            Console.WriteLine($"Error in SearchInExcelCache: {ex.Message}")
            Return New ExcelUtility.ExcelSearchResult() With {
            .SearchedProductCode = productCode,
            .ExcelFilePath = excelFilePath,
            .IsSuccess = False,
            .ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}",
            .SummaryMessage = "❌ เกิดข้อผิดพลาดในการค้นหา"
        }
        End Try
    End Function

    ''' <summary>
    ''' แสดงข้อมูลสถิติ Excel Cache
    ''' </summary>
    Private Sub ShowExcelCacheStats()
        Try
            If dataCache IsNot Nothing Then
                Dim stats = dataCache.GetMemoryStats()
                MessageBox.Show(stats, "สถิติ Excel Cache", MessageBoxButtons.OK, MessageBoxIcon.Information)
            Else
                MessageBox.Show("Excel Cache ยังไม่ได้เริ่มต้น", "ข้อมูล", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If
        Catch ex As Exception
            MessageBox.Show($"ไม่สามารถแสดงสถิติได้: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
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
                            message.AppendLine($"⚠️ มีโฟลเดอร์ที่เข้าถึงไม่ได้ {fileSearchResult.ErrorDirectories.Count} โฟลเดอร์")
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
    ''' ค้นหาข้อมูลในไฟล์ Excel โดยใช้ ExcelUtility
    ''' </summary>
    Private Function SearchProductInExcel(productCode As String) As ExcelUtility.ExcelSearchResult
        ' กำหนดเส้นทางไฟล์ Excel
        Dim excelFilePath As String = "\\10.24.179.2\OAFAB\OA2FAB\Film charecter check\Database.xlsx"

        ' เรียกใช้งานฟังก์ชัน SearchProductInExcel จากคลาส ExcelUtility
        Try
            Return ExcelUtility.SearchProductInExcel(excelFilePath, productCode)
        Catch ex As Exception
            Console.WriteLine($"Error in SearchProductInExcel: {ex.Message}")

            ' สร้าง result ที่แสดงข้อผิดพลาด
            Dim errorResult As New ExcelUtility.ExcelSearchResult()
            errorResult.SearchedProductCode = productCode
            errorResult.ExcelFilePath = excelFilePath
            errorResult.IsSuccess = False
            errorResult.ErrorMessage = $"เกิดข้อผิดพลาดในการค้นหา: {ex.Message}"
            errorResult.SummaryMessage = $"❌ ไม่สามารถค้นหาข้อมูลได้: {ex.Message}"

            Return errorResult
        End Try
    End Function

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
            Dim baseFolderPath As String = "\\10.24.179.2\OAFAB\OA2FAB\Film charecter check"

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
                            .RelativePath = GetRelativePath("\\10.24.179.2\OAFAB\OA2FAB\20250607 Pimploy S", filePath),
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
                           $"วันที่สแกน: {record.ScanDateTime:dd/MM/yyyy HH:mm:ss}{Environment.NewLine}" &
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
                                                      $"วันที่สแกน: {record.ScanDateTime}",
                                                      "ยืนยันการลบ", System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Warning)

            If result = System.Windows.Forms.DialogResult.Yes Then
                ' ลบข้อมูลจากฐานข้อมูล
                Dim success As Boolean = AccessDatabaseManager.DeleteScanRecord(record.Id)

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
                Dim excelApp As New Excel.Application()
                excelApp.Visible = False

                ' สร้าง Workbook
                Dim workbook As Excel.Workbook = excelApp.Workbooks.Add()
                Dim worksheet As Excel.Worksheet = CType(workbook.Worksheets(1), Excel.Worksheet)

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
                Dim headerRange As Excel.Range = worksheet.Range("A1:J1")
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

    ''' <summary>
    ''' คืนทรัพยากร COM object
    ''' </summary>
    ''' <param name="obj">COM object ที่ต้องการคืนทรัพยากร</param>
    Private Sub ReleaseObject(obj As Object)
        Try
            If obj IsNot Nothing Then
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
                obj = Nothing
            End If
        Catch ex As Exception
            obj = Nothing
            Console.WriteLine($"Error releasing COM object: {ex.Message}")
        Finally
            GC.Collect()
            GC.WaitForPendingFinalizers()
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
        End Try
    End Sub

#Region "Mission Management Functions"

    Private Function CreateNewMission(record As ScanDataRecord, creationCheck As MissionCreationCheck) As Boolean
        Try
            ' ตรวจสอบความถูกต้องของข้อมูล
            If record Is Nothing OrElse String.IsNullOrEmpty(record.ProductCode) OrElse Not record.IsValid Then
                MessageBox.Show("ข้อมูลไม่ถูกต้องหรือไม่สามารถสร้าง Mission ได้",
                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            If Not creationCheck.CanCreate Then
                MessageBox.Show($"ไม่สามารถสร้าง Mission ได้{vbCrLf}{creationCheck.Reason}",
                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

            ' แสดงกล่องโต้ตอบสำหรับการสร้าง Mission ที่ปรับปรุงแล้ว
            Dim missionForm As New Form()
            missionForm.Text = "สร้าง Mission ใหม่"
            missionForm.Size = New Size(850, 750)  ' เพิ่มขนาดเล็กน้อย
            missionForm.StartPosition = FormStartPosition.CenterParent
            missionForm.FormBorderStyle = FormBorderStyle.FixedDialog
            missionForm.MaximizeBox = False
            missionForm.MinimizeBox = False
            missionForm.BackColor = Color.WhiteSmoke
            missionForm.Font = New Font("Segoe UI", 9)
            missionForm.Padding = New Padding(15)  ' เพิ่ม padding รอบฟอร์ม

            ' ตัวแปรสำหรับจัดการ layout แบบ responsive
            Dim formWidth As Integer = missionForm.ClientSize.Width - 30  ' หัก padding ซ้าย-ขวา
            Dim panelMargin As Integer = 10
            Dim currentY As Integer = panelMargin

            ' Header Panel
            Dim headerPanel As New Panel()
            headerPanel.Size = New Size(formWidth, 90)  ' เพิ่มความสูง
            headerPanel.Location = New Point(15, currentY)
            headerPanel.BackColor = Color.FromArgb(52, 152, 219)
            headerPanel.BorderStyle = BorderStyle.None
            headerPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            Dim lblTitle As New Label()
            lblTitle.Text = "🚀 สร้าง Mission สำหรับข้อมูลที่สแกน"
            lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
            lblTitle.Location = New Point(20, 15)
            lblTitle.AutoSize = True
            lblTitle.ForeColor = Color.White
            lblTitle.BackColor = Color.Transparent

            Dim lblSubtitle As New Label()
            lblSubtitle.Text = "กรุณาตรวจสอบข้อมูลและกรอกรายละเอียดเพิ่มเติม"
            lblSubtitle.Font = New Font("Segoe UI", 11)
            lblSubtitle.Location = New Point(20, 50)
            lblSubtitle.Size = New Size(formWidth - 40, 25)  ' กำหนดขนาดเพื่อ wrap text
            lblSubtitle.ForeColor = Color.White
            lblSubtitle.BackColor = Color.Transparent

            headerPanel.Controls.AddRange({lblTitle, lblSubtitle})
            currentY += headerPanel.Height + panelMargin

            ' Info Panel
            Dim infoPanel As New Panel()
            infoPanel.Size = New Size(formWidth, 120)
            infoPanel.Location = New Point(15, currentY)
            infoPanel.BackColor = Color.White
            infoPanel.BorderStyle = BorderStyle.FixedSingle
            infoPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' สร้าง TableLayoutPanel สำหรับ Info Panel เพื่อจัดเรียงอัตโนมัติ
            Dim infoTable As New TableLayoutPanel()
            infoTable.Size = New Size(formWidth - 20, 100)
            infoTable.Location = New Point(10, 10)
            infoTable.ColumnCount = 2
            infoTable.RowCount = 3
            infoTable.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' กำหนดขนาดคอลัมน์
            infoTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 75))
            infoTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 25))

            ' กำหนดขนาดแถว
            infoTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            infoTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))
            infoTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 30))

            Dim lblProductCode As New Label()
            lblProductCode.Text = $"📦 รหัสผลิตภัณฑ์: {record.ProductCode}"
            lblProductCode.Font = New Font("Segoe UI", 11, FontStyle.Bold)
            lblProductCode.ForeColor = Color.FromArgb(52, 73, 94)
            lblProductCode.TextAlign = ContentAlignment.MiddleLeft
            lblProductCode.Dock = DockStyle.Fill

            Dim lblExcelInfo As New Label()
            lblExcelInfo.Text = $"📊 ข้อมูลจาก Excel: {creationCheck.ExcelMatch.Column4Value}"
            lblExcelInfo.ForeColor = Color.FromArgb(39, 174, 96)
            lblExcelInfo.Font = New Font("Segoe UI", 10)
            lblExcelInfo.TextAlign = ContentAlignment.MiddleLeft
            lblExcelInfo.Dock = DockStyle.Fill

            Dim lblFileInfo As New Label()
            lblFileInfo.Text = $"📁 ไฟล์ที่เกี่ยวข้อง: {creationCheck.FoundFile.FileName}"
            lblFileInfo.ForeColor = Color.FromArgb(41, 128, 185)
            lblFileInfo.Font = New Font("Segoe UI", 10)
            lblFileInfo.TextAlign = ContentAlignment.MiddleLeft
            lblFileInfo.Dock = DockStyle.Fill

            ' ปุ่มดูไฟล์
            Dim btnPreviewFile As New Button()
            btnPreviewFile.Text = "👁️ ดูไฟล์"
            btnPreviewFile.Size = New Size(100, 30)
            btnPreviewFile.BackColor = Color.FromArgb(52, 152, 219)
            btnPreviewFile.ForeColor = Color.White
            btnPreviewFile.FlatStyle = FlatStyle.Flat
            btnPreviewFile.FlatAppearance.BorderSize = 0
            btnPreviewFile.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnPreviewFile.Anchor = AnchorStyles.Top Or AnchorStyles.Right
            AddHandler btnPreviewFile.Click, Sub()
                                                 OpenFileWithErrorHandling(creationCheck.FoundFile.FullPath)
                                             End Sub

            ' เพิ่ม controls ลง TableLayoutPanel
            infoTable.Controls.Add(lblProductCode, 0, 0)
            infoTable.Controls.Add(lblExcelInfo, 0, 1)
            infoTable.Controls.Add(lblFileInfo, 0, 2)
            infoTable.Controls.Add(btnPreviewFile, 1, 2)

            infoPanel.Controls.Add(infoTable)
            currentY += infoPanel.Height + panelMargin

            ' Form Panel - ใช้ TableLayoutPanel เพื่อจัดการ layout
            Dim formPanel As New Panel()
            formPanel.Size = New Size(formWidth, 380)
            formPanel.Location = New Point(15, currentY)
            formPanel.BackColor = Color.White
            formPanel.BorderStyle = BorderStyle.FixedSingle
            formPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' Mission Name Section
            Dim lblMissionName As New Label()
            lblMissionName.Text = "📝 ชื่อ Mission:"
            lblMissionName.Location = New Point(20, 20)
            lblMissionName.AutoSize = True
            lblMissionName.Font = New Font("Segoe UI", 11, FontStyle.Bold)
            lblMissionName.ForeColor = Color.FromArgb(52, 73, 94)

            Dim txtMissionName As New TextBox()
            txtMissionName.Text = $"ตรวจสอบ {record.ProductCode} - {creationCheck.ExcelMatch.Column4Value}"
            txtMissionName.Location = New Point(20, 45)
            txtMissionName.Size = New Size(formWidth - 60, 25)
            txtMissionName.Font = New Font("Segoe UI", 9)
            txtMissionName.BorderStyle = BorderStyle.FixedSingle
            txtMissionName.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' Description Section
            Dim lblDescription As New Label()
            lblDescription.Text = "📋 รายละเอียด Mission:"
            lblDescription.Location = New Point(20, 85)
            lblDescription.AutoSize = True
            lblDescription.Font = New Font("Segoe UI", 11, FontStyle.Bold)
            lblDescription.ForeColor = Color.FromArgb(52, 73, 94)

            Dim txtDescription As New TextBox()
            txtDescription.Multiline = True
            txtDescription.Text = $"ตรวจสอบและดำเนินการกับข้อมูล QR Code{vbCrLf}" &
                         $"• รหัสผลิตภัณฑ์: {record.ProductCode}{vbCrLf}" &
                         $"• รหัสอ้างอิง: {record.ReferenceCode}{vbCrLf}" &
                         $"• จำนวน: {record.Quantity}{vbCrLf}" &
                         $"• วันที่ผลิต: {record.DateCode}{vbCrLf}" &
                         $"• ข้อมูล Excel: {creationCheck.ExcelMatch.Column4Value}{vbCrLf}" &
                         $"• ไฟล์เกี่ยวข้อง: {creationCheck.FoundFile.FileName}"
            txtDescription.Location = New Point(20, 110)
            txtDescription.Size = New Size(formWidth - 60, 180)
            txtDescription.ScrollBars = ScrollBars.Vertical
            txtDescription.WordWrap = True
            txtDescription.Font = New Font("Segoe UI", 9)
            txtDescription.BorderStyle = BorderStyle.FixedSingle
            txtDescription.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' Bottom Section - ใช้ TableLayoutPanel สำหรับ Assigned To และ Due Date
            Dim bottomTable As New TableLayoutPanel()
            bottomTable.Location = New Point(20, 310)
            bottomTable.Size = New Size(formWidth - 60, 50)
            bottomTable.ColumnCount = 4
            bottomTable.RowCount = 2
            bottomTable.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' กำหนดขนาดคอลัมน์
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))  ' Label ผู้รับผิดชอบ
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 40)) ' TextBox ผู้รับผิดชอบ
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))  ' Label กำหนดส่ง
            bottomTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 60)) ' DateTimePicker

            Dim lblAssignedTo As New Label()
            lblAssignedTo.Text = "👤 ผู้รับผิดชอบ:"
            lblAssignedTo.Font = New Font("Segoe UI", 11, FontStyle.Bold)
            lblAssignedTo.ForeColor = Color.FromArgb(52, 73, 94)
            lblAssignedTo.TextAlign = ContentAlignment.MiddleLeft
            lblAssignedTo.Dock = DockStyle.Fill

            Dim txtAssignedTo As New TextBox()
            txtAssignedTo.Text = record.UserName
            txtAssignedTo.Font = New Font("Segoe UI", 9)
            txtAssignedTo.BorderStyle = BorderStyle.FixedSingle
            txtAssignedTo.Dock = DockStyle.Fill
            txtAssignedTo.Margin = New Padding(5, 3, 10, 3)

            Dim lblDueDate As New Label()
            lblDueDate.Text = "📅 กำหนดส่ง:"
            lblDueDate.Font = New Font("Segoe UI", 11, FontStyle.Bold)
            lblDueDate.ForeColor = Color.FromArgb(52, 73, 94)
            lblDueDate.TextAlign = ContentAlignment.MiddleLeft
            lblDueDate.Dock = DockStyle.Fill

            Dim dtpDueDate As New DateTimePicker()
            dtpDueDate.Font = New Font("Segoe UI", 9)
            dtpDueDate.Value = DateTime.Now.AddDays(7)
            dtpDueDate.Format = DateTimePickerFormat.Custom
            dtpDueDate.CustomFormat = "dd/MM/yyyy HH:mm"
            dtpDueDate.Dock = DockStyle.Fill
            dtpDueDate.Margin = New Padding(5, 3, 0, 3)

            ' เพิ่ม controls ลง TableLayoutPanel
            bottomTable.Controls.Add(lblAssignedTo, 0, 0)
            bottomTable.Controls.Add(txtAssignedTo, 1, 0)
            bottomTable.Controls.Add(lblDueDate, 2, 0)
            bottomTable.Controls.Add(dtpDueDate, 3, 0)

            formPanel.Controls.AddRange({lblMissionName, txtMissionName, lblDescription, txtDescription, bottomTable})
            currentY += formPanel.Height + panelMargin

            ' Button Panel - ใช้ FlowLayoutPanel เพื่อจัดเรียงปุ่มอัตโนมัติ
            Dim buttonPanel As New Panel()
            buttonPanel.Size = New Size(formWidth, 60)
            buttonPanel.Location = New Point(15, currentY)
            buttonPanel.BackColor = Color.WhiteSmoke
            buttonPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            Dim buttonFlow As New FlowLayoutPanel()
            buttonFlow.Size = New Size(formWidth - 20, 50)
            buttonFlow.Location = New Point(10, 5)
            buttonFlow.FlowDirection = FlowDirection.RightToLeft  ' จัดเรียงจากขวาไปซ้าย
            buttonFlow.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' ปุ่มยกเลิก
            Dim btnCancel As New Button()
            btnCancel.Text = "❌ ยกเลิก"
            btnCancel.Size = New Size(100, 40)
            btnCancel.BackColor = Color.FromArgb(231, 76, 60)
            btnCancel.ForeColor = Color.White
            btnCancel.FlatStyle = FlatStyle.Flat
            btnCancel.FlatAppearance.BorderSize = 0
            btnCancel.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            btnCancel.DialogResult = DialogResult.Cancel
            btnCancel.Margin = New Padding(5)

            ' ปุ่มสร้าง Mission
            Dim btnConfirm As New Button()
            btnConfirm.Text = "✅ สร้าง Mission"
            btnConfirm.Size = New Size(140, 40)
            btnConfirm.BackColor = Color.FromArgb(39, 174, 96)
            btnConfirm.ForeColor = Color.White
            btnConfirm.FlatStyle = FlatStyle.Flat
            btnConfirm.FlatAppearance.BorderSize = 0
            btnConfirm.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            btnConfirm.DialogResult = DialogResult.OK
            btnConfirm.Margin = New Padding(5)

            ' ปุ่มดูตัวอย่าง
            Dim btnPreview As New Button()
            btnPreview.Text = "👀 ดูตัวอย่าง"
            btnPreview.Size = New Size(120, 40)
            btnPreview.BackColor = Color.FromArgb(155, 89, 182)
            btnPreview.ForeColor = Color.White
            btnPreview.FlatStyle = FlatStyle.Flat
            btnPreview.FlatAppearance.BorderSize = 0
            btnPreview.Font = New Font("Segoe UI", 10, FontStyle.Bold)
            btnPreview.Margin = New Padding(5)
            AddHandler btnPreview.Click, Sub()
                                             Dim preview As String = $"Mission Preview:{vbCrLf}{vbCrLf}" &
                                                               $"ชื่อ: {txtMissionName.Text}{vbCrLf}" &
                                                               $"ผู้รับผิดชอบ: {txtAssignedTo.Text}{vbCrLf}" &
                                                               $"กำหนดส่ง: {dtpDueDate.Value:dd/MM/yyyy HH:mm}{vbCrLf}{vbCrLf}" &
                                                               $"รายละเอียด:{vbCrLf}{txtDescription.Text}"
                                             MessageBox.Show(preview, "ดูตัวอย่าง Mission", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                         End Sub

            ' เพิ่มปุ่มลง FlowLayoutPanel (จะเรียงจากขวาไปซ้าย)
            buttonFlow.Controls.AddRange({btnCancel, btnConfirm, btnPreview})
            buttonPanel.Controls.Add(buttonFlow)

            ' เพิ่ม Panels เข้าฟอร์ม
            missionForm.Controls.AddRange({headerPanel, infoPanel, formPanel, buttonPanel})

            ' เพิ่ม Validation สำหรับข้อมูลที่จำเป็น
            AddHandler btnConfirm.Click, Sub(sender, e)
                                             If String.IsNullOrWhiteSpace(txtMissionName.Text) Then
                                                 MessageBox.Show("กรุณากรอกชื่อ Mission", "ข้อมูลไม่ครบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                 txtMissionName.Focus()
                                                 missionForm.DialogResult = DialogResult.None  ' ป้องกันไม่ให้ปิดฟอร์ม
                                                 Return
                                             End If

                                             If String.IsNullOrWhiteSpace(txtAssignedTo.Text) Then
                                                 MessageBox.Show("กรุณากรอกผู้รับผิดชอบ", "ข้อมูลไม่ครบ", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                 txtAssignedTo.Focus()
                                                 missionForm.DialogResult = DialogResult.None  ' ป้องกันไม่ให้ปิดฟอร์ม
                                                 Return
                                             End If
                                         End Sub

            ' แสดงฟอร์ม
            If missionForm.ShowDialog() = DialogResult.OK Then
                ' สร้าง Mission ID ใหม่
                Dim missionId As String = $"MISSION_{DateTime.Now:yyyyMMddHHmmss}_{record.Id}"

                ' บันทึก RelatedFilePath ลงฐานข้อมูล
                Try
                    If creationCheck.FoundFile IsNot Nothing AndAlso Not String.IsNullOrEmpty(creationCheck.FoundFile.FullPath) Then
                        Dim relatedFilePath As String = creationCheck.FoundFile.FullPath
                        Dim updateSuccess As Boolean = AccessDatabaseManager.UpdateRelatedFilePath(record.Id, relatedFilePath)
                        If updateSuccess Then
                            record.RelatedFilePath = relatedFilePath
                            Console.WriteLine($"Updated RelatedFilePath in database: {relatedFilePath}")
                        Else
                            Console.WriteLine($"Failed to update RelatedFilePath in database")
                        End If
                    End If
                Catch ex As Exception
                    Console.WriteLine($"Error updating RelatedFilePath: {ex.Message}")
                End Try

                ' บันทึกข้อมูล Mission (รวมข้อมูลจาก Excel และไฟล์)
                Dim missionData As String = $"ID: {missionId}{vbCrLf}" &
                                   $"ชื่อ: {txtMissionName.Text}{vbCrLf}" &
                                   $"รายละเอียด: {txtDescription.Text}{vbCrLf}" &
                                   $"ผู้รับผิดชอบ: {txtAssignedTo.Text}{vbCrLf}" &
                                   $"กำหนดส่ง: {dtpDueDate.Value:yyyy-MM-dd HH:mm:ss}{vbCrLf}" &
                                   $"วันที่สร้าง: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{vbCrLf}" &
                                   $"รหัสผลิตภัณฑ์: {record.ProductCode}{vbCrLf}" &
                                   $"ข้อมูล Excel: {creationCheck.ExcelMatch.Column4Value}{vbCrLf}" &
                                   $"ไฟล์เกี่ยวข้อง: {If(creationCheck.FoundFile IsNot Nothing, creationCheck.FoundFile.FullPath, "")}"

                ' บันทึกลงไฟล์
                Try
                    Dim missionDir As String = Path.Combine(Application.StartupPath, "Missions")
                    If Not Directory.Exists(missionDir) Then
                        Directory.CreateDirectory(missionDir)
                    End If

                    Dim missionFile As String = Path.Combine(missionDir, $"{missionId}.txt")
                    File.WriteAllText(missionFile, missionData, Encoding.UTF8)

                    Console.WriteLine($"Mission created: {missionId}")
                Catch ex As Exception
                    Console.WriteLine($"Error saving mission file: {ex.Message}")
                End Try

                MessageBox.Show($"🎉 สร้าง Mission สำเร็จ!{vbCrLf}{vbCrLf}" &
                       $"Mission ID: {missionId}{vbCrLf}" &
                       $"ชื่อ: {txtMissionName.Text}{vbCrLf}" &
                       $"ผู้รับผิดชอบ: {txtAssignedTo.Text}{vbCrLf}" &
                       $"กำหนดส่ง: {dtpDueDate.Value:dd/MM/yyyy HH:mm}{vbCrLf}" &
                       $"ไฟล์เกี่ยวข้อง: {creationCheck.FoundFile.FileName}",
                       "สร้าง Mission สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)

                Return True
            End If

            Return False

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการสร้าง Mission: {ex.Message}",
                   "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in CreateNewMission: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ตรวจสอบความถูกต้องของไฟล์ก่อนยืนยันความสำเร็จ
    ''' </summary>
    Private Function VerifyMissionCompletion(record As ScanDataRecord, scannedWorkpieceCode As String) As Boolean
        Try
            ' อ่านข้อมูล Mission ที่บันทึกไว้
            Dim missionData As Dictionary(Of String, String) = ReadMissionData(record)
            If missionData Is Nothing Then
                MessageBox.Show("ไม่พบข้อมูล Mission", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            ' ดึงชื่อไฟล์จาก Mission data
            Dim missionFilePath As String = ""
            If missionData.ContainsKey("ไฟล์เกี่ยวข้อง") Then
                missionFilePath = missionData("ไฟล์เกี่ยวข้อง")
            ElseIf Not String.IsNullOrEmpty(record.RelatedFilePath) Then
                missionFilePath = record.RelatedFilePath
            End If

            If String.IsNullOrEmpty(missionFilePath) Then
                MessageBox.Show("ไม่พบข้อมูลไฟล์ที่เกี่ยวข้อง", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            ' แยกชื่อไฟล์จาก path (เช่น "SN1B63P101XG-01E_US_VA.pdf" จาก path เต็ม)
            Dim missionFileName As String = Path.GetFileNameWithoutExtension(missionFilePath)

            ' เอาส่วนที่เป็น product code ออกมา (ตัดส่วน "_US_VA" ออก)
            Dim missionProductCode As String = missionFileName.Split("_"c)(0)

            ' เปรียบเทียบกับข้อมูลที่สแกนจากชิ้นงาน
            If scannedWorkpieceCode.Equals(missionProductCode, StringComparison.OrdinalIgnoreCase) Then
                Return True
            Else
                MessageBox.Show($"ไฟล์ไม่ตรงกัน!{vbCrLf}{vbCrLf}" &
                          $"ไฟล์ใน Mission: {missionProductCode}{vbCrLf}" &
                          $"ข้อมูลที่สแกน: {scannedWorkpieceCode}",
                          "การตรวจสอบไม่ผ่าน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตรวจสอบ: {ex.Message}",
                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    Private Function CheckMissionStatus(record As ScanDataRecord, rowIndex As Integer) As String
        Try
            If record Is Nothing Then
                MessageBox.Show("ไม่พบข้อมูลการสแกน", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return ""
            End If

            ' สร้างฟอร์มแสดงสถานะ Mission ที่ปรับปรุงแล้ว
            Dim statusForm As New Form()
            statusForm.Text = "ตรวจสอบสถานะ Mission"
            statusForm.Size = New Size(900, 750)  ' เพิ่มขนาด
            statusForm.StartPosition = FormStartPosition.CenterParent
            statusForm.FormBorderStyle = FormBorderStyle.FixedDialog
            statusForm.MaximizeBox = False
            statusForm.MinimizeBox = False
            statusForm.BackColor = Color.WhiteSmoke
            statusForm.Font = New Font("Segoe UI", 9)
            statusForm.Padding = New Padding(15)

            ' ตัวแปรสำหรับจัดการ layout แบบ responsive
            Dim formWidth As Integer = statusForm.ClientSize.Width - 30
            Dim panelMargin As Integer = 10
            Dim currentY As Integer = panelMargin

            ' Header Panel
            Dim headerPanel As New Panel()
            headerPanel.Size = New Size(formWidth, 80)
            headerPanel.Location = New Point(15, currentY)
            headerPanel.BackColor = Color.FromArgb(155, 89, 182)
            headerPanel.BorderStyle = BorderStyle.None
            headerPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            Dim lblTitle As New Label()
            lblTitle.Text = "📋 ตรวจสอบสถานะ Mission"
            lblTitle.Font = New Font("Segoe UI", 16, FontStyle.Bold)
            lblTitle.Location = New Point(20, 15)
            lblTitle.AutoSize = True
            lblTitle.ForeColor = Color.White
            lblTitle.BackColor = Color.Transparent

            Dim lblStatus As New Label()
            lblStatus.Text = $"สถานะปัจจุบัน: {record.MissionStatus}"
            lblStatus.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            lblStatus.Location = New Point(20, 50)
            lblStatus.AutoSize = True
            lblStatus.ForeColor = Color.White
            lblStatus.BackColor = Color.Transparent

            headerPanel.Controls.AddRange({lblTitle, lblStatus})
            currentY += headerPanel.Height + panelMargin

            ' Info Panel - ใช้ TableLayoutPanel
            Dim infoPanel As New Panel()
            infoPanel.Size = New Size(formWidth, 200)
            infoPanel.Location = New Point(15, currentY)
            infoPanel.BackColor = Color.White
            infoPanel.BorderStyle = BorderStyle.FixedSingle
            infoPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            Dim infoTable As New TableLayoutPanel()
            infoTable.Size = New Size(formWidth - 20, 180)
            infoTable.Location = New Point(10, 10)
            infoTable.ColumnCount = 2
            infoTable.RowCount = 8
            infoTable.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' กำหนดขนาดคอลัมน์
            infoTable.ColumnStyles.Add(New ColumnStyle(SizeType.AutoSize))
            infoTable.ColumnStyles.Add(New ColumnStyle(SizeType.Percent, 100))

            ' กำหนดขนาดแถว
            For i As Integer = 0 To 7
                infoTable.RowStyles.Add(New RowStyle(SizeType.Absolute, 22))
            Next

            ' สร้าง Labels สำหรับข้อมูล
            Dim infoLabels() As (icon As String, text As String, value As String) = {
            ("🔍", "รหัสผลิตภัณฑ์:", record.ProductCode),
            ("📅", "วันที่สแกน:", record.ScanDateTime.ToString("yyyy-MM-dd HH:mm:ss")),
            ("📋", "รหัสอ้างอิง:", record.ReferenceCode),
            ("🔢", "จำนวน:", record.Quantity.ToString()),
            ("📅", "วันที่ผลิต:", record.DateCode),
            ("👤", "ผู้ใช้:", record.UserName),
            ("💻", "เครื่อง:", record.ComputerName),
            ("✅", "สถานะข้อมูล:", If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง"))
        }

            For i As Integer = 0 To infoLabels.Length - 1
                Dim lblKey As New Label()
                lblKey.Text = $"{infoLabels(i).icon} {infoLabels(i).text}"
                lblKey.Font = New Font("Segoe UI", 10, FontStyle.Bold)
                lblKey.ForeColor = Color.FromArgb(52, 73, 94)
                lblKey.TextAlign = ContentAlignment.MiddleLeft
                lblKey.Dock = DockStyle.Fill

                Dim lblValue As New Label()
                lblValue.Text = infoLabels(i).value
                lblValue.Font = New Font("Segoe UI", 10)
                lblValue.ForeColor = Color.FromArgb(44, 62, 80)
                lblValue.TextAlign = ContentAlignment.MiddleLeft
                lblValue.Dock = DockStyle.Fill

                infoTable.Controls.Add(lblKey, 0, i)
                infoTable.Controls.Add(lblValue, 1, i)
            Next

            infoPanel.Controls.Add(infoTable)
            currentY += infoPanel.Height + panelMargin

            ' Excel & File Panel - ปรับปรุงให้ดูไฟล์ได้เหมือน Mission Form
            Dim filePanel As New Panel()
            filePanel.Size = New Size(formWidth, 160)
            filePanel.Location = New Point(15, currentY)
            filePanel.BackColor = Color.FromArgb(248, 249, 250)
            filePanel.BorderStyle = BorderStyle.FixedSingle
            filePanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            Dim lblFileTitle As New Label()
            lblFileTitle.Text = "📁 ข้อมูล Excel และไฟล์ที่เกี่ยวข้อง"
            lblFileTitle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            lblFileTitle.Location = New Point(15, 15)
            lblFileTitle.AutoSize = True
            lblFileTitle.ForeColor = Color.FromArgb(52, 73, 94)

            ' ค้นหาข้อมูล Excel และไฟล์ที่เกี่ยวข้องจาก Mission log
            Dim excelMatch As ExcelUtility.ExcelMatchResult = Nothing
            Dim relatedFilePath As String = Nothing
            Dim excelInfo As String = "ข้อมูลจาก RelatedFilePath ในฐานข้อมูล"

            ' ลองค้นหาข้อมูล Excel ถ้าต้องการ
            Try
                If dataCache.IsLoaded Then
                    Dim searchResult = dataCache.SearchInMemory(record.ProductCode)
                    If searchResult.IsSuccess AndAlso searchResult.HasMatches Then
                        excelMatch = searchResult.FirstMatch
                        excelInfo = $"ข้อมูลจาก Excel: {excelMatch.Column4Value}"
                    End If
                End If
            Catch ex As Exception
                Console.WriteLine($"Error searching Excel data: {ex.Message}")
            End Try

            ' อ่านข้อมูลจาก Mission file ที่บันทึกไว้
            Dim missionData As Dictionary(Of String, String) = ReadMissionData(record)

            If missionData IsNot Nothing Then
                ' ใช้ข้อมูลจาก Mission file ที่บันทึกไว้แล้ว
                If missionData.ContainsKey("ข้อมูล Excel") Then
                    excelInfo = $"ข้อมูลจาก Excel: {missionData("ข้อมูล Excel")}"
                End If

                ' ดึง path ไฟล์จาก Mission log
                If missionData.ContainsKey("ไฟล์เกี่ยวข้อง") Then
                    relatedFilePath = missionData("ไฟล์เกี่ยวข้อง")
                    Console.WriteLine($"พบไฟล์จาก Mission log: {relatedFilePath}")
                End If
            Else
                ' ถ้าไม่มี Mission file ให้ลองค้นหาจาก dataCache
                Try
                    If dataCache.IsLoaded Then
                        Dim searchResult = dataCache.SearchInMemory(record.ProductCode)
                        If searchResult.IsSuccess AndAlso searchResult.HasMatches Then
                            excelMatch = searchResult.FirstMatch
                            excelInfo = $"ข้อมูลจาก Excel: {excelMatch.Column4Value}"
                            Console.WriteLine($"ไม่มี Mission file, ค้นหาจาก dataCache: {excelMatch.Column4Value}")
                        End If
                    End If
                Catch ex As Exception
                    excelInfo = $"ข้อผิดพลาดในการค้นหา: {ex.Message}"
                    Console.WriteLine($"Error searching Excel data: {ex.Message}")
                End Try
            End If

            ' แสดงข้อมูล Excel
            Dim lblExcelInfo As New Label()
            lblExcelInfo.Text = $"📊 {excelInfo}"
            lblExcelInfo.Location = New Point(15, 50)
            lblExcelInfo.Size = New Size(formWidth - 140, 25)
            lblExcelInfo.Font = New Font("Segoe UI", 10)
            lblExcelInfo.ForeColor = Color.FromArgb(39, 174, 96)

            ' แสดงข้อมูลไฟล์
            Dim fileInfo As String = If(Not String.IsNullOrEmpty(relatedFilePath),
                                   $"ไฟล์ที่เกี่ยวข้อง: {Path.GetFileName(relatedFilePath)}",
                                   "ไม่พบไฟล์ที่เกี่ยวข้อง")

            Dim lblFileInfo As New Label()
            lblFileInfo.Text = $"📄 {fileInfo}"
            lblFileInfo.Location = New Point(15, 80)
            lblFileInfo.Size = New Size(formWidth - 140, 25)
            lblFileInfo.Font = New Font("Segoe UI", 10)
            lblFileInfo.ForeColor = Color.FromArgb(41, 128, 185)

            ' ปุ่มดูไฟล์ (เหมือนใน Mission Form)
            Dim btnPreviewFile As New Button()
            btnPreviewFile.Text = "👁️ ดูไฟล์"
            btnPreviewFile.Location = New Point(formWidth - 110, 77)
            btnPreviewFile.Size = New Size(100, 30)
            btnPreviewFile.BackColor = Color.FromArgb(52, 152, 219)
            btnPreviewFile.ForeColor = Color.White
            btnPreviewFile.FlatStyle = FlatStyle.Flat
            btnPreviewFile.FlatAppearance.BorderSize = 0
            btnPreviewFile.Font = New Font("Segoe UI", 9, FontStyle.Bold)

            ' Enable ถ้ามีไฟล์ path และไฟล์มีอยู่จริง
            Dim fileExists As Boolean = Not String.IsNullOrEmpty(relatedFilePath) AndAlso File.Exists(relatedFilePath)
            btnPreviewFile.Enabled = True

            ' เปลี่ยนสีปุ่มตามสถานะไฟล์
            If fileExists Then
                btnPreviewFile.BackColor = Color.FromArgb(52, 152, 219)  ' สีฟ้า = ไฟล์มีอยู่
            Else
                btnPreviewFile.BackColor = Color.FromArgb(149, 165, 166)  ' สีเทา = ไฟล์ไม่มี
            End If

            AddHandler btnPreviewFile.Click, Sub()
                                                 If Not String.IsNullOrEmpty(relatedFilePath) Then
                                                     If File.Exists(relatedFilePath) Then
                                                         Console.WriteLine($"เปิดไฟล์จากฐานข้อมูล: {relatedFilePath}")
                                                         OpenFileWithErrorHandling(relatedFilePath)
                                                     Else
                                                         MessageBox.Show($"ไม่พบไฟล์ที่ระบุ:{vbCrLf}{relatedFilePath}{vbCrLf}{vbCrLf}ไฟล์อาจถูกย้าย, ลบ, หรือเครือข่ายไม่เชื่อมต่อ",
                                                                    "ไม่พบไฟล์", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                     End If
                                                 Else
                                                     ' แสดง message และเสนอทางเลือก
                                                     Dim message As String = $"ไม่มีข้อมูลไฟล์ที่เกี่ยวข้องสำหรับ {record.ProductCode} ในฐานข้อมูล"
                                                     message += vbCrLf & vbCrLf & "ต้องการเปิดโฟลเดอร์เครือข่ายเพื่อค้นหาไฟล์เองหรือไม่?"

                                                     If MessageBox.Show(message, "ไม่มีข้อมูลไฟล์", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                                         Try
                                                             Dim networkFolder As String = "\\10.24.179.2\OAFAB\OA2FAB\Film charecter check"
                                                             Process.Start("explorer.exe", networkFolder)
                                                         Catch ex As Exception
                                                             MessageBox.Show($"ไม่สามารถเปิดโฟลเดอร์เครือข่ายได้: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                                                         End Try
                                                     End If
                                                 End If
                                             End Sub

            ' ค้นหาไฟล์ Mission ที่เกี่ยวข้อง
            Dim missionFiles = FindMissionFiles(record)
            Dim missionFileInfo As String = If(missionFiles.Count > 0,
                                          $"Mission Files: {missionFiles.Count} ไฟล์",
                                          "ไม่พบไฟล์ Mission")

            Dim lblMissionFiles As New Label()
            lblMissionFiles.Text = $"📋 {missionFileInfo}"
            lblMissionFiles.Location = New Point(15, 110)
            lblMissionFiles.Size = New Size(formWidth - 220, 25)
            lblMissionFiles.Font = New Font("Segoe UI", 10)
            lblMissionFiles.ForeColor = Color.FromArgb(142, 68, 173)

            ' ปุ่มเปิดไฟล์ Mission
            Dim btnOpenMission As New Button()
            btnOpenMission.Text = "📋 Mission"
            btnOpenMission.Location = New Point(formWidth - 210, 107)
            btnOpenMission.Size = New Size(90, 30)
            btnOpenMission.BackColor = Color.FromArgb(142, 68, 173)
            btnOpenMission.ForeColor = Color.White
            btnOpenMission.FlatStyle = FlatStyle.Flat
            btnOpenMission.FlatAppearance.BorderSize = 0
            btnOpenMission.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnOpenMission.Enabled = missionFiles.Count > 0
            AddHandler btnOpenMission.Click, Sub()
                                                 If missionFiles.Count > 0 Then
                                                     OpenFileWithErrorHandling(missionFiles(0).FullName)
                                                 End If
                                             End Sub

            ' ปุ่มเปิดโฟลเดอร์
            Dim btnOpenFolder As New Button()
            btnOpenFolder.Text = "📁 โฟลเดอร์"
            btnOpenFolder.Location = New Point(formWidth - 110, 107)
            btnOpenFolder.Size = New Size(100, 30)
            btnOpenFolder.BackColor = Color.FromArgb(46, 204, 113)
            btnOpenFolder.ForeColor = Color.White
            btnOpenFolder.FlatStyle = FlatStyle.Flat
            btnOpenFolder.FlatAppearance.BorderSize = 0
            btnOpenFolder.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            AddHandler btnOpenFolder.Click, Sub()
                                                OpenMissionFolder(record)
                                            End Sub

            filePanel.Controls.AddRange({lblFileTitle, lblExcelInfo, lblFileInfo, btnPreviewFile,
                                   lblMissionFiles, btnOpenMission, btnOpenFolder})
            currentY += filePanel.Height + panelMargin

            ' Action Panel - ใช้ FlowLayoutPanel
            Dim actionPanel As New Panel()
            actionPanel.Size = New Size(formWidth, 100)
            actionPanel.Location = New Point(15, currentY)
            actionPanel.BackColor = Color.White
            actionPanel.BorderStyle = BorderStyle.FixedSingle
            actionPanel.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            Dim lblActionTitle As New Label()
            lblActionTitle.Text = "⚡ การดำเนินการ"
            lblActionTitle.Font = New Font("Segoe UI", 12, FontStyle.Bold)
            lblActionTitle.Location = New Point(15, 15)
            lblActionTitle.AutoSize = True
            lblActionTitle.ForeColor = Color.FromArgb(52, 73, 94)

            ' ใช้ FlowLayoutPanel สำหรับปุ่มต่างๆ
            Dim actionFlow As New FlowLayoutPanel()
            actionFlow.Location = New Point(15, 50)
            actionFlow.Size = New Size(formWidth - 30, 40)
            actionFlow.FlowDirection = FlowDirection.LeftToRight
            actionFlow.Anchor = AnchorStyles.Top Or AnchorStyles.Left Or AnchorStyles.Right

            ' ปุ่มเปลี่ยนสถานะเป็น "สำเร็จ"
            Dim btnMarkComplete As New Button()
            btnMarkComplete.Text = "✅ ทำเครื่องหมายสำเร็จ"
            btnMarkComplete.Size = New Size(160, 35)
            btnMarkComplete.BackColor = Color.FromArgb(39, 174, 96)
            btnMarkComplete.ForeColor = Color.White
            btnMarkComplete.FlatStyle = FlatStyle.Flat
            btnMarkComplete.FlatAppearance.BorderSize = 0
            btnMarkComplete.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnMarkComplete.Margin = New Padding(0, 0, 10, 0)

            ' ปุ่มรีเซ็ตสถานะ
            Dim btnReset As New Button()
            btnReset.Text = "🔄 รีเซ็ตสถานะ"
            btnReset.Size = New Size(120, 35)
            btnReset.BackColor = Color.FromArgb(230, 126, 34)
            btnReset.ForeColor = Color.White
            btnReset.FlatStyle = FlatStyle.Flat
            btnReset.FlatAppearance.BorderSize = 0
            btnReset.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnReset.Margin = New Padding(0, 0, 10, 0)

            ' ปุ่มดูรายละเอียด Mission
            Dim btnViewDetails As New Button()
            btnViewDetails.Text = "📄 ดูรายละเอียด"
            btnViewDetails.Size = New Size(120, 35)
            btnViewDetails.BackColor = Color.FromArgb(52, 152, 219)
            btnViewDetails.ForeColor = Color.White
            btnViewDetails.FlatStyle = FlatStyle.Flat
            btnViewDetails.FlatAppearance.BorderSize = 0
            btnViewDetails.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnViewDetails.Margin = New Padding(0, 0, 10, 0)

            ' ปุ่มปิด
            Dim btnClose As New Button()
            btnClose.Text = "❌ ปิด"
            btnClose.Size = New Size(80, 35)
            btnClose.BackColor = Color.FromArgb(149, 165, 166)
            btnClose.ForeColor = Color.White
            btnClose.FlatStyle = FlatStyle.Flat
            btnClose.FlatAppearance.BorderSize = 0
            btnClose.Font = New Font("Segoe UI", 9, FontStyle.Bold)
            btnClose.DialogResult = DialogResult.Cancel

            actionFlow.Controls.AddRange({btnMarkComplete, btnReset, btnViewDetails, btnClose})
            actionPanel.Controls.AddRange({lblActionTitle, actionFlow})

            ' Event Handlers
            AddHandler btnMarkComplete.Click, Sub()
                                                  ' เพิ่มการสแกนรหัสชิ้นงานเพื่อยืนยันก่อน
                                                  Dim scannedCode As String = InputBox("กรุณาสแกนรหัสชิ้นงานเพื่อยืนยัน:", "ตรวจสอบชิ้นงาน")

                                                  If Not String.IsNullOrEmpty(scannedCode) Then
                                                      ' เรียกฟังก์ชันตรวจสอบไฟล์
                                                      If VerifyMissionCompletion(record, scannedCode) Then
                                                          If MessageBox.Show("ยืนยันการทำเครื่องหมายว่า Mission นี้สำเร็จแล้ว?",
                              "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                                              record.MissionStatus = "สำเร็จ"
                                                              If UpdateMissionStatus(record) Then
                                                                  ' อัปเดตการแสดงผลในตาราง
                                                                  dgvHistory.Rows(rowIndex).Cells("MissionStatus").Value = "สำเร็จ"
                                                                  dgvHistory.Rows(rowIndex).Cells("btnCreateMission").Value = "✅ สำเร็จ"
                                                                  dgvHistory.Rows(rowIndex).Cells("btnCreateMission").Style.ForeColor = Color.Green

                                                                  MessageBox.Show("อัปเดตสถานะ Mission เป็น 'สำเร็จ' แล้ว",
                                   "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                                  statusForm.Close()
                                                              End If
                                                          End If
                                                      End If
                                                  Else
                                                      MessageBox.Show("ต้องสแกนรหัสชิ้นงานก่อนยืนยัน", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                                                  End If
                                              End Sub

            AddHandler btnReset.Click, Sub()
                                           If MessageBox.Show("ยืนยันการรีเซ็ตสถานะ Mission กลับเป็น 'ไม่มี'?",
                                                        "ยืนยัน", MessageBoxButtons.YesNo, MessageBoxIcon.Question) = DialogResult.Yes Then
                                               record.MissionStatus = "ไม่มี"
                                               If UpdateMissionStatus(record) Then
                                                   ' อัปเดตการแสดงผลในตาราง
                                                   dgvHistory.Rows(rowIndex).Cells("MissionStatus").Value = "ไม่มี"
                                                   dgvHistory.Rows(rowIndex).Cells("btnCreateMission").Value = "🚀 สร้าง"
                                                   dgvHistory.Rows(rowIndex).Cells("btnCreateMission").Style.ForeColor = Color.Blue

                                                   MessageBox.Show("รีเซ็ตสถานะ Mission แล้ว",
                                                              "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
                                                   statusForm.Close()
                                               End If
                                           End If
                                       End Sub

            AddHandler btnViewDetails.Click, Sub()
                                                 ShowMissionDetailsDialog(record)
                                             End Sub

            ' เพิ่ม Panels เข้าฟอร์ม
            statusForm.Controls.AddRange({headerPanel, infoPanel, filePanel, actionPanel})

            ' แสดงฟอร์ม
            statusForm.ShowDialog()

            Return record.MissionStatus

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการตรวจสอบสถานะ Mission: {ex.Message}",
                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in CheckMissionStatus: {ex.Message}")
            Return ""
        End Try
    End Function

    Private Function UpdateMissionStatus(record As ScanDataRecord) As Boolean
        Try
            If record Is Nothing Then
                Console.WriteLine("Record is null in UpdateMissionStatus")
                Return False
            End If

            ' อัปเดตสถานะในฐานข้อมูล Access
            Dim success As Boolean = AccessDatabaseManager.UpdateMissionStatus(record.Id, record.MissionStatus)

            If success Then
                Console.WriteLine($"Mission status updated for record ID {record.Id}: {record.MissionStatus}")
                Return True
            Else
                Console.WriteLine($"Failed to update mission status for record ID {record.Id}")
                Return False
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in UpdateMissionStatus: {ex.Message}")
            MessageBox.Show($"เกิดข้อผิดพลาดในการอัปเดตสถานะ Mission: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' อ่านข้อมูลจาก Mission file
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>Dictionary ของข้อมูล Mission หรือ Nothing ถ้าไม่พบ</returns>
    Private Function ReadMissionData(record As ScanDataRecord) As Dictionary(Of String, String)
        Try
            ' ค้นหาไฟล์ Mission
            Dim missionFiles = FindMissionFiles(record)
            If missionFiles.Count = 0 Then
                Console.WriteLine($"ไม่พบไฟล์ Mission สำหรับ record ID: {record.Id}")
                Return Nothing
            End If

            ' อ่านไฟล์ Mission แรกที่เจอ
            Dim missionFile = missionFiles(0)
            Console.WriteLine($"อ่านไฟล์ Mission: {missionFile.FullName}")

            Dim content As String = File.ReadAllText(missionFile.FullName, Encoding.UTF8)
            Dim data As New Dictionary(Of String, String)()

            ' แยกข้อมูลจากไฟล์ (รูปแบบ "key: value")
            Dim lines() As String = content.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)

            For Each line In lines
                If line.Contains(":") Then
                    Dim parts() As String = line.Split({":", "="}, 2, StringSplitOptions.None)
                    If parts.Length = 2 Then
                        Dim key As String = parts(0).Trim()
                        Dim value As String = parts(1).Trim()
                        data(key) = value
                        Console.WriteLine($"Mission data: {key} = {value}")
                    End If
                End If
            Next

            Return data

        Catch ex As Exception
            Console.WriteLine($"Error reading Mission data: {ex.Message}")
            Return Nothing
        End Try
    End Function


    Public Function FindMissionFiles(record As ScanDataRecord) As List(Of FileInfo)
        Dim files As New List(Of FileInfo)()

        Try
            ' ค้นหาไฟล์ Mission
            Dim missionDir As String = Path.Combine(Application.StartupPath, "Missions")
            If Directory.Exists(missionDir) Then
                Dim missionPattern As String = $"MISSION_*_{record.Id}.txt"
                Dim missionFiles = Directory.GetFiles(missionDir, missionPattern)

                For Each file In missionFiles
                    files.Add(New FileInfo(file))
                Next

                ' ถ้าไม่เจอแบบ specific ให้ลองค้นหาแบบ general
                If files.Count = 0 Then
                    Dim generalPattern As String = $"*{record.ProductCode}*"
                    Dim generalFiles = Directory.GetFiles(missionDir, generalPattern)
                    For Each file In generalFiles
                        files.Add(New FileInfo(file))
                    Next
                End If
            End If

        Catch ex As Exception
            Console.WriteLine($"Error in FindMissionFiles: {ex.Message}")
        End Try

        Return files
    End Function

    Private Sub OpenMissionFolder(record As ScanDataRecord)
        Try
            Dim missionDir As String = Path.Combine(Application.StartupPath, "Missions")
            If Directory.Exists(missionDir) Then
                Process.Start("explorer.exe", missionDir)
            Else
                MessageBox.Show("ไม่พบโฟลเดอร์ Mission", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Warning)
            End If
        Catch ex As Exception
            MessageBox.Show($"ไม่สามารถเปิดโฟลเดอร์ได้: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub


    ''' <summary>
    ''' แสดงรายละเอียด Mission ที่เสร็จสิ้นแล้ว
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกนที่มี Mission เสร็จสิ้น</param>
    Private Sub ShowCompletedMissionDetails(record As ScanDataRecord)
        Try
            If record Is Nothing Then
                MessageBox.Show("ไม่พบข้อมูลการสแกน", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return
            End If

            If record.MissionStatus <> "สำเร็จ" Then
                MessageBox.Show($"Mission นี้ยังไม่เสร็จสิ้น (สถานะปัจจุบัน: {record.MissionStatus})",
                               "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Warning)
                Return
            End If

            ' สร้างฟอร์มแสดงรายละเอียด
            Dim detailForm As New Form()
            detailForm.Text = "รายละเอียด Mission ที่เสร็จสิ้น"
            detailForm.Size = New Size(700, 600)
            detailForm.StartPosition = FormStartPosition.CenterParent
            detailForm.FormBorderStyle = FormBorderStyle.FixedDialog
            detailForm.MaximizeBox = False
            detailForm.MinimizeBox = False

            ' สร้าง TabControl
            Dim tabControl As New TabControl()
            tabControl.Dock = DockStyle.Fill

            ' Tab 1: ข้อมูลทั่วไป
            Dim tabGeneral As New TabPage("ข้อมูลทั่วไป")
            Dim txtGeneral As New TextBox()
            txtGeneral.Multiline = True
            txtGeneral.ReadOnly = True
            txtGeneral.ScrollBars = ScrollBars.Vertical
            txtGeneral.Dock = DockStyle.Fill
            txtGeneral.Font = New Font("Consolas", 10)

            txtGeneral.Text = $"🎉 Mission เสร็จสิ้นแล้ว!{vbCrLf}{vbCrLf}" &
                             $"=== ข้อมูล Mission ==={vbCrLf}" &
                             $"🆔 รหัสอ้างอิง: MISSION_{record.ScanDateTime:yyyyMMddHHmmss}_{record.Id}{vbCrLf}" &
                             $"📊 สถานะ: {record.MissionStatus}{vbCrLf}" &
                             $"📅 วันที่สร้าง: {record.ScanDateTime:yyyy-MM-dd HH:mm:ss}{vbCrLf}" &
                             $"📅 วันที่เสร็จสิ้น: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{vbCrLf}{vbCrLf}" &
                             $"=== ข้อมูลการสแกน ==={vbCrLf}" &
                             $"🔍 รหัสผลิตภัณฑ์: {record.ProductCode}{vbCrLf}" &
                             $"📋 รหัสอ้างอิง: {record.ReferenceCode}{vbCrLf}" &
                             $"🔢 จำนวน: {record.Quantity}{vbCrLf}" &
                             $"📅 วันที่ผลิต: {record.DateCode}{vbCrLf}" &
                             $"✅ สถานะข้อมูล: {If(record.IsValid, "ถูกต้อง", "ไม่ถูกต้อง")}{vbCrLf}{vbCrLf}" &
                             $"=== ผู้ดำเนินการ ==={vbCrLf}" &
                             $"👤 ผู้ใช้: {record.UserName}{vbCrLf}" &
                             $"💻 เครื่อง: {record.ComputerName}{vbCrLf}{vbCrLf}" &
                             $"=== เวลาดำเนินการ ==={vbCrLf}" &
                             $"⏰ ระยะเวลา: {Math.Round((DateTime.Now - record.ScanDateTime).TotalMinutes, 1)} นาที"

            tabGeneral.Controls.Add(txtGeneral)
            tabControl.TabPages.Add(tabGeneral)

            ' Tab 2: ข้อมูลต้นฉบับ
            Dim tabRaw As New TabPage("ข้อมูลต้นฉบับ")
            Dim txtRaw As New TextBox()
            txtRaw.Multiline = True
            txtRaw.ReadOnly = True
            txtRaw.ScrollBars = ScrollBars.Both
            txtRaw.Dock = DockStyle.Fill
            txtRaw.Font = New Font("Consolas", 9)
            txtRaw.Text = $"ข้อมูลดิบจาก QR Code:{vbCrLf}{vbCrLf}{record.OriginalData}"
            tabRaw.Controls.Add(txtRaw)
            tabControl.TabPages.Add(tabRaw)

            ' Panel สำหรับปุ่ม
            Dim buttonPanel As New Panel()
            buttonPanel.Dock = DockStyle.Bottom
            buttonPanel.Height = 60

            Dim btnExportReport As New Button()
            btnExportReport.Text = "📄 ส่งออกรายงาน"
            btnExportReport.Location = New Point(20, 15)
            btnExportReport.Size = New Size(150, 30)
            btnExportReport.BackColor = Color.Blue
            btnExportReport.ForeColor = Color.White
            btnExportReport.FlatStyle = FlatStyle.Flat

            Dim btnPrintReport As New Button()
            btnPrintReport.Text = "🖨️ พิมพ์รายงาน"
            btnPrintReport.Location = New Point(180, 15)
            btnPrintReport.Size = New Size(120, 30)
            btnPrintReport.BackColor = Color.Green
            btnPrintReport.ForeColor = Color.White
            btnPrintReport.FlatStyle = FlatStyle.Flat

            Dim btnCloseDetail As New Button()
            btnCloseDetail.Text = "❌ ปิด"
            btnCloseDetail.Location = New Point(600, 15)
            btnCloseDetail.Size = New Size(70, 30)
            btnCloseDetail.BackColor = Color.Gray
            btnCloseDetail.ForeColor = Color.White
            btnCloseDetail.FlatStyle = FlatStyle.Flat
            btnCloseDetail.DialogResult = DialogResult.OK

            ' Event Handlers
            AddHandler btnExportReport.Click, Sub()
                                                  ExportMissionReport(record)
                                              End Sub

            AddHandler btnPrintReport.Click, Sub()
                                                 PrintMissionReport(record)
                                             End Sub

            buttonPanel.Controls.AddRange({btnExportReport, btnPrintReport, btnCloseDetail})

            ' เพิ่ม Controls เข้าฟอร์ม
            detailForm.Controls.Add(tabControl)
            detailForm.Controls.Add(buttonPanel)

            ' แสดงฟอร์ม
            detailForm.ShowDialog()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด Mission: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in ShowCompletedMissionDetails: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' แสดงรายละเอียด Mission ในกล่องโต้ตอบเล็ก
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    Private Sub ShowMissionDetailsDialog(record As ScanDataRecord)
        Try
            Dim missionId As String = $"MISSION_{record.ScanDateTime:yyyyMMddHHmmss}_{record.Id}"
            Dim missionFile As String = Path.Combine(Application.StartupPath, "Missions", $"{missionId}.txt")

            Dim details As String = ""
            If File.Exists(missionFile) Then
                details = File.ReadAllText(missionFile, Encoding.UTF8)
            Else
                details = $"📋 ข้อมูล Mission{vbCrLf}{vbCrLf}" &
                         $"ID: {missionId}{vbCrLf}" &
                         $"สถานะ: {record.MissionStatus}{vbCrLf}" &
                         $"รหัสผลิตภัณฑ์: {record.ProductCode}{vbCrLf}" &
                         $"วันที่สร้าง: {record.ScanDateTime:yyyy-MM-dd HH:mm:ss}"
            End If

            MessageBox.Show(details, "รายละเอียด Mission", MessageBoxButtons.OK, MessageBoxIcon.Information)

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายละเอียด: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ส่งออกรายงาน Mission
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    Private Sub ExportMissionReport(record As ScanDataRecord)
        Try
            Dim saveDialog As New SaveFileDialog()
            saveDialog.Filter = "Text Files (*.txt)|*.txt|All Files (*.*)|*.*"
            saveDialog.FileName = $"MissionReport_{record.ProductCode}_{DateTime.Now:yyyyMMdd}.txt"

            If saveDialog.ShowDialog() = DialogResult.OK Then
                Dim reportContent As String = GenerateMissionReport(record)
                File.WriteAllText(saveDialog.FileName, reportContent, Encoding.UTF8)

                MessageBox.Show($"ส่งออกรายงานสำเร็จ!{vbCrLf}ไฟล์: {saveDialog.FileName}",
                               "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการส่งออกรายงาน: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' พิมพ์รายงาน Mission
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    Private Sub PrintMissionReport(record As ScanDataRecord)
        Try
            Dim reportContent As String = GenerateMissionReport(record)

            ' สร้าง PrintDocument
            Dim printDoc As New System.Drawing.Printing.PrintDocument()
            Dim reportText As String = reportContent

            AddHandler printDoc.PrintPage, Sub(sender As Object, e As System.Drawing.Printing.PrintPageEventArgs)
                                               Dim font As New Font("Arial", 10)
                                               Dim brush As New SolidBrush(Color.Black)
                                               Dim leftMargin As Single = e.MarginBounds.Left
                                               Dim topMargin As Single = e.MarginBounds.Top
                                               Dim lineHeight As Single = font.GetHeight(e.Graphics)

                                               Dim lines() As String = reportText.Split({vbCrLf, vbLf}, StringSplitOptions.None)
                                               Dim yPos As Single = topMargin

                                               For Each line As String In lines
                                                   If yPos + lineHeight > e.MarginBounds.Bottom Then
                                                       e.HasMorePages = True
                                                       Exit For
                                                   End If

                                                   e.Graphics.DrawString(line, font, brush, leftMargin, yPos)
                                                   yPos += lineHeight
                                               Next

                                               font.Dispose()
                                               brush.Dispose()
                                           End Sub

            ' แสดง Print Dialog
            Dim printDialog As New PrintDialog()
            printDialog.Document = printDoc

            If printDialog.ShowDialog() = DialogResult.OK Then
                printDoc.Print()
                MessageBox.Show("พิมพ์รายงานสำเร็จ!", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการพิมพ์รายงาน: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' สร้างเนื้อหารายงาน Mission
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>เนื้อหารายงาน</returns>
    Private Function GenerateMissionReport(record As ScanDataRecord) As String
        Try
            Dim report As New StringBuilder()

            report.AppendLine("=".PadRight(80, "="c))
            report.AppendLine("                    รายงาน MISSION ที่เสร็จสิ้น")
            report.AppendLine("=".PadRight(80, "="c))
            report.AppendLine()

            report.AppendLine($"📅 วันที่สร้างรายงาน: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            report.AppendLine($"🆔 Mission ID: MISSION_{record.ScanDateTime:yyyyMMddHHmmss}_{record.Id}")
            report.AppendLine()

            report.AppendLine("📊 ข้อมูล Mission:")
            report.AppendLine("-".PadRight(50, "-"c))
            report.AppendLine($"   สถานะ Mission: {record.MissionStatus}")
            report.AppendLine($"   วันที่เริ่มต้น: {record.ScanDateTime:yyyy-MM-dd HH:mm:ss}")
            report.AppendLine($"   วันที่เสร็จสิ้น: {DateTime.Now:yyyy-MM-dd HH:mm:ss}")
            report.AppendLine($"   ระยะเวลาดำเนินการ: {Math.Round((DateTime.Now - record.ScanDateTime).TotalHours, 2)} ชั่วโมง")
            report.AppendLine()

            report.AppendLine("🔍 ข้อมูลการสแกน:")
            report.AppendLine("-".PadRight(50, "-"c))
            report.AppendLine($"   รหัสผลิตภัณฑ์: {record.ProductCode}")
            report.AppendLine($"   รหัสอ้างอิง: {record.ReferenceCode}")
            report.AppendLine($"   จำนวน: {record.Quantity}")
            report.AppendLine($"   วันที่ผลิต: {record.DateCode}")
            report.AppendLine($"   สถานะข้อมูล: {If(record.IsValid, "✅ ถูกต้อง", "❌ ไม่ถูกต้อง")}")
            report.AppendLine()

            report.AppendLine("👤 ข้อมูลผู้ดำเนินการ:")
            report.AppendLine("-".PadRight(50, "-"c))
            report.AppendLine($"   ผู้ใช้: {record.UserName}")
            report.AppendLine($"   เครื่อง: {record.ComputerName}")
            report.AppendLine()

            report.AppendLine("📋 ข้อมูลต้นฉบับ:")
            report.AppendLine("-".PadRight(50, "-"c))
            report.AppendLine($"   {record.OriginalData}")
            report.AppendLine()

            report.AppendLine("=".PadRight(80, "="c))
            report.AppendLine("                       สิ้นสุดรายงาน")
            report.AppendLine("=".PadRight(80, "="c))

            Return report.ToString()

        Catch ex As Exception
            Console.WriteLine($"Error generating mission report: {ex.Message}")
            Return $"เกิดข้อผิดพลาดในการสร้างรายงาน: {ex.Message}"
        End Try
    End Function

    ''' <summary>
    ''' ดึงรายการ Mission ทั้งหมดจากไฟล์
    ''' </summary>
    ''' <returns>รายการ Mission</returns>
    Private Function GetAllMissions() As List(Of MissionInfo)
        Dim missions As New List(Of MissionInfo)()

        Try
            Dim missionDir As String = Path.Combine(Application.StartupPath, "Missions")
            If Not Directory.Exists(missionDir) Then
                Return missions
            End If

            Dim missionFiles() As String = Directory.GetFiles(missionDir, "MISSION_*.txt")

            For Each filePath As String In missionFiles
                Try
                    Dim content As String = File.ReadAllText(filePath, Encoding.UTF8)
                    Dim mission As New MissionInfo()

                    ' แยกข้อมูลจากไฟล์
                    Dim lines() As String = content.Split({vbCrLf, vbLf}, StringSplitOptions.RemoveEmptyEntries)
                    For Each line As String In lines
                        If line.StartsWith("ID: ") Then
                            mission.Id = line.Substring(4).Trim()
                        ElseIf line.StartsWith("ชื่อ: ") Then
                            mission.Name = line.Substring(4).Trim()
                        ElseIf line.StartsWith("ผู้รับผิดชอบ: ") Then
                            mission.AssignedTo = line.Substring(13).Trim()
                        ElseIf line.StartsWith("วันที่สร้าง: ") Then
                            DateTime.TryParse(line.Substring(11).Trim(), mission.CreatedDate)
                        ElseIf line.StartsWith("รหัสผลิตภัณฑ์: ") Then
                            mission.ProductCode = line.Substring(15).Trim()
                        End If
                    Next

                    mission.FilePath = filePath
                    missions.Add(mission)

                Catch ex As Exception
                    Console.WriteLine($"Error reading mission file {filePath}: {ex.Message}")
                End Try
            Next

        Catch ex As Exception
            Console.WriteLine($"Error getting all missions: {ex.Message}")
        End Try

        Return missions
    End Function

    ''' <summary>
    ''' แสดงรายการ Mission ทั้งหมด
    ''' </summary>
    Private Sub ShowAllMissions()
        Try
            Dim missions As List(Of MissionInfo) = GetAllMissions()

            If missions.Count = 0 Then
                MessageBox.Show("ไม่มี Mission ในระบบ", "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' สร้างฟอร์มแสดงรายการ Mission
            Dim listForm As New Form()
            listForm.Text = "รายการ Mission ทั้งหมด"
            listForm.Size = New Size(800, 600)
            listForm.StartPosition = FormStartPosition.CenterParent

            ' สร้าง DataGridView
            Dim dgvMissions As New DataGridView()
            dgvMissions.Dock = DockStyle.Fill
            dgvMissions.AutoGenerateColumns = False
            dgvMissions.AllowUserToAddRows = False
            dgvMissions.AllowUserToDeleteRows = False
            dgvMissions.ReadOnly = True
            dgvMissions.SelectionMode = DataGridViewSelectionMode.FullRowSelect

            ' สร้างคอลัมน์
            dgvMissions.Columns.Add("Id", "Mission ID")
            dgvMissions.Columns.Add("Name", "ชื่อ Mission")
            dgvMissions.Columns.Add("ProductCode", "รหัสผลิตภัณฑ์")
            dgvMissions.Columns.Add("AssignedTo", "ผู้รับผิดชอบ")
            dgvMissions.Columns.Add("CreatedDate", "วันที่สร้าง")

            ' ปรับความกว้างคอลัมน์
            dgvMissions.Columns("Id").Width = 200
            dgvMissions.Columns("Name").Width = 200
            dgvMissions.Columns("ProductCode").Width = 150
            dgvMissions.Columns("AssignedTo").Width = 120
            dgvMissions.Columns("CreatedDate").Width = 150

            ' เพิ่มข้อมูล
            For Each mission As MissionInfo In missions
                dgvMissions.Rows.Add(mission.Id, mission.Name, mission.ProductCode,
                                   mission.AssignedTo, mission.CreatedDate.ToString("yyyy-MM-dd HH:mm"))
            Next

            listForm.Controls.Add(dgvMissions)
            listForm.ShowDialog()

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการแสดงรายการ Mission: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
        End Try
    End Sub

    ''' <summary>
    ''' ลบ Mission
    ''' </summary>
    ''' <param name="missionId">ID ของ Mission ที่ต้องการลบ</param>
    ''' <returns>True ถ้าลบสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Private Function DeleteMission(missionId As String) As Boolean
        Try
            If String.IsNullOrEmpty(missionId) Then
                MessageBox.Show("ไม่พบ Mission ID", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

            Dim missionDir As String = Path.Combine(Application.StartupPath, "Missions")
            Dim missionFile As String = Path.Combine(missionDir, $"{missionId}.txt")

            If File.Exists(missionFile) Then
                File.Delete(missionFile)
                MessageBox.Show($"ลบ Mission '{missionId}' สำเร็จ", "สำเร็จ", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return True
            Else
                MessageBox.Show($"ไม่พบไฟล์ Mission: {missionId}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
                Return False
            End If

        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการลบ Mission: {ex.Message}",
                           "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Console.WriteLine($"Error in DeleteMission: {ex.Message}")
            Return False
        End Try
    End Function

#End Region

#Region "Support Classes for Mission"

    ''' <summary>
    ''' คลาสสำหรับเก็บข้อมูล Mission
    ''' </summary>
    Public Class MissionInfo
        Public Property Id As String = ""
        Public Property Name As String = ""
        Public Property Description As String = ""
        Public Property ProductCode As String = ""
        Public Property AssignedTo As String = ""
        Public Property CreatedDate As DateTime = DateTime.MinValue
        Public Property CompletedDate As DateTime? = Nothing
        Public Property Status As String = "รอดำเนินการ"
        Public Property FilePath As String = ""

        Public Sub New()
            CreatedDate = DateTime.Now
        End Sub

        Public Overrides Function ToString() As String
            Return $"{Id} - {Name} ({Status})"
        End Function
    End Class



    ''' <summary>
    ''' คลาสสำหรับเก็บผลการตรวจสอบการสร้าง Mission
    ''' </summary>
    Public Class MissionCreationCheck
        Public Property CanCreate As Boolean = False
        Public Property Reason As String = ""
        Public Property ExcelMatch As ExcelUtility.ExcelMatchResult = Nothing
        Public Property FoundFile As FileDetail = Nothing
    End Class

#End Region

#Region "Support Classes"
    ' คลาสสำหรับผลลัพธ์การตรวจสอบเครือข่าย
    Public Class NetworkCheckResult
        Public Property IsConnected As Boolean = False
        Public Property NetworkType As String = ""
        Public Property ErrorMessage As String = ""
    End Class
#End Region

    ''' <summary>
    ''' อัปเดตชื่อโปรแกรมด้วยเวอร์ชันจาก Assembly
    ''' </summary>
    Private Sub UpdateFormTitleWithVersion()
        Try
            Dim version As Version = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Dim versionString As String = $"v{version.Major}.{version.Minor}.{version.Build}"

            ' อัปเดตชื่อในหัวข้อฟอร์ม
            Me.Text = $"ประวัติการสแกน QR Code - QR Code Scanner System {versionString}"

        Catch ex As Exception
            ' ถ้าอ่านเวอร์ชันไม่ได้ ให้ใช้ชื่อเดิม
            Me.Text = "ประวัติการสแกน QR Code - QR Code Scanner System"
            Console.WriteLine($"Error reading assembly version in frmHistory: {ex.Message}")
        End Try
    End Sub



#Region "Additional Helper Methods"

    ''' <summary>
    ''' รีเฟรชข้อมูล Excel เบื้องหลัง
    ''' </summary>
    Private Sub RefreshExcelDataAsync()
        Try
            Console.WriteLine("เริ่มรีเฟรชข้อมูล Excel...")

            Dim result = dataCache.RefreshData()

            Me.Invoke(Sub()
                          Try
                              If result.IsSuccess Then
                                  ShowExcelLoadingStatus($"รีเฟรชข้อมูลสำเร็จ: {dataCache.RowCount:N0} แถว")
                                  EnableExcelSearchControls(True)

                                  ' แสดงการแจ้งเตือนสั้นๆ
                                  ShowSuccessNotification($"รีเฟรชข้อมูล {dataCache.RowCount:N0} แถว สำเร็จ")
                              Else
                                  ShowExcelLoadingStatus($"รีเฟรชไม่สำเร็จ: {result.Message}")
                                  ' ไม่ปิด Controls เพราะยังใช้ข้อมูลเก่าได้
                              End If
                          Catch uiEx As Exception
                              Console.WriteLine($"Error updating UI after Excel refresh: {uiEx.Message}")
                          End Try
                      End Sub)

        Catch ex As Exception
            Console.WriteLine($"Error in RefreshExcelDataAsync: {ex.Message}")
            Me.Invoke(Sub()
                          ShowExcelLoadingStatus($"รีเฟรชข้อผิดพลาด: {ex.Message}")
                      End Sub)
        End Try
    End Sub

    ''' <summary>
    ''' สร้าง Mission ทั้งหมดที่ยังไม่มี
    ''' </summary>
    Private Async Sub CreateAllMissions()
        Try
            ' ค้นหารายการที่ยังไม่มี Mission
            Dim recordsWithoutMission As New List(Of ScanDataRecord)()

            For Each record As ScanDataRecord In filteredHistory
                If String.IsNullOrEmpty(record.MissionStatus) OrElse record.MissionStatus = "ไม่มี" Then
                    recordsWithoutMission.Add(record)
                End If
            Next

            If recordsWithoutMission.Count = 0 Then
                MessageBox.Show("ไม่มีรายการที่ต้องสร้าง Mission",
                           "แจ้งเตือน", MessageBoxButtons.OK, MessageBoxIcon.Information)
                Return
            End If

            ' ยืนยันการสร้าง
            Dim confirmResult = MessageBox.Show(
            $"พบรายการที่ยังไม่มี Mission จำนวน {recordsWithoutMission.Count} รายการ{vbCrLf}{vbCrLf}" &
            "ต้องการสร้าง Mission สำหรับรายการทั้งหมดหรือไม่?",
            "ยืนยันการสร้าง Mission",
            MessageBoxButtons.YesNo,
            MessageBoxIcon.Question)

            If confirmResult <> DialogResult.Yes Then
                Return
            End If

            ' แสดง Progress Bar
            toolStripProgressBar.Visible = True
            toolStripProgressBar.Style = ProgressBarStyle.Continuous
            toolStripProgressBar.Maximum = recordsWithoutMission.Count
            toolStripProgressBar.Value = 0

            Dim successCount As Integer = 0
            Dim failCount As Integer = 0
            Dim failedItems As New List(Of String)()

            For i As Integer = 0 To recordsWithoutMission.Count - 1
                Dim record = recordsWithoutMission(i)

                ' อัปเดต Status Bar
                lblCount.Text = $"กำลังสร้าง Mission... {i + 1}/{recordsWithoutMission.Count} - {record.ProductCode}"
                Application.DoEvents()

                Try
                    ' ตรวจสอบว่าสามารถสร้าง Mission ได้หรือไม่
                    Dim canCreateResult As MissionCreationCheck = Await CheckMissionCreationRequirementsAsync(record.ProductCode)

                    If canCreateResult.CanCreate Then
                        ' สร้าง Mission แบบอัตโนมัติ
                        If CreateAutoMission(record, canCreateResult) Then
                            record.MissionStatus = "รอดำเนินการ"
                            UpdateMissionStatus(record)
                            successCount += 1
                        Else
                            failCount += 1
                            failedItems.Add($"{record.ProductCode} - ไม่สามารถสร้าง Mission ได้")
                        End If
                    Else
                        failCount += 1
                        failedItems.Add($"{record.ProductCode} - {canCreateResult.Reason}")
                    End If

                Catch ex As Exception
                    failCount += 1
                    failedItems.Add($"{record.ProductCode} - ข้อผิดพลาด: {ex.Message}")
                End Try

                ' อัปเดต Progress Bar
                toolStripProgressBar.Value = i + 1
            Next

            ' ซ่อน Progress Bar
            toolStripProgressBar.Visible = False
            lblCount.Text = $"จำนวนรายการ: {filteredHistory.Count} จาก {scanHistory.Count} รายการ"

            ' รีเฟรชข้อมูล - แก้ไขตรงนี้
            LoadScanHistory()

            ' แสดงผลลัพธ์
            Dim resultMessage As String = $"สร้าง Mission เสร็จสิ้น!{vbCrLf}{vbCrLf}" &
                                    $"✅ สำเร็จ: {successCount} รายการ{vbCrLf}" &
                                    $"❌ ไม่สำเร็จ: {failCount} รายการ"

            If failCount > 0 Then
                resultMessage += $"{vbCrLf}{vbCrLf}รายการที่ไม่สำเร็จ:{vbCrLf}"
                For Each failedItem In failedItems.Take(10) ' แสดงเฉพาะ 10 รายการแรก
                    resultMessage += $"• {failedItem}{vbCrLf}"
                Next

                If failedItems.Count > 10 Then
                    resultMessage += $"• และอีก {failedItems.Count - 10} รายการ..."
                End If
            End If

            MessageBox.Show(resultMessage, "ผลการสร้าง Mission",
                       MessageBoxButtons.OK,
                       If(failCount = 0, MessageBoxIcon.Information, MessageBoxIcon.Warning))

        Catch ex As Exception
        toolStripProgressBar.Visible = False
        MessageBox.Show($"เกิดข้อผิดพลาดในการสร้าง Mission: {ex.Message}",
                       "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try
End Sub

''' <summary>
''' สร้าง Mission แบบอัตโนมัติ (ไม่ต้องกรอกข้อมูล)
''' </summary>
Private Function CreateAutoMission(record As ScanDataRecord, creationCheck As MissionCreationCheck) As Boolean
    Try
        ' สร้าง Mission ID ใหม่
        Dim missionId As String = $"MISSION_{DateTime.Now:yyyyMMddHHmmss}_{record.Id}"
        
        ' ข้อมูล Mission แบบอัตโนมัติ
        Dim missionName As String = $"Mission_{record.ProductCode}"
        Dim assignedTo As String = "Auto Generated"
        Dim dueDate As DateTime = DateTime.Now.AddDays(7) ' กำหนดส่ง 7 วัน
        Dim description As String = $"Mission ที่สร้างอัตโนมัติสำหรับ {record.ProductCode}"
        
        ' บันทึก RelatedFilePath ลงฐานข้อมูล
        If creationCheck.FoundFile IsNot Nothing AndAlso Not String.IsNullOrEmpty(creationCheck.FoundFile.FullPath) Then
            Dim relatedFilePath As String = creationCheck.FoundFile.FullPath
            Dim updateSuccess As Boolean = AccessDatabaseManager.UpdateRelatedFilePath(record.Id, relatedFilePath)
            If updateSuccess Then
                record.RelatedFilePath = relatedFilePath
            End If
        End If
        
        ' บันทึกข้อมูล Mission
        Dim missionData As String = $"ID: {missionId}{vbCrLf}" &
                                   $"ชื่อ: {missionName}{vbCrLf}" &
                                   $"รายละเอียด: {description}{vbCrLf}" &
                                   $"ผู้รับผิดชอบ: {assignedTo}{vbCrLf}" &
                                   $"กำหนดส่ง: {dueDate:yyyy-MM-dd HH:mm:ss}{vbCrLf}" &
                                   $"วันที่สร้าง: {DateTime.Now:yyyy-MM-dd HH:mm:ss}{vbCrLf}" &
                                   $"รหัสผลิตภัณฑ์: {record.ProductCode}{vbCrLf}" &
                                   $"ข้อมูล Excel: {creationCheck.ExcelMatch.Column4Value}{vbCrLf}" &
                                   $"ไฟล์เกี่ยวข้อง: {If(creationCheck.FoundFile IsNot Nothing, creationCheck.FoundFile.FullPath, "")}"
        
        ' บันทึกลงไฟล์
        Dim missionDir As String = Path.Combine(Application.StartupPath, "Missions")
        If Not Directory.Exists(missionDir) Then
            Directory.CreateDirectory(missionDir)
        End If
        
        Dim missionFile As String = Path.Combine(missionDir, $"{missionId}.txt")
        File.WriteAllText(missionFile, missionData, Encoding.UTF8)
        
        Return True
        
    Catch ex As Exception
        Console.WriteLine($"Error in CreateAutoMission: {ex.Message}")
        Return False
    End Try
End Function

#End Region

End Class