Imports System.Data
Imports System.Data.OleDb
Imports System.IO
Imports System.Xml
Imports System.Windows.Forms


''' <summary>
''' คลาสสำหรับจัดการฐานข้อมูล MS Access
''' </summary>
Public Class AccessDatabaseManager

    ' กำหนดพาธของไฟล์การตั้งค่า
    Private Shared ReadOnly CONFIG_FILE As String = "Settings.config"

    ' กำหนดพาธของฐานข้อมูล Access
    Private Shared _databasePath As String = "\\fls951\OAFAB\OA2FAB\Film charecter check\dbSystems\QRCodeScanner.accdb"
    Private Shared _password As String = "" ' รหัสผ่านฐานข้อมูล (ถ้ามี)

    ''' <summary>
    ''' Connection string สำหรับเชื่อมต่อฐานข้อมูล Access
    ''' </summary>
    Public Shared ReadOnly Property ConnectionString As String
        Get
            Dim connectionStrings As String = $"Provider=Microsoft.ACE.OLEDB.12.0;Data Source={_databasePath};"

            ' เพิ่มรหัสผ่านถ้ามี
            If Not String.IsNullOrEmpty(_password) Then
                connectionStrings += $"Jet OLEDB:Database Password={_password};"
            End If

            Return connectionStrings
        End Get
    End Property



    ''' <summary>
    ''' ตรวจสอบการเชื่อมต่อฐานข้อมูล
    ''' </summary>
    ''' <returns>True ถ้าเชื่อมต่อได้, False ถ้าเชื่อมต่อไม่ได้</returns>
    Public Shared Function IsConnected() As Boolean
        Try
            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()
                conn.Close()
                Return True
            End Using
        Catch ex As Exception
            Console.WriteLine($"Connection error: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ตั้งค่าพาธของฐานข้อมูล
    ''' </summary>
    ''' <param name="path">พาธของไฟล์ฐานข้อมูล</param>
    Public Shared Sub SetDatabasePath(path As String)
        If Not String.IsNullOrEmpty(path) Then
            _databasePath = path
            Console.WriteLine($"Database path set to: {_databasePath}")
        End If
    End Sub

    ''' <summary>
    ''' ตั้งค่ารหัสผ่านของฐานข้อมูล
    ''' </summary>
    ''' <param name="password">รหัสผ่านของฐานข้อมูล</param>
    Public Shared Sub SetDatabasePassword(password As String)
        _password = password
        Console.WriteLine("Database password updated")
    End Sub

    ''' <summary>
    ''' เริ่มต้นการใช้งานฐานข้อมูล
    ''' </summary>
    ''' <returns>True ถ้าเริ่มต้นสำเร็จ, False ถ้าเริ่มต้นไม่สำเร็จ</returns>
    Public Shared Function Initialize() As Boolean
        Try
            Console.WriteLine($"Initializing database with path: {_databasePath}")

            ' สร้างฐานข้อมูลใหม่ถ้ายังไม่มี
            If Not File.Exists(_databasePath) Then
                Console.WriteLine("Database file not found, creating new database...")
                CreateNewDatabase()
            End If

            ' สร้างตารางหากยังไม่มี
            CreateTablesIfNotExists()

            Return IsConnected()
        Catch ex As Exception
            Console.WriteLine($"Error initializing database: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' สร้างไฟล์ฐานข้อมูล Access ใหม่
    ''' </summary>
    Private Shared Sub CreateNewDatabase()
        Try
            ' สร้างโฟลเดอร์ถ้ายังไม่มี
            Dim directory As String = Path.GetDirectoryName(_databasePath)
            If Not IO.Directory.Exists(directory) Then
                IO.Directory.CreateDirectory(directory)
            End If

            ' สร้างฐานข้อมูล Access ใหม่ด้วย ADOX
            Try
                ' วิธีที่ 1: ใช้ ADOX (แนะนำ)
                Dim catalog As Object = CreateObject("ADOX.Catalog")
                catalog.Create(ConnectionString)
                catalog = Nothing
                Console.WriteLine($"Access database created successfully: {_databasePath}")
            Catch adoxEx As Exception
                Console.WriteLine($"ADOX method failed: {adoxEx.Message}")

                ' วิธีที่ 2: สร้างด้วย OleDb (backup method)
                CreateDatabaseWithOleDb()
            End Try

        Catch ex As Exception
            Console.WriteLine($"Error creating database: {ex.Message}")
            Throw New Exception($"ไม่สามารถสร้างฐานข้อมูล Access ได้: {ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' สร้างฐานข้อมูลด้วย OleDb (วิธีสำรอง)
    ''' </summary>
    Private Shared Sub CreateDatabaseWithOleDb()
        Try
            ' สร้างไฟล์ .accdb เปล่า
            File.WriteAllBytes(_databasePath, New Byte() {})

            ' ลองเชื่อมต่อเพื่อ initialize
            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()
                conn.Close()
            End Using

            Console.WriteLine($"Database created with OleDb method: {_databasePath}")
        Catch ex As Exception
            Console.WriteLine($"OleDb creation method also failed: {ex.Message}")
            Throw
        End Try 
    End Sub

    ''' <summary>
    ''' บันทึกข้อมูลการสแกน
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>ID ของรายการที่บันทึก หรือ 0 ถ้าบันทึกไม่สำเร็จ</returns>
    Public Shared Function SaveScanData(record As ScanDataRecord) As Integer
        Try
            Return AddScanRecord(record)
        Catch ex As Exception
            Console.WriteLine($"Error in SaveScanData: {ex.Message}")
            Return 0
        End Try
    End Function
    Public Shared Function AddScanRecord(record As ScanDataRecord) As Integer
        Try
            ' ตรวจสอบข้อมูลพื้นฐาน
            If record Is Nothing Then
                Throw New ArgumentException("ข้อมูล record เป็น null")
            End If

            ' ตรวจสอบว่าฐานข้อมูลพร้อมใช้งาน
            If Not IsConnected() Then
                Throw New Exception("ไม่สามารถเชื่อมต่อฐานข้อมูลได้")
            End If

            ' ตรวจสอบว่าตาราง ScanRecords มีอยู่หรือไม่
            Try
                Using testConn As New OleDbConnection(ConnectionString)
                    testConn.Open()

                    ' ตรวจสอบว่ามีตาราง ScanRecords หรือไม่
                    Dim tableExists As Boolean = False
                    Dim tables As DataTable = testConn.GetSchema("Tables")
                    For Each row As DataRow In tables.Rows
                        If row("TABLE_NAME").ToString().Equals("ScanRecords", StringComparison.OrdinalIgnoreCase) Then
                            tableExists = True
                            Exit For
                        End If
                    Next

                    If Not tableExists Then
                        Throw New Exception("ตาราง ScanRecords ไม่มีอยู่ในฐานข้อมูล - กรุณาเรียกใช้ CreateTablesIfNotExists() ก่อน")
                    End If

                    Console.WriteLine("Table ScanRecords exists and is accessible")
                    testConn.Close()
                End Using
            Catch ex As Exception
                Throw New Exception($"ไม่สามารถตรวจสอบโครงสร้างฐานข้อมูลได้: {ex.Message}", ex)
            End Try

            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                ' เริ่มต้นการทำธุรกรรม
                Using trans As OleDbTransaction = conn.BeginTransaction()
                    Try
                        ' แก้ไข SQL statement - ไม่ใส่ id เพราะเป็น AUTOINCREMENT
                        Dim insertSql As String =
                        "INSERT INTO ScanRecords " &
                        "(scandatetime, productcode, referencecode, quantity, datecode, isvalid, " &
                        "originaldata, extracteddata, validationmessages, computername, username, MissionStatus) " &
                        "VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)"

                        Using insertCmd As New OleDbCommand(insertSql, conn, trans)
                            ' เพิ่มพารามิเตอร์โดยไม่ใส่ id
                            insertCmd.Parameters.Add("@scandatetime", OleDbType.Date).Value = If(record.ScanDateTime = DateTime.MinValue, DateTime.Now, record.ScanDateTime)
                            insertCmd.Parameters.Add("@productcode", OleDbType.VarChar, 50).Value = If(String.IsNullOrEmpty(record.ProductCode), DBNull.Value, record.ProductCode)
                            insertCmd.Parameters.Add("@referencecode", OleDbType.VarChar, 50).Value = If(String.IsNullOrEmpty(record.ReferenceCode), DBNull.Value, record.ReferenceCode)
                            insertCmd.Parameters.Add("@quantity", OleDbType.Integer).Value = record.Quantity
                            insertCmd.Parameters.Add("@datecode", OleDbType.VarChar, 20).Value = If(String.IsNullOrEmpty(record.DateCode), DBNull.Value, record.DateCode)
                            insertCmd.Parameters.Add("@isvalid", OleDbType.Boolean).Value = record.IsValid
                            insertCmd.Parameters.Add("@originaldata", OleDbType.LongVarChar).Value = If(String.IsNullOrEmpty(record.OriginalData), DBNull.Value, record.OriginalData)
                            insertCmd.Parameters.Add("@extracteddata", OleDbType.LongVarChar).Value = If(String.IsNullOrEmpty(record.ExtractedData), DBNull.Value, record.ExtractedData)
                            insertCmd.Parameters.Add("@validationmessages", OleDbType.LongVarChar).Value = If(String.IsNullOrEmpty(record.ValidationMessages), DBNull.Value, record.ValidationMessages)
                            insertCmd.Parameters.Add("@computername", OleDbType.VarChar, 50).Value = If(String.IsNullOrEmpty(record.ComputerName), Environment.MachineName, record.ComputerName)
                            insertCmd.Parameters.Add("@username", OleDbType.VarChar, 50).Value = If(String.IsNullOrEmpty(record.UserName), Environment.UserName, record.UserName)
                            insertCmd.Parameters.Add("@MissionStatus", OleDbType.VarChar, 50).Value = If(String.IsNullOrEmpty(record.MissionStatus), "ไม่มี", record.MissionStatus)

                            Console.WriteLine($"Executing INSERT command for ProductCode: {record.ProductCode}")
                            Console.WriteLine($"SQL: {insertSql}")
                            Console.WriteLine($"Parameters count: {insertCmd.Parameters.Count}")

                            ' แสดงค่าพารามิเตอร์ทั้งหมด
                            For i As Integer = 0 To insertCmd.Parameters.Count - 1
                                Console.WriteLine($"Parameter {i}: {insertCmd.Parameters(i).Value}")
                            Next

                            Dim rowsAffected As Integer = insertCmd.ExecuteNonQuery()

                            If rowsAffected = 0 Then
                                trans.Rollback()
                                Throw New Exception("ไม่มีแถวใดได้รับการเพิ่ม - INSERT ไม่สำเร็จ")
                            End If

                            Console.WriteLine($"INSERT successful, rows affected: {rowsAffected}")
                        End Using

                        ' ดึง ID ที่เพิ่มล่าสุด
                        Using idCmd As New OleDbCommand("SELECT @@IDENTITY", conn, trans)
                            Dim result = idCmd.ExecuteScalar()
                            If result Is Nothing OrElse IsDBNull(result) Then
                                trans.Rollback()
                                Throw New Exception("ไม่สามารถดึง ID ของข้อมูลที่เพิ่มได้")
                            End If

                            Dim id As Integer = Convert.ToInt32(result)
                            record.Id = id
                            trans.Commit()

                            Console.WriteLine($"Added scan record with ID: {id}")
                            Return id
                        End Using
                    Catch
                        trans.Rollback()
                        Throw
                    End Try
                End Using
            End Using

        Catch ex As OleDbException
            Console.WriteLine($"Database error adding scan record: {ex.Message}")
            Console.WriteLine($"Error Code: {ex.ErrorCode}")
            Console.WriteLine($"SQL State: {ex.Errors(0).SQLState}")
            Throw New Exception($"ข้อผิดพลาดฐานข้อมูล: {ex.Message}", ex)
        Catch ex As Exception
            Console.WriteLine($"General error adding scan record: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")
            Throw New Exception($"ไม่สามารถเพิ่มข้อมูลการสแกนได้: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' ลบข้อมูลการสแกนตาม ID
    ''' </summary>
    ''' <param name="id">ID ของรายการที่ต้องการลบ</param>
    ''' <returns>True ถ้าลบสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Public Shared Function DeleteScanRecord(id As Integer) As Boolean
        Try
            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                Using deleteCmd As New OleDbCommand("DELETE FROM ScanRecords WHERE Id = ?", conn)
                    deleteCmd.Parameters.AddWithValue("@Id", id)
                    Dim rowsAffected As Integer = deleteCmd.ExecuteNonQuery()

                    Console.WriteLine($"Deleted scan record ID {id}, rows affected: {rowsAffected}")
                    Return rowsAffected > 0
                End Using
            End Using

        Catch ex As Exception
            Console.WriteLine($"Error deleting scan record: {ex.Message}")
            Throw New Exception($"ไม่สามารถลบข้อมูลการสแกนได้: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' อัปเดตข้อมูลการสแกน
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกนที่อัปเดต</param>
    ''' <returns>True ถ้าอัปเดตสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Public Shared Function UpdateScanRecord(record As ScanDataRecord) As Boolean
        Try
            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                Dim updateSql As String =
                    "UPDATE ScanRecords SET " &
                    "ProductCode = ?, " &
                    "ReferenceCode = ?, " &
                    "Quantity = ?, " &
                    "DateCode = ?, " &
                    "IsValid = ?, " &
                    "ExtractedData = ?, " &
                    "ValidationMessages = ? " &
                    "WHERE Id = ?"

                Using updateCmd As New OleDbCommand(updateSql, conn)
                    ' เพิ่มพารามิเตอร์
                    updateCmd.Parameters.AddWithValue("@ProductCode", If(record.ProductCode, DBNull.Value))
                    updateCmd.Parameters.AddWithValue("@ReferenceCode", If(record.ReferenceCode, DBNull.Value))
                    updateCmd.Parameters.AddWithValue("@Quantity", record.Quantity)
                    updateCmd.Parameters.AddWithValue("@DateCode", If(record.DateCode, DBNull.Value))
                    updateCmd.Parameters.AddWithValue("@IsValid", record.IsValid)
                    updateCmd.Parameters.AddWithValue("@ExtractedData", If(record.ExtractedData, DBNull.Value))
                    updateCmd.Parameters.AddWithValue("@ValidationMessages", If(record.ValidationMessages, DBNull.Value))
                    updateCmd.Parameters.AddWithValue("@Id", record.Id)

                    Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()

                    Console.WriteLine($"Updated scan record ID {record.Id}, rows affected: {rowsAffected}")
                    Return rowsAffected > 0
                End Using

            End Using

        Catch ex As Exception
            Console.WriteLine($"Error updating scan record: {ex.Message}")
            Throw New Exception($"ไม่สามารถอัปเดตข้อมูลการสแกนได้: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' ค้นหาข้อมูลการสแกนตามรหัสผลิตภัณฑ์
    ''' </summary>
    ''' <param name="productCode">รหัสผลิตภัณฑ์</param>
    ''' <returns>รายการข้อมูลการสแกนที่ตรงกับรหัสผลิตภัณฑ์</returns>
    Public Shared Function SearchByProductCode(productCode As String) As List(Of ScanDataRecord)
        Try
            Dim results As New List(Of ScanDataRecord)()

            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                ' Access ใช้ * แทน % สำหรับ wildcard
                Using selectCmd As New OleDbCommand("SELECT * FROM ScanRecords WHERE ProductCode LIKE ? ORDER BY ScanDateTime DESC", conn)
                    selectCmd.Parameters.AddWithValue("@ProductCode", $"*{productCode}*")

                    Using reader As OleDbDataReader = selectCmd.ExecuteReader()
                        While reader.Read()
                            Dim record As New ScanDataRecord()

                            record.Id = Convert.ToInt32(reader("Id"))
                            record.ScanDateTime = Convert.ToDateTime(reader("ScanDateTime"))
                            record.ProductCode = If(reader("ProductCode") IsNot DBNull.Value, reader("ProductCode").ToString(), "")
                            record.ReferenceCode = If(reader("ReferenceCode") IsNot DBNull.Value, reader("ReferenceCode").ToString(), "")
                            record.Quantity = If(reader("Quantity") IsNot DBNull.Value, Convert.ToInt32(reader("Quantity")), 0)
                            record.DateCode = If(reader("DateCode") IsNot DBNull.Value, reader("DateCode").ToString(), "")
                            record.IsValid = Convert.ToBoolean(reader("IsValid"))
                            record.OriginalData = If(reader("OriginalData") IsNot DBNull.Value, reader("OriginalData").ToString(), "")
                            record.ExtractedData = If(reader("ExtractedData") IsNot DBNull.Value, reader("ExtractedData").ToString(), "")
                            record.ValidationMessages = If(reader("ValidationMessages") IsNot DBNull.Value, reader("ValidationMessages").ToString(), "")
                            record.ComputerName = If(reader("ComputerName") IsNot DBNull.Value, reader("ComputerName").ToString(), "")
                            record.UserName = If(reader("UserName") IsNot DBNull.Value, reader("UserName").ToString(), "")

                            results.Add(record)
                        End While
                    End Using
                End Using
            End Using

            Console.WriteLine($"Found {results.Count} records for product code: {productCode}")
            Return results

        Catch ex As Exception
            Console.WriteLine($"Error searching by product code: {ex.Message}")
            Throw New Exception($"ไม่สามารถค้นหาข้อมูลตามรหัสผลิตภัณฑ์ได้: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' ได้รับข้อมูลสถิติการใช้งาน
    ''' </summary>
    ''' <returns>ข้อมูลสถิติ</returns>
    Public Shared Function GetStatistics() As DatabaseStatistics
        Try
            Dim stats As New DatabaseStatistics()

            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                ' นับจำนวนรายการทั้งหมด
                Using countCmd As New OleDbCommand("SELECT COUNT(*) FROM ScanRecords", conn)
                    stats.TotalRecords = Convert.ToInt32(countCmd.ExecuteScalar())
                End Using

                ' นับจำนวนรายการที่ถูกต้อง
                Using validCmd As New OleDbCommand("SELECT COUNT(*) FROM ScanRecords WHERE IsValid = True", conn)
                    stats.ValidRecords = Convert.ToInt32(validCmd.ExecuteScalar())
                End Using

                ' นับจำนวนรายการที่ไม่ถูกต้อง
                stats.InvalidRecords = stats.TotalRecords - stats.ValidRecords

                ' หาวันที่สแกนล่าสุด
                Using lastCmd As New OleDbCommand("SELECT MAX(ScanDateTime) FROM ScanRecords", conn)
                    Dim lastScan = lastCmd.ExecuteScalar()
                    If lastScan IsNot DBNull.Value Then
                        stats.LastScanDate = Convert.ToDateTime(lastScan)
                    End If
                End Using

            End Using

            Return stats

        Catch ex As Exception
            Console.WriteLine($"Error getting statistics: {ex.Message}")
            Return New DatabaseStatistics()
        End Try
    End Function

    ''' <summary>
    ''' สำรองข้อมูลฐานข้อมูล
    ''' </summary>
    ''' <param name="backupPath">พาธที่จะสำรองข้อมูล</param>
    ''' <returns>True ถ้าสำรองสำเร็จ</returns>
    Public Shared Function BackupDatabase(backupPath As String) As Boolean
        Try
            If File.Exists(_databasePath) Then
                ' สร้างโฟลเดอร์ backup ถ้ายังไม่มี
                Dim backupDir As String = Path.GetDirectoryName(backupPath)
                If Not Directory.Exists(backupDir) Then
                    Directory.CreateDirectory(backupDir)
                End If

                ' คัดลอกไฟล์ฐานข้อมูล
                File.Copy(_databasePath, backupPath, True)

                Console.WriteLine($"Database backed up to: {backupPath}")
                Return True
            End If

            Return False

        Catch ex As Exception
            Console.WriteLine($"Error backing up database: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' คืนค่าฐานข้อมูลจากไฟล์สำรอง
    ''' </summary>
    ''' <param name="backupPath">พาธของไฟล์สำรอง</param>
    ''' <returns>True ถ้าคืนค่าสำเร็จ</returns>
    Public Shared Function RestoreDatabase(backupPath As String) As Boolean
        Try
            If File.Exists(backupPath) Then
                ' สำรองไฟล์ปัจจุบันก่อน (ถ้ามี)
                If File.Exists(_databasePath) Then
                    Dim currentBackup As String = _databasePath + ".old"
                    File.Copy(_databasePath, currentBackup, True)
                End If

                ' คัดลอกไฟล์สำรองมาแทนที่
                File.Copy(backupPath, _databasePath, True)

                Console.WriteLine($"Database restored from: {backupPath}")
                Return True
            End If

            Return False

        Catch ex As Exception
            Console.WriteLine($"Error restoring database: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ทำความสะอาดข้อมูลเก่า
    ''' </summary>
    ''' <param name="daysOld">จำนวนวันที่จะลบข้อมูลเก่า</param>
    ''' <returns>จำนวนรายการที่ลบ</returns>
    Public Shared Function CleanupOldData(daysOld As Integer) As Integer
        Try
            Dim cutoffDate As DateTime = DateTime.Now.AddDays(-daysOld)

            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                Using deleteCmd As New OleDbCommand("DELETE FROM ScanRecords WHERE ScanDateTime < ?", conn)
                    deleteCmd.Parameters.AddWithValue("@CutoffDate", cutoffDate)
                    Dim deletedCount As Integer = deleteCmd.ExecuteNonQuery()

                    Console.WriteLine($"Cleaned up {deletedCount} old records older than {daysOld} days")
                    Return deletedCount
                End Using

            End Using

        Catch ex As Exception
            Console.WriteLine($"Error cleaning up old data: {ex.Message}")
            Return 0
        End Try
    End Function

    ''' <summary>
    ''' อัปเดตสถานะ Mission ของรายการสแกน
    ''' </summary>
    ''' <param name="recordId">ID ของรายการที่ต้องการอัปเดต</param>
    ''' <param name="missionStatus">สถานะ Mission ใหม่</param>
    ''' <returns>True ถ้าอัปเดตสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Public Shared Function UpdateMissionStatus(recordId As Integer, missionStatus As String) As Boolean
        Try
            Using conn As New OleDbConnection(ConnectionString)
                conn.Open()

                ' ตรวจสอบว่ามีคอลัมน์ MissionStatus หรือไม่
                Dim columnExists As Boolean = False
                Try
                    Dim schema As DataTable = conn.GetSchema("Columns", New String() {Nothing, Nothing, "ScanRecords", "MissionStatus"})
                    columnExists = schema.Rows.Count > 0
                Catch ex As Exception
                    Console.WriteLine($"Error checking MissionStatus column: {ex.Message}")
                    columnExists = False
                End Try

                ' ถ้าไม่มีคอลัมน์ MissionStatus ให้เพิ่มคอลัมน์
                If Not columnExists Then
                    Try
                        Using alterCmd As New OleDbCommand("ALTER TABLE ScanRecords ADD COLUMN MissionStatus TEXT(50)", conn)
                            alterCmd.ExecuteNonQuery()
                            Console.WriteLine("Added MissionStatus column to ScanRecords table")
                        End Using
                    Catch ex As Exception
                        Console.WriteLine($"Error adding MissionStatus column: {ex.Message}")
                        ' ถ้าไม่สามารถเพิ่มคอลัมน์ได้ ให้ข้ามไปและพยายามอัปเดตต่อ
                    End Try
                End If

                ' อัปเดตสถานะ Mission
                Using updateCmd As New OleDbCommand("UPDATE ScanRecords SET MissionStatus = ? WHERE Id = ?", conn)
                    updateCmd.Parameters.AddWithValue("@MissionStatus", missionStatus)
                    updateCmd.Parameters.AddWithValue("@Id", recordId)

                    Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()

                    Console.WriteLine($"Updated mission status for record ID {recordId} to {missionStatus}, rows affected: {rowsAffected}")
                    Return rowsAffected > 0
                End Using

                conn.Close()
            End Using

        Catch ex As Exception
            Console.WriteLine($"Error updating mission status: {ex.Message}")
            Return False
        End Try
    End Function

        ' ใน AccessDatabaseManager.vb - เพิ่มฟังก์ชั่นสำหรับจัดการคอลัมน์ RelatedFilePath

''' <summary>
''' เพิ่มคอลัมน์ RelatedFilePath ในตาราง ScanRecords ถ้ายังไม่มี
''' </summary>
Public Shared Sub AddRelatedFilePathColumnIfNotExists()
    Try
        Using conn As New OleDbConnection(ConnectionString)
            conn.Open()
            
            ' ตรวจสอบว่ามีคอลัมน์ RelatedFilePath หรือไม่
            Dim columnExists As Boolean = False
            Try
                Dim schema As DataTable = conn.GetSchema("Columns", New String() {Nothing, Nothing, "ScanRecords", "RelatedFilePath"})
                columnExists = schema.Rows.Count > 0
            Catch ex As Exception
                Console.WriteLine($"Error checking RelatedFilePath column: {ex.Message}")
                columnExists = False
            End Try
            
            ' ถ้าไม่มีคอลัมน์ RelatedFilePath ให้เพิ่มคอลัมน์
            If Not columnExists Then
                Try
                    Using alterCmd As New OleDbCommand("ALTER TABLE ScanRecords ADD COLUMN RelatedFilePath MEMO", conn)
                        alterCmd.ExecuteNonQuery()
                        Console.WriteLine("Added RelatedFilePath column to ScanRecords table")
                    End Using
                Catch ex As Exception
                    Console.WriteLine($"Error adding RelatedFilePath column: {ex.Message}")
                    Throw
                End Try
            Else
                Console.WriteLine("RelatedFilePath column already exists")
            End If
            
        End Using
    Catch ex As Exception
        Console.WriteLine($"Error in AddRelatedFilePathColumnIfNotExists: {ex.Message}")
        Throw New Exception($"ไม่สามารถเพิ่มคอลัมน์ RelatedFilePath ได้: {ex.Message}", ex)
    End Try
End Sub

''' <summary>
''' อัปเดต RelatedFilePath ของรายการสแกน
''' </summary>
''' <param name="recordId">ID ของรายการที่ต้องการอัปเดต</param>
''' <param name="filePath">เส้นทางไฟล์ที่เกี่ยวข้อง</param>
''' <returns>True ถ้าอัปเดตสำเร็จ, False ถ้าไม่สำเร็จ</returns>
Public Shared Function UpdateRelatedFilePath(recordId As Integer, filePath As String) As Boolean
    Try
        Using conn As New OleDbConnection(ConnectionString)
            conn.Open()
            
            ' ตรวจสอบและเพิ่มคอลัมน์ถ้าจำเป็น
            AddRelatedFilePathColumnIfNotExists()
            
            ' อัปเดต RelatedFilePath
            Using updateCmd As New OleDbCommand("UPDATE ScanRecords SET RelatedFilePath = ? WHERE Id = ?", conn)
                updateCmd.Parameters.AddWithValue("@FilePath", If(String.IsNullOrEmpty(filePath), DBNull.Value, filePath))
                updateCmd.Parameters.AddWithValue("@Id", recordId)
                
                Dim rowsAffected As Integer = updateCmd.ExecuteNonQuery()
                
                If rowsAffected > 0 Then
                    Console.WriteLine($"Updated RelatedFilePath for record ID {recordId}: {filePath}")
                    Return True
                Else
                    Console.WriteLine($"No record found with ID {recordId} to update RelatedFilePath")
                    Return False
                End If
            End Using
        End Using
        
    Catch ex As Exception
        Console.WriteLine($"Error updating RelatedFilePath: {ex.Message}")
        Return False
    End Try
End Function

''' <summary>
''' อัปเดต CreateTablesIfNotExists เพื่อรวมคอลัมน์ RelatedFilePath ในการสร้างตารางใหม่
''' </summary>
Public Shared Sub CreateTablesIfNotExists()
    Try
        Using conn As New OleDbConnection(ConnectionString)
            conn.Open()

            ' ตรวจสอบว่ามีตาราง ScanRecords หรือไม่
            Dim tableExists As Boolean = False

            Try
                Dim tables As DataTable = conn.GetSchema("Tables")
                For Each row As DataRow In tables.Rows
                    If row("TABLE_NAME").ToString().Equals("ScanRecords", StringComparison.OrdinalIgnoreCase) Then
                        tableExists = True
                        Exit For
                    End If
                Next
            Catch ex As Exception
                Console.WriteLine($"Error checking table existence: {ex.Message}")
                tableExists = False
            End Try

            ' สร้างตาราง ScanRecords ถ้ายังไม่มี (รวมคอลัมน์ RelatedFilePath)
            If Not tableExists Then
                Dim createTableSql As String =
                    "CREATE TABLE ScanRecords (" &
                    "Id AUTOINCREMENT PRIMARY KEY, " &
                    "ScanDateTime DATETIME NOT NULL, " &
                    "ProductCode TEXT(255), " &
                    "ReferenceCode TEXT(255), " &
                    "Quantity INTEGER, " &
                    "DateCode TEXT(255), " &
                    "IsValid YESNO, " &
                    "OriginalData MEMO, " &
                    "ExtractedData MEMO, " &
                    "ValidationMessages MEMO, " &
                    "ComputerName TEXT(255), " &
                    "UserName TEXT(255), " &
                    "MissionStatus TEXT(50) DEFAULT 'ไม่มี', " &
                    "RelatedFilePath MEMO" &
                    ")"

                Using createCmd As New OleDbCommand(createTableSql, conn)
                    createCmd.ExecuteNonQuery()
                    Console.WriteLine("ScanRecords table created successfully with RelatedFilePath column")
                End Using
            Else
                ' ถ้าตารางมีอยู่แล้ว ให้ตรวจสอบและเพิ่มคอลัมน์ที่ขาดหาย
                AddRelatedFilePathColumnIfNotExists()
            End If

            conn.Close()
        End Using

        Console.WriteLine("Tables created or updated successfully")
    Catch ex As Exception
        Console.WriteLine($"Error creating/updating tables: {ex.Message}")
        Throw New Exception($"ไม่สามารถสร้างหรืออัปเดตตารางในฐานข้อมูลได้: {ex.Message}", ex)
    End Try
End Sub

''' <summary>
''' อัปเดต GetScanHistory เพื่อรวมคอลัมน์ RelatedFilePath
''' </summary>
Public Shared Function GetScanHistory(Optional limit As Integer = 1000) As List(Of ScanDataRecord)
    Try
        Dim results As New List(Of ScanDataRecord)()

        Using conn As New OleDbConnection(ConnectionString)
            conn.Open()

            ' ตรวจสอบและเพิ่มคอลัมน์ที่ขาดหาย
            AddRelatedFilePathColumnIfNotExists()

            ' Access ใช้ TOP แทน LIMIT
            Dim selectSql As String = If(limit > 0,
                $"SELECT TOP {limit} * FROM ScanRecords ORDER BY ScanDateTime DESC",
                "SELECT * FROM ScanRecords ORDER BY ScanDateTime DESC")

            Using selectCmd As New OleDbCommand(selectSql, conn)
                Using reader As OleDbDataReader = selectCmd.ExecuteReader()
                    While reader.Read()
                        Dim record As New ScanDataRecord()

                        record.Id = Convert.ToInt32(reader("Id"))
                        record.ScanDateTime = Convert.ToDateTime(reader("ScanDateTime"))
                        record.ProductCode = If(reader("ProductCode") IsNot DBNull.Value, reader("ProductCode").ToString(), "")
                        record.ReferenceCode = If(reader("ReferenceCode") IsNot DBNull.Value, reader("ReferenceCode").ToString(), "")
                        record.Quantity = If(reader("Quantity") IsNot DBNull.Value, Convert.ToInt32(reader("Quantity")), 0)
                        record.DateCode = If(reader("DateCode") IsNot DBNull.Value, reader("DateCode").ToString(), "")
                        record.IsValid = Convert.ToBoolean(reader("IsValid"))
                        record.OriginalData = If(reader("OriginalData") IsNot DBNull.Value, reader("OriginalData").ToString(), "")
                        record.ExtractedData = If(reader("ExtractedData") IsNot DBNull.Value, reader("ExtractedData").ToString(), "")
                        record.ValidationMessages = If(reader("ValidationMessages") IsNot DBNull.Value, reader("ValidationMessages").ToString(), "")
                        record.ComputerName = If(reader("ComputerName") IsNot DBNull.Value, reader("ComputerName").ToString(), "")
                        record.UserName = If(reader("UserName") IsNot DBNull.Value, reader("UserName").ToString(), "")
                        
                        ' MissionStatus - ตรวจสอบว่ามีคอลัมน์หรือไม่
                        Try
                            record.MissionStatus = If(reader("MissionStatus") IsNot DBNull.Value, reader("MissionStatus").ToString(), "ไม่มี")
                        Catch
                            record.MissionStatus = "ไม่มี"
                        End Try
                        
                        ' RelatedFilePath - ตรวจสอบว่ามีคอลัมน์หรือไม่
                        Try
                            record.RelatedFilePath = If(reader("RelatedFilePath") IsNot DBNull.Value, reader("RelatedFilePath").ToString(), "")
                        Catch
                            record.RelatedFilePath = ""
                        End Try

                        results.Add(record)
                    End While
                End Using
            End Using
        End Using

        Console.WriteLine($"Retrieved {results.Count} scan records from database")
        Return results

    Catch ex As Exception
        Console.WriteLine($"Error retrieving scan history: {ex.Message}")
        Return New List(Of ScanDataRecord)()
    End Try
End Function

End Class

''' <summary>
''' คลาสสำหรับเก็บสถิติการใช้งานฐานข้อมูล
''' </summary>
Public Class DatabaseStatistics
    Public Property TotalRecords As Integer = 0
    Public Property ValidRecords As Integer = 0
    Public Property InvalidRecords As Integer = 0
    Public Property LastScanDate As DateTime = DateTime.MinValue

    Public ReadOnly Property ValidPercentage As Double
        Get
            If TotalRecords = 0 Then Return 0
            Return (ValidRecords / TotalRecords) * 100
        End Get
    End Property

    Public ReadOnly Property InvalidPercentage As Double
        Get
            If TotalRecords = 0 Then Return 0
            Return (InvalidRecords / TotalRecords) * 100
        End Get
    End Property
End Class