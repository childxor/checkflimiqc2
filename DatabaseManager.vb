Imports System.Data
Imports System.Data.SqlClient
Imports System.IO
Imports System.Xml

''' <summary>
''' คลาสสำหรับจัดการฐานข้อมูล
''' </summary>
Public Class DatabaseManager

    ' กำหนดพาธของไฟล์การตั้งค่า
    Private Shared ReadOnly CONFIG_FILE As String = "Settings.config"

    ' ตัวแปรสำหรับการเชื่อมต่อฐานข้อมูล
    Private Shared _server As String = "localhost"
    Private Shared _database As String = "ScanData"
    Private Shared _username As String = ""
    Private Shared _password As String = ""
    Private Shared _integratedSecurity As Boolean = True

    ' Connection string สำหรับเชื่อมต่อฐานข้อมูล
    Private Shared ReadOnly Property ConnectionString As String
        Get
            Dim builder As New SqlConnectionStringBuilder()
            builder.DataSource = _server
            builder.InitialCatalog = _database

            If _integratedSecurity Then
                builder.IntegratedSecurity = True
            Else
                builder.UserID = _username
                builder.Password = _password
            End If

            builder.ConnectTimeout = 30
            Return builder.ConnectionString
        End Get
    End Property

    ''' <summary>
    ''' โหลดการตั้งค่าฐานข้อมูลจากไฟล์ config
    ''' </summary>
    Private Shared Sub LoadDatabaseSettings()
        Try
            If File.Exists(CONFIG_FILE) Then
                Dim doc As New XmlDocument()
                doc.Load(CONFIG_FILE)

                _server = GetSettingValueFromXML(doc, "Server", "localhost")
                _database = GetSettingValueFromXML(doc, "Database", "ScanData")
                _username = GetSettingValueFromXML(doc, "Username", "")
                _password = GetSettingValueFromXML(doc, "Password", "")
                _integratedSecurity = GetSettingValueFromXML(doc, "IntegratedSecurity", True)

                Console.WriteLine($"Database settings loaded - Server: {_server}, Database: {_database}")
            End If
        Catch ex As Exception
            Console.WriteLine($"Error loading database settings: {ex.Message}")
        End Try
    End Sub

    ''' <summary>
    ''' ดึงค่าจาก XML
    ''' </summary>
    Private Shared Function GetSettingValueFromXML(doc As XmlDocument, key As String, defaultValue As Object) As Object
        Try
            Dim node As XmlNode = doc.SelectSingleNode($"//Setting[@key='{key}']")
            If node IsNot Nothing Then
                Dim value As String = node.Attributes("value").Value

                Select Case defaultValue.GetType()
                    Case GetType(Boolean)
                        Return Boolean.Parse(value)
                    Case GetType(Integer)
                        Return Integer.Parse(value)
                    Case GetType(Decimal)
                        Return Decimal.Parse(value)
                    Case Else
                        If key = "Password" Then
                            Return DecryptPassword(value)
                        End If
                        Return value
                End Select
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    ''' <summary>
    ''' ถอดรหัสพาสเวิร์ดแบบง่าย
    ''' </summary>
    Private Shared Function DecryptPassword(encryptedPassword As String) As String
        Try
            If String.IsNullOrEmpty(encryptedPassword) Then Return ""

            ' การถอดรหัสแบบง่าย (จาก Base64)
            Dim bytes As Byte() = Convert.FromBase64String(encryptedPassword)
            Return System.Text.Encoding.UTF8.GetString(bytes)
        Catch
            Return encryptedPassword
        End Try
    End Function

    ''' <summary>
    ''' ตรวจสอบการเชื่อมต่อฐานข้อมูล
    ''' </summary>
    ''' <returns>True ถ้าเชื่อมต่อได้, False ถ้าเชื่อมต่อไม่ได้</returns>
    Public Shared Function IsConnected() As Boolean
        Try
            Using conn As New SqlConnection(ConnectionString)
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
    ''' เริ่มต้นการใช้งานฐานข้อมูล
    ''' </summary>
    ''' <returns>True ถ้าเริ่มต้นสำเร็จ, False ถ้าเริ่มต้นไม่สำเร็จ</returns>
    Public Shared Function Initialize() As Boolean
        Try
            ' โหลดการตั้งค่าฐานข้อมูล
            LoadDatabaseSettings()

            ' ตรวจสอบการเชื่อมต่อและสร้างฐานข้อมูล
            InitializeDatabase()
            CreateTablesIfNotExists()
            Return IsConnected()
        Catch ex As Exception
            Console.WriteLine($"Error initializing database: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' สร้างตารางในฐานข้อมูลถ้ายังไม่มี
    ''' </summary>
    Public Shared Sub CreateTablesIfNotExists()
        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                ' ตรวจสอบว่ามีตาราง ScanRecords หรือไม่
                Dim tableExists As Boolean = False
                Dim checkTableCmd As New SqlCommand(
                    "SELECT COUNT(*) FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = 'ScanRecords'", conn)
                tableExists = (Convert.ToInt32(checkTableCmd.ExecuteScalar()) > 0)

                ' สร้างตาราง ScanRecords ถ้ายังไม่มี
                If Not tableExists Then
                    Dim createTableCmd As New SqlCommand(
                        "CREATE TABLE ScanRecords (" &
                        "Id INT IDENTITY(1,1) PRIMARY KEY, " &
                        "ScanDateTime DATETIME NOT NULL, " &
                        "ProductCode NVARCHAR(255), " &
                        "ReferenceCode NVARCHAR(255), " &
                        "Quantity INT, " &
                        "DateCode NVARCHAR(255), " &
                        "IsValid BIT, " &
                        "OriginalData NVARCHAR(MAX), " &
                        "ExtractedData NVARCHAR(MAX), " &
                        "ValidationMessages NVARCHAR(MAX), " &
                        "ComputerName NVARCHAR(255), " &
                        "UserName NVARCHAR(255)" &
                        ")", conn)

                    createTableCmd.ExecuteNonQuery()
                    Console.WriteLine("ScanRecords table created successfully")
                End If

                conn.Close()
            End Using

            Console.WriteLine("Tables created or already exist")
        Catch ex As Exception
            Console.WriteLine($"Error creating tables: {ex.Message}")
            Throw New Exception($"ไม่สามารถสร้างตารางในฐานข้อมูลได้: {ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' บันทึกข้อมูลการสแกน
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>ID ของรายการที่บันทึก</returns>
    Public Shared Function SaveScanData(record As ScanDataRecord) As Integer
        Return AddScanRecord(record)
    End Function

    ''' <summary>
    ''' สร้างฐานข้อมูลถ้ายังไม่มี
    ''' </summary> 
    Public Shared Sub InitializeDatabase()
        Try
            ' ตรวจสอบว่าสามารถเชื่อมต่อกับเซิร์ฟเวอร์ได้หรือไม่
            Using masterConn As New SqlConnection($"Data Source={_server};Initial Catalog=master;{If(_integratedSecurity, "Integrated Security=True", $"User ID={_username};Password={_password}")}")
                masterConn.Open()

                ' ตรวจสอบว่ามีฐานข้อมูลหรือไม่
                Dim checkDatabaseCmd As New SqlCommand($"SELECT COUNT(*) FROM sys.databases WHERE name = '{_database}'", masterConn)
                Dim databaseExists As Boolean = (Convert.ToInt32(checkDatabaseCmd.ExecuteScalar()) > 0)

                ' สร้างฐานข้อมูลถ้ายังไม่มี
                If Not databaseExists Then
                    Dim createDatabaseCmd As New SqlCommand($"CREATE DATABASE [{_database}]", masterConn)
                    createDatabaseCmd.ExecuteNonQuery()
                    Console.WriteLine($"Database '{_database}' created successfully")
                End If

                masterConn.Close()
            End Using

            Console.WriteLine($"Database '{_database}' initialized successfully")

        Catch ex As Exception
            Console.WriteLine($"Error initializing database: {ex.Message}")
            Throw New Exception($"ไม่สามารถสร้างฐานข้อมูลได้: {ex.Message}", ex)
        End Try
    End Sub

    ''' <summary>
    ''' เพิ่มข้อมูลการสแกนใหม่
    ''' </summary>
    ''' <param name="record">ข้อมูลการสแกน</param>
    ''' <returns>ID ของรายการที่เพิ่ม</returns>
    Public Shared Function AddScanRecord(record As ScanDataRecord) As Integer
        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim insertCmd As New SqlCommand(
                    "INSERT INTO ScanRecords " &
                    "(ScanDateTime, ProductCode, ReferenceCode, Quantity, DateCode, IsValid, OriginalData, ExtractedData, ValidationMessages, ComputerName, UserName) " &
                    "VALUES (@ScanDateTime, @ProductCode, @ReferenceCode, @Quantity, @DateCode, @IsValid, @OriginalData, @ExtractedData, @ValidationMessages, @ComputerName, @UserName); " &
                    "SELECT SCOPE_IDENTITY();", conn)

                ' เพิ่มพารามิเตอร์
                insertCmd.Parameters.AddWithValue("@ScanDateTime", record.ScanDateTime)
                insertCmd.Parameters.AddWithValue("@ProductCode", If(record.ProductCode, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@ReferenceCode", If(record.ReferenceCode, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@Quantity", record.Quantity)
                insertCmd.Parameters.AddWithValue("@DateCode", If(record.DateCode, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@IsValid", record.IsValid)
                insertCmd.Parameters.AddWithValue("@OriginalData", If(record.OriginalData, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@ExtractedData", If(record.ExtractedData, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@ValidationMessages", If(record.ValidationMessages, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@ComputerName", If(record.ComputerName, DBNull.Value))
                insertCmd.Parameters.AddWithValue("@UserName", If(record.UserName, DBNull.Value))

                ' ดึง ID ที่เพิ่มล่าสุด
                Dim id As Integer = Convert.ToInt32(insertCmd.ExecuteScalar())
                record.Id = id

                conn.Close()

                Console.WriteLine($"Added scan record with ID: {id}")
                Return id
            End Using
 
        Catch ex As Exception
            Console.WriteLine($"Error adding scan record: {ex.Message}")
            Throw New Exception($"ไม่สามารถเพิ่มข้อมูลการสแกนได้: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' ดึงข้อมูลประวัติการสแกนทั้งหมด
    ''' </summary>
    ''' <param name="limit">จำนวนรายการสูงสุดที่ต้องการดึง</param>
    ''' <returns>รายการข้อมูลการสแกน</returns>
    Public Shared Function GetScanHistory(Optional limit As Integer = 1000) As List(Of ScanDataRecord)
        Try
            Dim results As New List(Of ScanDataRecord)()

            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim limitClause As String = If(limit > 0, $" TOP {limit}", "")
                Dim selectCmd As New SqlCommand($"SELECT{limitClause} * FROM ScanRecords ORDER BY ScanDateTime DESC", conn)

                Using reader As SqlDataReader = selectCmd.ExecuteReader()
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

                conn.Close()
            End Using

            Console.WriteLine($"Retrieved {results.Count} scan records")
            Return results

        Catch ex As Exception
            Console.WriteLine($"Error getting scan history: {ex.Message}")
            Throw New Exception($"ไม่สามารถดึงข้อมูลประวัติการสแกนได้: {ex.Message}", ex)
        End Try
    End Function

    ''' <summary>
    ''' ลบข้อมูลการสแกนตาม ID
    ''' </summary>
    ''' <param name="id">ID ของรายการที่ต้องการลบ</param>
    ''' <returns>True ถ้าลบสำเร็จ, False ถ้าไม่สำเร็จ</returns>
    Public Shared Function DeleteScanRecord(id As Integer) As Boolean
        Try
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim deleteCmd As New SqlCommand("DELETE FROM ScanRecords WHERE Id = @Id", conn)
                deleteCmd.Parameters.AddWithValue("@Id", id)

                Dim rowsAffected As Integer = deleteCmd.ExecuteNonQuery()
                conn.Close()

                Console.WriteLine($"Deleted scan record ID {id}, rows affected: {rowsAffected}")
                Return rowsAffected > 0
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
            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim updateCmd As New SqlCommand(
                    "UPDATE ScanRecords SET " &
                    "ProductCode = @ProductCode, " &
                    "ReferenceCode = @ReferenceCode, " &
                    "Quantity = @Quantity, " &
                    "DateCode = @DateCode, " &
                    "IsValid = @IsValid, " &
                    "ExtractedData = @ExtractedData, " &
                    "ValidationMessages = @ValidationMessages " &
                    "WHERE Id = @Id", conn)

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
                conn.Close()

                Console.WriteLine($"Updated scan record ID {record.Id}, rows affected: {rowsAffected}")
                Return rowsAffected > 0
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

            Using conn As New SqlConnection(ConnectionString)
                conn.Open()

                Dim selectCmd As New SqlCommand("SELECT * FROM ScanRecords WHERE ProductCode LIKE @ProductCode ORDER BY ScanDateTime DESC", conn)
                selectCmd.Parameters.AddWithValue("@ProductCode", $"%{productCode}%")

                Using reader As SqlDataReader = selectCmd.ExecuteReader()
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
                
                conn.Close()
            End Using
            
            Console.WriteLine($"Found {results.Count} records for product code: {productCode}")
            Return results
            
        Catch ex As Exception
            Console.WriteLine($"Error searching by product code: {ex.Message}")
            Throw New Exception($"ไม่สามารถค้นหาข้อมูลตามรหัสผลิตภัณฑ์ได้: {ex.Message}", ex)
        End Try
    End Function
End Class

