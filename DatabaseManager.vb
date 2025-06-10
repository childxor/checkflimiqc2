Imports System.Data.SqlClient
Imports System.Configuration
Imports System.IO
Imports System.Xml

Public Class DatabaseManager

#Region "Constants & Variables"
    Private Shared _connectionString As String = ""
    Private Shared _isConnected As Boolean = False
#End Region

#Region "Connection Management"
    ''' <summary>
    ''' เริ่มต้นการเชื่อมต่อฐานข้อมูล
    ''' </summary>
    Public Shared Function Initialize() As Boolean
        Try
            _connectionString = BuildConnectionString()
            Return TestConnection()
        Catch ex As Exception
            LogError($"Database initialization failed: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' สร้าง connection string จากการตั้งค่า
    ''' </summary>
    Private Shared Function BuildConnectionString() As String
        Try
            If File.Exists("Settings.config") Then
                Dim doc As New XmlDocument()
                doc.Load("Settings.config")
                
                Dim server As String = GetSettingValue(doc, "Server", "localhost")
                Dim database As String = GetSettingValue(doc, "Database", "")
                Dim username As String = GetSettingValue(doc, "Username", "")
                Dim password As String = GetSettingValue(doc, "Password", "")
                Dim integratedSecurity As Boolean = GetSettingValue(doc, "IntegratedSecurity", False)
                
                Dim builder As New SqlConnectionStringBuilder()
                builder.DataSource = server
                
                If Not String.IsNullOrEmpty(database) Then
                    builder.InitialCatalog = database
                End If
                
                If integratedSecurity Then
                    builder.IntegratedSecurity = True
                Else
                    builder.UserID = username
                    builder.Password = DecryptPassword(password)
                End If
                
                builder.ConnectTimeout = 30

                Return builder.ConnectionString
            End If
        Catch ex As Exception
            LogError($"Error building connection string: {ex.Message}")
        End Try
        
        Return ""
    End Function

    ''' <summary>
    ''' ทดสอบการเชื่อมต่อฐานข้อมูล
    ''' </summary>
    Public Shared Function TestConnection() As Boolean
        Try
            If String.IsNullOrEmpty(_connectionString) Then
                Return False
            End If

            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                _isConnected = True
                Return True
            End Using
        Catch ex As Exception
            _isConnected = False
            LogError($"Database connection test failed: {ex.Message}")
            Return False
        End Try
    End Function
#End Region

#Region "Scan Data Management"
    ''' <summary>
    ''' บันทึกข้อมูลการสแกน
    ''' </summary>
    Public Shared Function SaveScanData(scanData As ScanDataRecord) As Boolean
        Try
            If Not _isConnected Then
                Initialize()
            End If

            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                
                Dim sql As String = "
                    INSERT INTO ScanHistory 
                    (ScanDateTime, OriginalData, ExtractedData, ProductCode, 
                     ReferenceCode, Quantity, DateCode, IsValid, ValidationMessages, 
                     ComputerName, UserName)
                    VALUES 
                    (@ScanDateTime, @OriginalData, @ExtractedData, @ProductCode, 
                     @ReferenceCode, @Quantity, @DateCode, @IsValid, @ValidationMessages, 
                     @ComputerName, @UserName)"
                
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@ScanDateTime", scanData.ScanDateTime)
                    cmd.Parameters.AddWithValue("@OriginalData", scanData.OriginalData)
                    cmd.Parameters.AddWithValue("@ExtractedData", scanData.ExtractedData)
                    cmd.Parameters.AddWithValue("@ProductCode", scanData.ProductCode)
                    cmd.Parameters.AddWithValue("@ReferenceCode", scanData.ReferenceCode)
                    cmd.Parameters.AddWithValue("@Quantity", scanData.Quantity)
                    cmd.Parameters.AddWithValue("@DateCode", scanData.DateCode)
                    cmd.Parameters.AddWithValue("@IsValid", scanData.IsValid)
                    cmd.Parameters.AddWithValue("@ValidationMessages", scanData.ValidationMessages)
                    cmd.Parameters.AddWithValue("@ComputerName", Environment.MachineName)
                    cmd.Parameters.AddWithValue("@UserName", Environment.UserName)
                    
                    cmd.ExecuteNonQuery()
                End Using
            End Using
            
            Return True
            
        Catch ex As Exception
            LogError($"Error saving scan data: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' ดึงประวัติการสแกน
    ''' </summary>
    Public Shared Function GetScanHistory(Optional limit As Integer = 100) As List(Of ScanDataRecord)
        Dim history As New List(Of ScanDataRecord)()
        
        Try
            Console.WriteLine($"GetScanHistory called with limit: {limit}")
            Console.WriteLine($"IsConnected: {_isConnected}")
            Console.WriteLine($"ConnectionString: {If(String.IsNullOrEmpty(_connectionString), "Empty", "Set")}")
            
            If Not _isConnected Then
                Console.WriteLine("Not connected, attempting to initialize...")
                Initialize()
            End If

            If String.IsNullOrEmpty(_connectionString) Then
                Console.WriteLine("Connection string is empty, returning empty list")
                Return history
            End If

            Using conn As New SqlConnection(_connectionString)
                Console.WriteLine("Opening connection...")
                conn.Open()
                Console.WriteLine("Connection opened successfully")
                
                Dim sql As String = $"
                    SELECT TOP {limit} 
                        ScanDateTime, OriginalData, ExtractedData, ProductCode, 
                        ReferenceCode, Quantity, DateCode, IsValid, ValidationMessages,
                        ComputerName, UserName
                    FROM ScanHistory 
                    ORDER BY ScanDateTime DESC"
                
                Console.WriteLine($"Executing SQL: {sql}")
                
                Using cmd As New SqlCommand(sql, conn)
                    Using reader As SqlDataReader = cmd.ExecuteReader()
                        Console.WriteLine("SQL executed, reading data...")
                        
                        Dim recordCount As Integer = 0
                        While reader.Read()
                            recordCount += 1
                            
                            Dim record As New ScanDataRecord() With {
                                .ScanDateTime = If(IsDBNull(reader("ScanDateTime")), DateTime.MinValue, CDate(reader("ScanDateTime"))),
                                .OriginalData = If(IsDBNull(reader("OriginalData")), "", reader("OriginalData").ToString()),
                                .ExtractedData = If(IsDBNull(reader("ExtractedData")), "", reader("ExtractedData").ToString()),
                                .ProductCode = If(IsDBNull(reader("ProductCode")), "", reader("ProductCode").ToString()),
                                .ReferenceCode = If(IsDBNull(reader("ReferenceCode")), "", reader("ReferenceCode").ToString()),
                                .Quantity = If(IsDBNull(reader("Quantity")), "", reader("Quantity").ToString()),
                                .DateCode = If(IsDBNull(reader("DateCode")), "", reader("DateCode").ToString()),
                                .IsValid = If(IsDBNull(reader("IsValid")), False, CBool(reader("IsValid"))),
                                .ValidationMessages = If(IsDBNull(reader("ValidationMessages")), "", reader("ValidationMessages").ToString()),
                                .ComputerName = If(IsDBNull(reader("ComputerName")), "", reader("ComputerName").ToString()),
                                .UserName = If(IsDBNull(reader("UserName")), "", reader("UserName").ToString())
                            }
                            history.Add(record)
                        End While
                        
                        Console.WriteLine($"Read {recordCount} records from database")
                    End Using
                End Using
            End Using
            
        Catch ex As Exception
            Console.WriteLine($"Error in GetScanHistory: {ex.Message}")
            Console.WriteLine($"Stack trace: {ex.StackTrace}")
            LogError($"Error getting scan history: {ex.Message}")
        End Try
        
        Console.WriteLine($"Returning {history.Count} records")
        Return history
    End Function

    ''' <summary>
    ''' ลบข้อมูลเก่า
    ''' </summary>
    Public Shared Function CleanupOldData(daysToKeep As Integer) As Integer
        Try
            If Not _isConnected Then
                Initialize()
            End If

            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                
                Dim cutoffDate As DateTime = DateTime.Now.AddDays(-daysToKeep)
                Dim sql As String = "DELETE FROM ScanHistory WHERE ScanDateTime < @CutoffDate"
                
                Using cmd As New SqlCommand(sql, conn)
                    cmd.Parameters.AddWithValue("@CutoffDate", cutoffDate)
                    Return cmd.ExecuteNonQuery()
                End Using
            End Using
            
        Catch ex As Exception
            LogError($"Error cleaning up old data: {ex.Message}")
            Return 0
        End Try
    End Function
#End Region

#Region "Database Schema Management"
    ''' <summary>
    ''' สร้างตารางฐานข้อมูลหากยังไม่มี
    ''' </summary>
    Public Shared Function CreateTablesIfNotExists() As Boolean
        Try
            If Not _isConnected Then
                Initialize()
            End If

            Using conn As New SqlConnection(_connectionString)
                conn.Open()
                
                Dim sql As String = "
                    IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='ScanHistory' AND xtype='U')
                    BEGIN
                        CREATE TABLE ScanHistory (
                            ID int IDENTITY(1,1) PRIMARY KEY,
                            ScanDateTime datetime NOT NULL,
                            OriginalData nvarchar(max) NOT NULL,
                            ExtractedData nvarchar(500),
                            ProductCode nvarchar(100),
                            ReferenceCode nvarchar(100),
                            Quantity nvarchar(50),
                            DateCode nvarchar(50),
                            IsValid bit NOT NULL DEFAULT(0),
                            ValidationMessages nvarchar(max),
                            ComputerName nvarchar(100),
                            UserName nvarchar(100),
                            CreatedAt datetime DEFAULT(GETDATE())
                        )
                        
                        CREATE INDEX IX_ScanHistory_ScanDateTime ON ScanHistory(ScanDateTime DESC)
                        CREATE INDEX IX_ScanHistory_ProductCode ON ScanHistory(ProductCode)
                        CREATE INDEX IX_ScanHistory_IsValid ON ScanHistory(IsValid)
                    END"
                
                Using cmd As New SqlCommand(sql, conn)
                    cmd.ExecuteNonQuery()
                End Using
            End Using
            
            Return True
            
        Catch ex As Exception
            LogError($"Error creating database tables: {ex.Message}")
            Return False
        End Try
    End Function
#End Region

#Region "Utility Methods"
    Private Shared Function GetSettingValue(doc As XmlDocument, key As String, defaultValue As Object) As Object
        Try
            Dim node As XmlNode = doc.SelectSingleNode($"//Setting[@key='{key}']")
            If node IsNot Nothing Then
                Dim value As String = node.Attributes("value").Value
                Select Case defaultValue.GetType()
                    Case GetType(Boolean)
                        Return Boolean.Parse(value)
                    Case GetType(Integer)
                        Return Integer.Parse(value)
                    Case Else
                        Return value
                End Select
            End If
        Catch
        End Try
        Return defaultValue
    End Function

    Private Shared Function DecryptPassword(encryptedPassword As String) As String
        Try
            If String.IsNullOrEmpty(encryptedPassword) Then Return ""
            Dim bytes As Byte() = Convert.FromBase64String(encryptedPassword)
            Return System.Text.Encoding.UTF8.GetString(bytes)
        Catch
            Return encryptedPassword
        End Try
    End Function

    Private Shared Sub LogError(message As String)
        Try
            Dim logPath As String = Path.Combine(Application.StartupPath, "Logs")
            If Not Directory.Exists(logPath) Then
                Directory.CreateDirectory(logPath)
            End If
            
            Dim logFile As String = Path.Combine(logPath, $"Database_{DateTime.Now:yyyyMMdd}.log")
            Dim logEntry As String = $"{DateTime.Now:yyyy-MM-dd HH:mm:ss} - {message}{Environment.NewLine}"
            
            File.AppendAllText(logFile, logEntry)
        Catch
            ' ไม่ต้องทำอะไร
        End Try
    End Sub

    Public Shared ReadOnly Property IsConnected As Boolean
        Get
            Return _isConnected
        End Get
    End Property
#End Region

End Class

#Region "Supporting Classes"
''' <summary>
''' โครงสร้างข้อมูลการสแกน
''' </summary>
Public Class ScanDataRecord
    Public Property ScanDateTime As DateTime
    Public Property OriginalData As String
    Public Property ExtractedData As String
    Public Property ProductCode As String
    Public Property ReferenceCode As String
    Public Property Quantity As String
    Public Property DateCode As String
    Public Property IsValid As Boolean
    Public Property ValidationMessages As String
    Public Property ComputerName As String
    Public Property UserName As String
End Class
#End Region