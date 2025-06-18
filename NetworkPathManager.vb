Imports System.Net.NetworkInformation
Imports System.IO

''' <summary>
''' คลาสสำหรับจัดการพาธ network ทั้ง OA และ FAB
''' </summary>
Public Class NetworkPathManager
    
    ' ค่าคงที่สำหรับ IP addresses
    Private Shared ReadOnly OA_SERVER_IP As String = "10.24.179.2"
    Private Shared ReadOnly FAB_SERVER_IP As String = "172.24.0.3"
    
    ' ค่าคงที่สำหรับ base paths
    Private Shared ReadOnly BASE_SHARE_PATH As String = "OAFAB\OA2FAB"
    
    ' Timeout สำหรับการ ping (มิลลิวินาที)
    Private Shared ReadOnly PING_TIMEOUT As Integer = 3000
    
    ''' <summary>
    ''' ผลลัพธ์การตรวจสอบ network
    ''' </summary>
    Public Class NetworkCheckResult
        Public Property IsConnected As Boolean = False
        Public Property NetworkType As String = ""  ' "OA" หรือ "FAB"
        Public Property ServerIP As String = ""
        Public Property ErrorMessage As String = ""
        Public Property BasePath As String = ""
    End Class
    
    ''' <summary>
    ''' ตรวจสอบการเชื่อมต่อ network และกำหนดประเภท network
    ''' Logic ใหม่: ถ้าปิง 172.24.0.3 ไม่ได้ = OA, ถ้าปิงได้ทั้งสอง = FAB
    ''' </summary>
    ''' <returns>ผลลัพธ์การตรวจสอบ network</returns>
    Public Shared Function CheckNetworkConnection() As NetworkCheckResult
        Dim result As New NetworkCheckResult()
        
        Try
            Dim ping As New Ping()
            Dim canPingOA As Boolean = False
            Dim canPingFAB As Boolean = False
            
            ' ทดสอบการเชื่อมต่อ OA (10.24.179.2)
            Try
                Dim replyOa As PingReply = ping.Send(OA_SERVER_IP, PING_TIMEOUT)
                canPingOA = (replyOa.Status = IPStatus.Success)
                If canPingOA Then
                    Console.WriteLine($"OA network ({OA_SERVER_IP}) is reachable")
                End If
            Catch ex As Exception
                Console.WriteLine($"OA network test failed: {ex.Message}")
                canPingOA = False
            End Try
            
            ' ทดสอบการเชื่อมต่อ FAB (172.24.0.3)
            Try
                Dim replyFab As PingReply = ping.Send(FAB_SERVER_IP, PING_TIMEOUT)
                canPingFAB = (replyFab.Status = IPStatus.Success)
                If canPingFAB Then
                    Console.WriteLine($"FAB network ({FAB_SERVER_IP}) is reachable")
                End If
            Catch ex As Exception
                Console.WriteLine($"FAB network test failed: {ex.Message}")
                canPingFAB = False
            End Try
            
            ' กำหนดประเภทเครือข่ายตาม logic ใหม่
            If Not canPingFAB Then
                ' ถ้าปิง FAB (172.24.0.3) ไม่ได้ = เครือข่าย OA
                If canPingOA Then
                    result.IsConnected = True
                    result.NetworkType = "OA"
                    result.ServerIP = OA_SERVER_IP
                    result.BasePath = $"\\{OA_SERVER_IP}\{BASE_SHARE_PATH}"
                    Console.WriteLine($"Network determined as OA (FAB unreachable)")
                    Return result
                Else
                    result.IsConnected = False
                    result.ErrorMessage = "ไม่สามารถเชื่อมต่อกับเครือข่าย OA ได้"
                    Console.WriteLine("Both networks unreachable, but assuming OA environment")
                    Return result
                End If
            ElseIf canPingOA And canPingFAB Then
                ' ถ้าปิงได้ทั้งสอง = เครือข่าย FAB
                result.IsConnected = True
                result.NetworkType = "FAB"
                result.ServerIP = FAB_SERVER_IP
                result.BasePath = $"\\{FAB_SERVER_IP}\{BASE_SHARE_PATH}"
                Console.WriteLine($"Network determined as FAB (both networks reachable)")
                Return result
            ElseIf canPingFAB And Not canPingOA Then
                ' ถ้าปิงได้แค่ FAB = เครือข่าย FAB
                result.IsConnected = True
                result.NetworkType = "FAB"
                result.ServerIP = FAB_SERVER_IP
                result.BasePath = $"\\{FAB_SERVER_IP}\{BASE_SHARE_PATH}"
                Console.WriteLine($"Network determined as FAB (only FAB reachable)")
                Return result
            Else
                ' ไม่สามารถเชื่อมต่อได้เลย
                result.IsConnected = False
                result.ErrorMessage = "ไม่สามารถเชื่อมต่อกับเครือข่าย OA หรือ FAB ได้"
                Console.WriteLine("No network connection available")
                Return result
            End If
            
        Catch ex As Exception
            result.IsConnected = False
            result.ErrorMessage = $"เกิดข้อผิดพลาดในการตรวจสอบเครือข่าย: {ex.Message}"
            Console.WriteLine($"Network check error: {ex.Message}")
        End Try
        
        Return result
    End Function
    
    ''' <summary>
    ''' ได้รับพาธของไฟล์ Excel Database
    ''' </summary>
    ''' <returns>พาธของไฟล์ Excel หรือ String.Empty ถ้าไม่พบ network</returns>
    Public Shared Function GetExcelDatabasePath() As String
        Dim networkResult = CheckNetworkConnection()
        If networkResult.IsConnected Then
            Return Path.Combine(networkResult.BasePath, "Film charecter check\Database.xlsx")
        End If
        Return String.Empty
    End Function
    
    ''' <summary>
    ''' ได้รับพาธของไฟล์ Access Database
    ''' </summary>
    ''' <returns>พาธของไฟล์ Access Database หรือ String.Empty ถ้าไม่พบ network</returns>
    Public Shared Function GetAccessDatabasePath() As String
        Dim networkResult = CheckNetworkConnection()
        If networkResult.IsConnected Then
            Return Path.Combine(networkResult.BasePath, "Film charecter check\dbSystems\QRCodeScanner.accdb")
        End If
        Return String.Empty
    End Function
    
    ''' <summary>
    ''' ได้รับพาธของโฟลเดอร์ Debug Systems สำหรับการอัพเดท
    ''' </summary>
    ''' <returns>พาธของโฟลเดอร์ Debug Systems หรือ String.Empty ถ้าไม่พบ network</returns>
    Public Shared Function GetUpdateSystemPath() As String
        Dim networkResult = CheckNetworkConnection()
        If networkResult.IsConnected Then
            Return Path.Combine(networkResult.BasePath, "Film charecter check\DebugSystems\net8.0-windows\")
        End If
        Return String.Empty
    End Function
    
    ''' <summary>
    ''' ได้รับพาธของโฟลเดอร์หลักสำหรับ Film Character Check
    ''' </summary>
    ''' <returns>พาธของโฟลเดอร์หลัก หรือ String.Empty ถ้าไม่พบ network</returns>
    Public Shared Function GetFilmCharacterCheckPath() As String
        Dim networkResult = CheckNetworkConnection()
        If networkResult.IsConnected Then
            Return Path.Combine(networkResult.BasePath, "Film charecter check")
        End If
        Return String.Empty
    End Function
    
    ''' <summary>
    ''' ได้รับพาธของโฟลเดอร์ Drawing
    ''' </summary>
    ''' <returns>พาธของโฟลเดอร์ Drawing หรือ String.Empty ถ้าไม่พบ network</returns>
    Public Shared Function GetDrawingFolderPath() As String
        Dim networkResult = CheckNetworkConnection()
        If networkResult.IsConnected Then
            Return Path.Combine(networkResult.BasePath, "Film charecter check\Drawing")
        End If
        Return String.Empty
    End Function
    
    ''' <summary>
    ''' สร้างพาธแบบกำหนดเอง
    ''' </summary>
    ''' <param name="relativePath">พาธที่ต้องการต่อจาก base path</param>
    ''' <returns>พาธเต็ม หรือ String.Empty ถ้าไม่พบ network</returns>
    Public Shared Function GetCustomPath(relativePath As String) As String
        Dim networkResult = CheckNetworkConnection()
        If networkResult.IsConnected AndAlso Not String.IsNullOrEmpty(relativePath) Then
            Return Path.Combine(networkResult.BasePath, relativePath)
        End If
        Return String.Empty
    End Function
    
    ''' <summary>
    ''' ตรวจสอบว่าไฟล์หรือโฟลเดอร์มีอยู่หรือไม่
    ''' </summary>
    ''' <param name="path">พาธที่ต้องการตรวจสอบ</param>
    ''' <returns>True ถ้ามีอยู่, False ถ้าไม่มี</returns>
    Public Shared Function PathExists(path As String) As Boolean
        Try
            If String.IsNullOrEmpty(path) Then
                Return False
            End If
            
            Return File.Exists(path) OrElse Directory.Exists(path)
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' ได้รับข้อมูลสถานะของ network ปัจจุบัน
    ''' </summary>
    ''' <returns>ข้อความอธิบายสถานะ network</returns>
    Public Shared Function GetNetworkStatus() As String
        Dim networkResult = CheckNetworkConnection()
        
        If networkResult.IsConnected Then
            Return $"เชื่อมต่อกับเครือข่าย {networkResult.NetworkType} ({networkResult.ServerIP}) ✅{vbNewLine}Base Path: {networkResult.BasePath}"
        Else
            Return $"ไม่พบการเชื่อมต่อเครือข่าย ❌{vbNewLine}ข้อผิดพลาด: {networkResult.ErrorMessage}"
        End If
    End Function
    
End Class 