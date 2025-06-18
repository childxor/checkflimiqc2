Imports System.IO
Imports System.Diagnostics
Imports System.Windows.Forms

Public Class UpdateManager
    Private Shared ReadOnly CURRENT_EXE_PATH As String = Application.ExecutablePath
    Private Shared ReadOnly CURRENT_APP_FOLDER As String = Path.GetDirectoryName(Application.ExecutablePath)
    
    ''' <summary>
    ''' ได้รับพาธของเซิร์ฟเวอร์อัพเดทตาม network ปัจจุบัน
    ''' </summary>
    Private Shared ReadOnly Property UPDATE_SERVER_PATH As String
        Get
            Dim networkPath = NetworkPathManager.GetUpdateSystemPath()
            If Not String.IsNullOrEmpty(networkPath) Then
                Return networkPath
            End If
            
            ' ถ้าไม่พบ network ให้ส่งคืนข้อความแสดงข้อผิดพลาด
            Return "ERROR: ไม่สามารถเชื่อมต่อกับเครือข่าย OA หรือ FAB ได้"
        End Get
    End Property
    
    Private Shared ReadOnly Property VERSION_FILE As String
        Get
            Return Path.Combine(UPDATE_SERVER_PATH, "version.txt")
        End Get
    End Property
    
    Public Shared Function CheckForUpdates() As UpdateCheckResult
        Try
            ' ตรวจสอบการเชื่อมต่อ network ก่อน
            Dim networkResult = NetworkPathManager.CheckNetworkConnection()
            If Not networkResult.IsConnected Then
                Return New UpdateCheckResult() With {.HasUpdate = False, .ErrorMessage = $"ไม่สามารถเชื่อมต่อกับเครือข่าย: {networkResult.ErrorMessage}"}
            End If
            
            Console.WriteLine($"Update check - Connected to {networkResult.NetworkType} network ({networkResult.ServerIP})")
            
            If Not Directory.Exists(UPDATE_SERVER_PATH) Then
                Return New UpdateCheckResult() With {.HasUpdate = False, .ErrorMessage = $"ไม่พบโฟลเดอร์อัปเดต: {UPDATE_SERVER_PATH}"}
            End If
            
            Dim currentVersion = GetCurrentVersion()
            Dim serverVersion = GetServerVersion()
            
            Console.WriteLine($"Current Version: {currentVersion}")
            Console.WriteLine($"Server Version: {serverVersion}")
            
            If serverVersion > currentVersion Then
                Return New UpdateCheckResult() With {
                    .HasUpdate = True,
                    .CurrentVersion = currentVersion,
                    .NewVersion = serverVersion,
                    .UpdatePath = UPDATE_SERVER_PATH
                }
            End If
            
            Return New UpdateCheckResult() With {.HasUpdate = False, .CurrentVersion = currentVersion, .NewVersion = serverVersion}
            
        Catch ex As Exception
            Return New UpdateCheckResult() With {.HasUpdate = False, .ErrorMessage = ex.Message}
        End Try
    End Function
    
    Private Shared Function GetCurrentVersion() As Version
        Try
            ' ดูจากไฟล์ version.txt ในโฟลเดอร์โปรแกรมก่อน
            Dim localVersionFile = GetLocalVersionPath()
            If File.Exists(localVersionFile) Then
                Dim versionText = File.ReadAllText(localVersionFile).Trim()
                If Not String.IsNullOrEmpty(versionText) Then
                    Return New Version(versionText)
                End If
            End If
            
            ' ถ้าไม่มีไฟล์ หรือไฟล์เสีย ให้ดูจาก Assembly
            Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            
        Catch ex As Exception
            ' ถ้าเกิดข้อผิดพลาด ให้ใช้ Assembly version
            Return System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
        End Try
    End Function

    Private Shared Function GetServerVersion() As Version
        Try
            If File.Exists(VERSION_FILE) Then
                Dim versionText = File.ReadAllText(VERSION_FILE).Trim()
                If Not String.IsNullOrEmpty(versionText) Then
                    Return New Version(versionText)
                End If
            End If
            
            ' หรือดูจาก Properties ของไฟล์ exe ในเซิร์ฟเวอร์
            Dim serverExe = Path.Combine(UPDATE_SERVER_PATH, "CheckFlimIQC2.exe")
            If File.Exists(serverExe) Then
                Return New Version(FileVersionInfo.GetVersionInfo(serverExe).FileVersion)
            End If
            
            Return New Version("0.0.0.0")
            
        Catch ex As Exception
            Console.WriteLine($"Error getting server version: {ex.Message}")
            Return New Version("0.0.0.0")
        End Try
    End Function
    
    Public Shared Function PerformUpdate(updateSourcePath As String) As Boolean
        Try
            Dim currentFolder = CURRENT_APP_FOLDER
            Dim backupFolder = currentFolder & "_backup"
            Dim tempFolder = currentFolder & "_temp"
            
            Console.WriteLine($"Starting update process...")
            Console.WriteLine($"Current folder: {currentFolder}")
            Console.WriteLine($"Update source: {updateSourcePath}")
            Console.WriteLine($"Backup folder: {backupFolder}")
            Console.WriteLine($"Temp folder: {tempFolder}")
            
            ' ลบโฟลเดอร์ backup และ temp เก่า (ถ้ามี)
            If Directory.Exists(backupFolder) Then
                Directory.Delete(backupFolder, True)
            End If
            If Directory.Exists(tempFolder) Then
                Directory.Delete(tempFolder, True)
            End If
            
            ' สร้าง Batch Script สำหรับอัปเดตทั้งโฟลเดอร์
            Dim batchScript = CreateUpdateBatchScript(currentFolder, updateSourcePath, backupFolder, tempFolder)
            
            Dim batchPath = Path.Combine(Path.GetTempPath(), "update_full_folder.bat")
            File.WriteAllText(batchPath, batchScript, System.Text.Encoding.Default)
            
            Console.WriteLine($"Batch script created: {batchPath}")
            Console.WriteLine("Batch script content:")
            Console.WriteLine(batchScript)
            
            ' รัน Batch และปิดโปรแกรม
            Process.Start(New ProcessStartInfo() With {
                .FileName = batchPath,
                .WindowStyle = ProcessWindowStyle.Normal,
                .CreateNoWindow = False
            })
            
            ' รอสักครู่แล้วปิดโปรแกรม
            System.Threading.Thread.Sleep(1000)
            Application.Exit()
            Return True
            
        Catch ex As Exception
            MessageBox.Show($"เกิดข้อผิดพลาดในการอัปเดต: {ex.Message}", "ข้อผิดพลาด", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End Try
    End Function
    
    Private Shared Function CreateUpdateBatchScript(currentFolder As String, sourceFolder As String, backupFolder As String, tempFolder As String) As String
        Dim exeName = Path.GetFileName(CURRENT_EXE_PATH)
        
        Return $"@echo off
chcp 65001 >nul
echo ===== เริ่มต้นการอัปเดตโปรแกรม =====
echo Current folder: {currentFolder}
echo Source folder: {sourceFolder}
echo Backup folder: {backupFolder}
echo Temp folder: {tempFolder}
echo.

echo รอให้โปรแกรมปิดสมบูรณ์...
timeout /t 3 /nobreak

echo ===== ขั้นตอนที่ 1: สำรองโฟลเดอร์เดิม =====
if exist ""{backupFolder}"" (
    echo ลบโฟลเดอร์สำรองเก่า...
    rmdir /s /q ""{backupFolder}""
)
echo สำรองโฟลเดอร์ปัจจุบัน...
xcopy ""{currentFolder}"" ""{backupFolder}\"" /e /i /h /y
if %errorlevel% neq 0 (
    echo ข้อผิดพลาด: ไม่สามารถสำรองโฟลเดอร์เดิมได้
    pause
    goto :end
)
echo ✅ สำรองโฟลเดอร์เดิมเสร็จสิ้น

echo.
echo ===== ขั้นตอนที่ 2: คัดลอกไฟล์ใหม่ =====
if exist ""{tempFolder}"" (
    echo ลบโฟลเดอร์ temp เก่า...
    rmdir /s /q ""{tempFolder}""
)
echo คัดลอกไฟล์ใหม่จากเซิร์ฟเวอร์...
xcopy ""{sourceFolder}"" ""{tempFolder}\"" /e /i /h /y
if %errorlevel% neq 0 (
    echo ข้อผิดพลาด: ไม่สามารถคัดลอกไฟล์ใหม่ได้
    pause
    goto :end
)
echo ✅ คัดลอกไฟล์ใหม่เสร็จสิ้น

echo.
echo ===== ขั้นตอนที่ 3: แทนที่ไฟล์ทั้งหมด =====
echo ลบไฟล์เก่าในโฟลเดอร์ปัจจุบัน...
for /f ""delims="" %%i in ('dir /b ""{currentFolder}""') do (
    if exist ""{currentFolder}\%%i"" (
        if /i ""%%i"" neq ""{Path.GetFileName(backupFolder)}"" (
            if /i ""%%i"" neq ""{Path.GetFileName(tempFolder)}"" (
                echo ลบ: %%i
                if exist ""{currentFolder}\%%i\*"" (
                    rmdir /s /q ""{currentFolder}\%%i""
                ) else (
                    del /f /q ""{currentFolder}\%%i""
                )
            )
        )
    )
)

echo คัดลอกไฟล์ใหม่ทั้งหมด...
xcopy ""{tempFolder}\*"" ""{currentFolder}\"" /e /h /y
if %errorlevel% neq 0 (
    echo ข้อผิดพลาด: ไม่สามารถแทนที่ไฟล์ได้ กำลังกู้คืนไฟล์เดิม...
    rmdir /s /q ""{currentFolder}""
    xcopy ""{backupFolder}"" ""{currentFolder}\"" /e /i /h /y
    echo ไฟล์ถูกกู้คืนเรียบร้อยแล้ว
    pause
    goto :end
)
echo ✅ แทนที่ไฟล์ทั้งหมดเสร็จสิ้น

echo.
echo ===== ขั้นตอนที่ 4: ตรวจสอบและเปิดโปรแกรม =====
if exist ""{currentFolder}\{exeName}"" (
    echo ✅ อัปเดตสำเร็จ! กำลังเปิดโปรแกรมใหม่...
    start """" ""{currentFolder}\{exeName}""
    
    echo ทำความสะอาดไฟล์ชั่วคราว...
    timeout /t 2 /nobreak
    if exist ""{tempFolder}"" rmdir /s /q ""{tempFolder}""
    if exist ""{backupFolder}"" rmdir /s /q ""{backupFolder}""
    
    echo ✅ อัปเดตเสร็จสมบูรณ์!
) else (
    echo ❌ ข้อผิดพลาด: ไม่พบไฟล์โปรแกรมหลังอัปเดต
    echo กำลังกู้คืนไฟล์เดิม...
    rmdir /s /q ""{currentFolder}""
    xcopy ""{backupFolder}"" ""{currentFolder}\"" /e /i /h /y
    start """" ""{currentFolder}\{exeName}""
    echo ไฟล์ถูกกู้คืนและเปิดโปรแกรมเดิมแล้ว
    pause
)

:end
echo การอัปเดตเสร็จสิ้น
timeout /t 3 /nobreak
del ""%~f0""
"
    End Function
    
    Private Shared Function GetLocalVersionPath() As String
        Return Path.Combine(CURRENT_APP_FOLDER, "version.txt")
    End Function
    
    Public Shared Sub InitializeVersionFile()
        Try
            Dim localVersionFile = GetLocalVersionPath()
            Dim assemblyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            
            ' แสดงข้อมูล debug
            Console.WriteLine($"🔍 InitializeVersionFile Debug Info:")
            Console.WriteLine($"   Application.ExecutablePath: {Application.ExecutablePath}")
            Console.WriteLine($"   CURRENT_APP_FOLDER: {CURRENT_APP_FOLDER}")
            Console.WriteLine($"   Local version file path: {localVersionFile}")
            Console.WriteLine($"   Assembly version: {assemblyVersion}")
            Console.WriteLine($"   Directory exists: {Directory.Exists(CURRENT_APP_FOLDER)}")
            
            Dim shouldUpdateVersionFile As Boolean = False
            Dim reason As String = ""
            
            ' ตรวจสอบสถานการณ์ต่างๆ
            If Not File.Exists(localVersionFile) Then
                shouldUpdateVersionFile = True
                reason = "ไม่พบไฟล์ version.txt"
            Else
                Try
                    Dim fileContent = File.ReadAllText(localVersionFile).Trim()
                    Console.WriteLine($"   Current file content: '{fileContent}'")
                    Dim fileVersion = New Version(fileContent)
                    If fileVersion <> assemblyVersion Then
                        shouldUpdateVersionFile = True
                        reason = $"เวอร์ชันไม่ตรงกัน (Assembly: {assemblyVersion}, File: {fileVersion})"
                    End If
                Catch ex As Exception
                    shouldUpdateVersionFile = True
                    reason = $"ไฟล์ version.txt เสียหาย: {ex.Message}"
                End Try
            End If
            
            ' อัปเดตไฟล์ version.txt ถ้าจำเป็น
            If shouldUpdateVersionFile Then
                Try
                    ' ตรวจสอบสิทธิ์การเขียนไฟล์
                    Dim directoryPath = Path.GetDirectoryName(localVersionFile)
                    If Not Directory.Exists(directoryPath) Then
                        Directory.CreateDirectory(directoryPath)
                        Console.WriteLine($"   สร้างโฟลเดอร์: {directoryPath}")
                    End If
                    
                    ' เขียนไฟล์
                    File.WriteAllText(localVersionFile, assemblyVersion.ToString())
                    Console.WriteLine($"✅ อัปเดต version.txt สำเร็จ: {reason}")
                    Console.WriteLine($"   เวอร์ชันใหม่: {assemblyVersion}")
                    Console.WriteLine($"   ไฟล์ตำแหน่ง: {localVersionFile}")
                    
                    ' ตรวจสอบว่าไฟล์ถูกเขียนจริงหรือไม่
                    If File.Exists(localVersionFile) Then
                        Dim writtenContent = File.ReadAllText(localVersionFile).Trim()
                        Console.WriteLine($"   เนื้อหาที่เขียน: '{writtenContent}'")
                    End If
                    
                Catch writeEx As Exception
                    Console.WriteLine($"❌ ไม่สามารถเขียนไฟล์ version.txt ได้: {writeEx.Message}")
                    Console.WriteLine($"   Path: {localVersionFile}")
                    Console.WriteLine($"   Directory writable: {IsDirectoryWritable(Path.GetDirectoryName(localVersionFile))}")
                End Try
            Else
                Console.WriteLine($"✅ version.txt ถูกต้องแล้ว: {assemblyVersion}")
            End If
            
        Catch ex As Exception
            Console.WriteLine($"❌ Error initializing version file: {ex.Message}")
            Console.WriteLine($"   Stack trace: {ex.StackTrace}")
        End Try
    End Sub
    
    ''' <summary>
    ''' ฟังก์ชันสำหรับซิงค์ไฟล์ version.txt ให้ตรงกับ Assembly version
    ''' </summary>
    Public Shared Sub SyncVersionWithAssembly()
        Try
            Dim localVersionFile = GetLocalVersionPath()
            Dim assemblyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            
            ' อัปเดตไฟล์ version.txt ให้ตรงกับ Assembly
            File.WriteAllText(localVersionFile, assemblyVersion.ToString())
            Console.WriteLine($"🔄 Synced version.txt with Assembly version: {assemblyVersion}")
            
        Catch ex As Exception
            Console.WriteLine($"❌ Error syncing version file: {ex.Message}")
        End Try
    End Sub
    
    ''' <summary>
    ''' ตรวจสอบว่าสามารถเขียนไฟล์ในโฟลเดอร์ได้หรือไม่
    ''' </summary>
    Private Shared Function IsDirectoryWritable(directoryPath As String) As Boolean
        Try
            If Not Directory.Exists(directoryPath) Then
                Return False
            End If
            
            Dim testFile = Path.Combine(directoryPath, "write_test.tmp")
            File.WriteAllText(testFile, "test")
            File.Delete(testFile)
            Return True
        Catch
            Return False
        End Try
    End Function
    
    ''' <summary>
    ''' ตรวจสอบความสอดคล้องระหว่าง Assembly version และไฟล์ version.txt
    ''' </summary>
    Public Shared Function CheckVersionConsistency() As VersionConsistencyResult
        Try
            Dim result As New VersionConsistencyResult()
            Dim assemblyVersion = System.Reflection.Assembly.GetExecutingAssembly().GetName().Version
            Dim localVersionFile = GetLocalVersionPath()
            
            result.AssemblyVersion = assemblyVersion
            
            If File.Exists(localVersionFile) Then
                Try
                    Dim fileVersionText = File.ReadAllText(localVersionFile).Trim()
                    result.FileVersion = New Version(fileVersionText)
                    result.IsConsistent = (result.AssemblyVersion = result.FileVersion)
                    result.FileExists = True
                Catch ex As Exception
                    result.IsConsistent = False
                    result.ErrorMessage = "ไฟล์ version.txt เสียหาย"
                    result.FileExists = True
                End Try
            Else
                result.IsConsistent = False
                result.ErrorMessage = "ไม่พบไฟล์ version.txt"
                result.FileExists = False
            End If
            
            Return result
            
        Catch ex As Exception
            Return New VersionConsistencyResult() With {
                .IsConsistent = False,
                .ErrorMessage = ex.Message
            }
        End Try
    End Function
End Class

Public Class UpdateCheckResult
    Public Property HasUpdate As Boolean = False
    Public Property CurrentVersion As Version
    Public Property NewVersion As Version
    Public Property UpdatePath As String = ""
    Public Property ErrorMessage As String = ""
End Class

Public Class VersionConsistencyResult
    Public Property AssemblyVersion As Version
    Public Property FileVersion As Version
    Public Property IsConsistent As Boolean = False
    Public Property FileExists As Boolean = False
    Public Property ErrorMessage As String = ""
    
    Public Overrides Function ToString() As String
        If IsConsistent Then
            Return $"✅ เวอร์ชันสอดคล้องกัน: {AssemblyVersion}"
        Else
            Return $"❌ เวอร์ชันไม่สอดคล้อง: Assembly={AssemblyVersion}, File={If(FileVersion, "N/A")} - {ErrorMessage}"
        End If
    End Function
End Class