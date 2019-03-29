Attribute VB_Name = "AMC"
Option Explicit

'Application Mode Changer
'Copyright (c) 2002 - 2004 Nir Sofer
'
'Web site: http://nirsoft.mirrorz.com
'
'This small utility allows you to change the mode of an executable file (.exe) from
'GUI mode to console mode and vise versa.
'You can use this utility for creating console applications in Visual Basic environment.
'
'In order to do this, follow the instructions below:
'1. Write a console application in Visual Basic by using API calls
'(You can learn how to do it from my sample project)
'2. Create an executable file from your VB project.
'3. Run this utility, select your executable file, and click the
'"Change Mode" button. Your executable will be switched from GUI mode to console mode.
'
'If you successfully do the above 3 steps, you'll get a real console application !

' In use by ChangeMode
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" _
(ByVal lpFileName As String, _
ByVal dwDesiredAccess As Long, _
ByVal dwShareMode As Long, ByVal _
lpSecurityAttributes As Long, _
ByVal dwCreationDisposition As Long, _
ByVal dwFlagsAndAttributes As Long, _
ByVal hTemplateFile As Long) As Long

' In use by ChangeMode
Private Declare Function CloseHandle Lib "kernel32" _
(ByVal hObject As Long) As Long

' In use by ChangeMode
Private Declare Function ReadFile Lib "kernel32" _
(ByVal hFile As Long, _
lpBuffer As Any, _
ByVal nNumberOfBytesToRead As Long, _
lpNumberOfBytesRead As Long, _
ByVal lpOverlapped As Long) As Long

' In use by ChangeMode
Private Declare Function WriteFile Lib "kernel32" _
(ByVal hFile As Long, _
lpBuffer As Any, _
ByVal nNumberOfBytesToWrite As Long, _
lpNumberOfBytesWritten As Long, _
ByVal lpOverlapped As Long) As Long

' In use by ChangeMode
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
(Destination As Any, _
Source As Any, _
ByVal Length As Long)

' In use by ChangeMode
Private Declare Function SetFilePointer Lib "kernel32" _
(ByVal hFile As Long, _
ByVal lDistanceToMove As Long, _
lpDistanceToMoveHigh As Long, _
ByVal dwMoveMethod As Long) As Long

'' In generic use
'' https://social.msdn.microsoft.com/Forums/sqlserver/en-US/d6e76731-8e3b-465f-9d5a-12c6498d6b6c/how-to-return-exit-code-from-vb6-form?forum=winforms
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
' In use by ChangeMode
Private Const FILE_BEGIN = 0

' In use by ChangeMode
Private Const INVALID_HANDLE_VALUE = -1
Private Const GENERIC_READ = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3

' In use by GetErrorDesc
Private Const ERROR_CANNOT_WRITE_FILE = 1
Private Const ERROR_BAD_FILE_FORMAT = 2
Private Const ERROR_CANNOT_READ_FILE = 3
Private Const ERROR_CANNOT_OPEN_FILE = 4

' pseudo consts, see setup()
Public APP_NAME As String
Public VER As String
Public DEBUGGER As Boolean

Public SUBSYSTEMS() As Variant
Public SUBSYSTEM_INVALID_NAME As String
Public SUBSYSTEM_INVALID_DESC As String
Public Function Subsys_construct(ByVal ID As Integer, Name As String, Desc As String) As String
    Subsys_construct = CStr(ID) + " " + Name + " " + Desc
End Function



Public Function setup()
    APP_NAME = "AMC Patcher CLI"
    VER = CStr(App.Major) + "." + CStr(App.Minor)
    DEBUGGER = GetRunningInIDE()
    
    ' ftp://www.esrf.eu/scisoft/GrWinlib/CURRENT/GrWinTk/src/CheckUI.c
    ' http://eljay.github.io/directx-sys/winapi/winnt/
    ' https://ebourg.github.io/jsign/apidocs/src-html/net/jsign/pe/Subsystem.html
    
    ' Application Modes:
    ' Has to be in the chronological ascending order by design
    SUBSYSTEMS = Array( _
        Subsys_construct(0, "UNKNOWN", "Unknown system"), _
        Subsys_construct(1, "NATIVE", "Image doesn't require a subsystem"), _
        Subsys_construct(2, "WINDOWS_GUI", "Image runs in the Windows GUI subsystem"), _
        Subsys_construct(3, "WINDOWS_CUI", "Image runs in the Windows CLI subsystem (console)"), _
        Subsys_construct(4, "?", "No description"), _
        Subsys_construct(5, "OS2_CUI", "image runs in the OS/2 character subsystem"), _
        Subsys_construct(6, "?", "No description"), _
        Subsys_construct(7, "POSIX_CUI", "image runs in the Posix character subsystem"), _
        Subsys_construct(8, "NATIVE_WINDOWS", "image is a native Win9x driver"), _
        Subsys_construct(9, "WINDOWS_CE_GUI", "Image runs in the Windows CE subsystem"), _
        Subsys_construct(10, "EFI_APPLICATION", "An Extensible Firmware Interface (EFI) application"), _
        Subsys_construct(11, "EFI_BOOT_SERVICE_DRIVER", "An EFI driver with boot services"), _
        Subsys_construct(12, "EFI_RUNTIME_DRIVER", "An EFI driver with run-time services"), _
        Subsys_construct(13, "EFI_ROM", "An EFI ROM image"), _
        Subsys_construct(14, "IMAGE_SUBSYSTEM_XBOX", "No description"), _
        Subsys_construct(15, "?", "No description"), _
        Subsys_construct(16, "IMAGE_SUBSYSTEM_WINDOWS_BOOT_APPLICATION", "No description") _
    )

    SUBSYSTEM_INVALID_NAME = "?"
    SUBSYSTEM_INVALID_DESC = "No description"
End Function

' return something from the number
Public Function Subsys_ret(ByVal i As Integer) As String
    If (i <= (UBound(SUBSYSTEMS) - LBound(SUBSYSTEMS))) Then
        Subsys_ret = SUBSYSTEMS(i)
    Else
        Subsys_ret = Subsys_construct(i, SUBSYSTEM_INVALID_NAME, SUBSYSTEM_INVALID_DESC)
    End If
End Function

Private Function GetErrorDesc(iError As Integer) As String
    Select Case iError
        Case ERROR_CANNOT_WRITE_FILE
            GetErrorDesc = "Cannot write to the file"
        
        Case ERROR_BAD_FILE_FORMAT
            GetErrorDesc = "Bad file format"
        
        Case ERROR_CANNOT_READ_FILE
            GetErrorDesc = "Cannot read from the file"
        
        Case ERROR_CANNOT_OPEN_FILE
            GetErrorDesc = "Cannot open the file"
    End Select
End Function

Public Function ChangeMode(sFilename As String, iAppMode As Integer) As Integer
    Dim hFile           As Long
    Dim iError          As Integer
    Dim buffer(1024)    As Byte
    Dim lBytesReadWrite As Long
    Dim lPELocation     As Long
    
    ' just to be safe
    iError = 0
    
    'Open the file with Win32 API call
    hFile = CreateFile(sFilename, _
    GENERIC_READ Or GENERIC_WRITE, 0, 0, OPEN_EXISTING, 0, 0)
    If hFile <> INVALID_HANDLE_VALUE Then
        'Read the Dos header.
        If ReadFile(hFile, buffer(0), &H40, lBytesReadWrite, 0) <> 0 Then
            If buffer(0) = &H4D And buffer(1) = &H5A Then
                'Get the location of the portable executable header.
                CopyMemory lPELocation, buffer(&H3C), 4
                If lPELocation = 0 Then
                    iError = ERROR_BAD_FILE_FORMAT
                Else
                    SetFilePointer hFile, lPELocation, ByVal 0, FILE_BEGIN
                    'Read the portable executable signature.
                    If ReadFile(hFile, buffer(0), 2, lBytesReadWrite, 0) <> 0 Then
                        'Check the signature.
                        If buffer(0) = &H50 And buffer(1) = &H45 Then
                            buffer(0) = iAppMode
                            SetFilePointer hFile, lPELocation + &H5C, ByVal 0, FILE_BEGIN
                            'Change the byte in the executable file to the desired application mode.
                            If WriteFile(hFile, buffer(0), 1, lBytesReadWrite, 0) = 0 Then
                                iError = ERROR_CANNOT_WRITE_FILE
                            End If
                            
                        Else
                            iError = ERROR_BAD_FILE_FORMAT
                        End If
                    Else
                        iError = ERROR_CANNOT_READ_FILE
                    End If
                End If
                
            Else
                iError = ERROR_BAD_FILE_FORMAT
            End If
        Else
            iError = ERROR_CANNOT_READ_FILE
        End If
                
        'Close the file
        CloseHandle hFile
    Else
        iError = ERROR_CANNOT_OPEN_FILE
    End If
    ChangeMode = iError
End Function

Public Sub output_err(errMsg As String)
    CLI.Sendln "Error: " & errMsg
End Sub

Public Sub output_result(ByVal res As Integer, ByVal iError As Integer)
    If iError = 0 Then
        CLI.Sendln "Success. The application mode was successfully changed to"
        CLI.Sendln Subsys_ret(res)
    Else
        output_err GetErrorDesc(iError)
    End If
    
End Sub

Public Function quit(code As Integer)
    On Error Resume Next

    CLI.Send vbNewLine

    If DEBUGGER Then
        Debug.Print "End"
    Else
        ExitProcess code
    End If
End Function

' https://stackoverflow.com/a/9068210
Public Function GetRunningInIDE() As Boolean
   Dim x As Long
   Debug.Assert Not TestIDE(x)
   GetRunningInIDE = x = 1
End Function

' https://stackoverflow.com/a/9068210
Private Function TestIDE(x As Long) As Boolean
    x = 1
End Function


