Attribute VB_Name = "CLI"
'Console Application Sample
'Written by Nir Sofer.
'
'Web site: http://nirsoft.mirrorz.com
'
'In order to create a console application from this code,
' follow the instructions below:
'1. Make an executable file from this project.
'2. Run the "Application Mode Changer" utility and change the
' mode of the executable file
' to "Console Application".



Public Declare Function GetStdHandle Lib "kernel32" _
(ByVal nStdHandle As Long) As Long

Private Declare Function WriteFile Lib "kernel32" _
(ByVal hFile As Long, _
lpBuffer As Any, _
ByVal nNumberOfBytesToWrite As Long, _
lpNumberOfBytesWritten As Long, _
lpOverlapped As Any) As Long

Public Const STD_OUTPUT_HANDLE = -11&

Private Type COORD
        x As Integer
        y As Integer
End Type

Private Type SMALL_RECT
        Left As Integer
        Top As Integer
        Right As Integer
        Bottom As Integer
End Type

Private Type CONSOLE_SCREEN_BUFFER_INFO
        dwSize As COORD
        dwCursorPosition As COORD
        wAttributes As Integer
        srWindow As SMALL_RECT
        dwMaximumWindowSize As COORD
End Type
Private Declare Function GetConsoleScreenBufferInfo Lib "kernel32" _
(ByVal hConsoleOutput As Long, _
lpConsoleScreenBufferInfo As CONSOLE_SCREEN_BUFFER_INFO) As Long

Private Declare Function SetConsoleTextAttribute Lib "kernel32" _
(ByVal hConsoleOutput As Long, ByVal wAttributes As Long) As Long

Public Const FOREGROUND_BLUE = &H1     '  text color contains blue.
Public Const FOREGROUND_GREEN = &H2     '  text color contains green.
Public Const FOREGROUND_INTENSITY = &H8     '  text color is intensified.
Public Const FOREGROUND_RED = &H4     '  text color contains red.



Private scrbuf      As CONSOLE_SCREEN_BUFFER_INFO
Private hOutput             As Long

'The following function writes the content of sText variable into the console window:
Public Function Send(sText As String) As Boolean
    Dim lWritten            As Long
    
    If WriteFile(hOutput, ByVal sText, Len(sText), lWritten, ByVal 0) = 0 Then
        WriteToConsole = False
    Else
        WriteToConsole = True
    End If
End Function

Public Function Sendln(t As String) As Boolean
    Sendln = Send(t + vbNewLine)
End Function

Public Function SetTextColour(colour As Long) As Long
    SetTextColour = SetConsoleTextAttribute(hOutput, colour)
End Function

Public Sub setup()
    hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
    GetConsoleScreenBufferInfo hOutput, scrbuf
End Sub

'Public Sub Main()
'   Dim scrbuf      As CONSOLE_SCREEN_BUFFER_INFO
'
'   'Get the standard output handle
'   hOutput = GetStdHandle(STD_OUTPUT_HANDLE)
'   GetConsoleScreenBufferInfo hOutput, scrbuf
'   WriteToConsole "Console Application Example In Visual Basic." & vbCrLf
'   WriteToConsole "Written by Nir Sofer" & vbCrLf
'   WriteToConsole "Web site: http://nirsoft.mirrorz.com" & vbCrLf & vbCrLf
'
'   'Change the text color to blue
'   SetConsoleTextAttribute hOutput, FOREGROUND_BLUE Or FOREGROUND_INTENSITY
'   WriteToConsole "Blue Color !!" & vbCrLf
'
'   'Change the text color to yellow
'   SetConsoleTextAttribute hOutput, FOREGROUND_RED Or _
'       FOREGROUND_GREEN Or FOREGROUND_INTENSITY
'   WriteToConsole "Yellow Color !!" & vbCrLf
'
'   'Restore the previous text attributes.
'   SetConsoleTextAttribute hOutput, scrbuf.wAttributes
'   If Len(Command$) <> 0 Then
'           Show the command line parameters:
'       WriteToConsole vbCrLf & "Command Line Parameters: " & Command$ & vbCrLf
'   End If
'End Sub

