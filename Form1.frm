VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "amc_patcher_cli"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    AMC.setup
    CLI.setup
    
    ' Initialise args
    Dim argw() As String
    Dim argc As Integer
    Dim TrimArg As String

    TrimArg = Trim(Command)
    argw = Split(TrimArg, " ")
    argc = UBound(argw) - LBound(argw) + 1 ' https://forums.windowssecrets.com/showthread.php/28214-counting-array-elements-(vb6)
    
    If (argc >= 2) Then ' If number of args suffice
        ' INPUT
        Dim last_arg As String
        Dim last_arg_int As Integer
        
        Dim path As String
        
        last_arg = argw(argc - 1)
        If (IsNumeric(last_arg)) Then
            last_arg_int = CInt(last_arg)
            
            path = Left(TrimArg, Len(TrimArg) - Len(last_arg) - 1)
            path = Replace(path, Chr(34), "")

            CLI.Sendln "Path: " + path
            CLI.Sendln "Arg: " + last_arg

            ' PROCESS
            Dim buf As Integer
            buf = ChangeMode(path, last_arg_int)
            
            ' OUTPUT

            AMC.output_result last_arg_int, buf
            quit buf
        Else
            AMC.output_err "Invalid mode (number expected)"
            quit -1
            ' Error bad subsys
        End If
    Else
        showHelp
        quit 0
    End If
End Sub

Private Sub showHelp()
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "AppModeChange - CLI mod v" + VER
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "(Original GUI code by Nirsoft)"
        CLI.Sendln ""
        
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "USAGE:"
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
        CLI.Sendln "amc <path_to_app> <new_mode>"
        CLI.Sendln ""
        
        CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "FOR EXAMPLE:"
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
        CLI.Sendln "amc " + Chr(34) + "C:/Projects/My supa CLI project/Project1.exe" + Chr(34) + " 3"
        CLI.Sendln "(to set the Project1 application to the CLI mode)"
        CLI.Sendln ""
        
        
        CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE Or CLI.FOREGROUND_INTENSITY
        CLI.Sendln "MANUAL:"
        CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
        CLI.Sendln "<path_to_app> - Path to your executable. " + Chr(34) + "-tolerable"
        CLI.Sendln ""
        CLI.Sendln "<new_mode> - New app SUBSYSTEM mode to set"
        CLI.Sendln "Informally, one'd need to only know of modes: 2 (CLI) and 3 (GUI)"
        CLI.Sendln "But below all known modes are listed:"
        
        Dim i As Integer

        For i = LBound(AMC.SUBSYSTEMS) To UBound(AMC.SUBSYSTEMS)
            CLI.SetTextColour CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
            CLI.Send vbTab + "* "
            CLI.SetTextColour CLI.FOREGROUND_RED Or CLI.FOREGROUND_GREEN Or CLI.FOREGROUND_BLUE
            CLI.Sendln CStr(AMC.SUBSYSTEMS(i))
        Next
End Sub
