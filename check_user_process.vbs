'==============================================================================
' Auteur      : John Gonzalez
'==============================================================================

'==============================================================================
' Version     : 1.0.0
'==============================================================================

Set arguments = WScript.Arguments
script = "check_user_process"
version = "1.0.0"
author = "John Gonzalez"
verbose = False

For i = 0 To arguments.Count - 1
    If arguments.Item(i) = "-v" Then
        verbose = True
        Exit For
    ElseIf Left(arguments.Item(i), 1) = "-" Then
        WScript.Echo "Invalid option: " & arguments.Item(i)
        WScript.Quit 1
    End If
Next

If verbose Then
    WScript.Echo script & " - " & author & " - " & version
    WScript.Quit 0
End If

On Error Resume Next

Dim processName, userName, objWMIService, colProcesses, objProcess

processName = "<process>.exe"
userName = "<user>"

Set objWMIService = GetObject("winmgmts:\\.\root\cimv2")
Set colProcesses = objWMIService.ExecQuery("SELECT * FROM Win32_Process WHERE Name = '" & processName & "'")

If colProcesses.Count > 0 Then
    For Each objProcess in colProcesses
        objProcess.GetOwner strUser, strDomain
        If LCase(strUser) = LCase(userName) Then
            WScript.Echo "OK - " & processName & " process is running with user " & userName
            WScript.Quit(0) 'Exit with OK status (0)
        End If
    Next
End If

WScript.Echo "CRITICAL - " & processName & " process is not running with user " & userName
WScript.Quit(2) 'Exit with CRITICAL status (2)