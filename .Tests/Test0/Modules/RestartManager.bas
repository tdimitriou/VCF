Attribute VB_Name = "RestartManager"
Option Explicit

Private Declare Function ShellExecute _
                Lib "shell32.dll" _
                Alias "ShellExecuteA" (ByVal hwnd As Long, _
                                       ByVal lpOperation As String, _
                                       ByVal lpFile As String, _
                                       ByVal lpParameters As String, _
                                       ByVal lpDirectory As String, _
                                       ByVal nShowCmd As Long) As Long

Public Sub Restart(Optional AllowMultiInstance As Boolean = False)

    Dim sScriptFile As String

    '<Replace with specific Close Function for the Application if necessary>
    CloseApplication
    '</Replace with specific Close Function for the Application if necessary>
    
    sScriptFile = Environ$("TEMP") & "\r.vbs"
    CreateRestartScript sScriptFile, AllowMultiInstance
    Call ShellExecute(0, "Open", sScriptFile, "", "", 1)

End Sub

Private Sub CloseApplication()
    Dim frm As Form

    For Each frm In Forms
        Unload frm
    Next

    Cairo.WidgetForms.RemoveAll
End Sub

Private Sub CreateRestartScript(ScriptFileName As String, Optional AllowMultiInstance As Boolean = False)

    Dim lFile As Long
    Dim sCmd  As String

    lFile = FreeFile
    sCmd = App.Path

    If Right$(sCmd, 1) <> "\" Then sCmd = sCmd & "\"
    sCmd = sCmd & App.EXEName & ".exe"
    
    Open ScriptFileName For Output As #lFile
    Print #lFile, BuildScript(ScriptFileName, sCmd, AllowMultiInstance)
    Close #lFile

End Sub

Private Function BuildScript(ScriptFileName As String, ShellCommand As String, Optional AllowMultiInstance As Boolean = False) As String
    Dim cmd As String
    
    If AllowMultiInstance = False Then
        cmd = "IsStillRunning = 1" & vbCrLf & _
                "For i = 0 to 20" & vbCrLf & _
                "   If CheckRunningApp(""" & App.EXEName & ".exe" & """) = 0 Then" & vbCrLf & _
                "       IsStillRunning = 0" & vbCrLf & _
                "       Exit For" & vbCrLf & _
                "   End If" & vbCrLf & _
                "   WScript.Sleep 50" & vbCrLf & _
                "Next" & vbCrLf & vbCrLf
    
        cmd = cmd & _
                "If IsStillRunning = 0 Then" & vbCrLf & _
                "   StartApp" & vbCrLf & _
                "   KillScriptFile" & vbCrLf & _
                "Else" & vbCrLf & _
                "   MsgBox ""Could not restart " & App.Title & "."" , vbInformation + vbOKOnly, """ & App.Title & """" & vbCrLf & _
                "   KillScriptFile" & vbCrLf & _
                "End If" & vbCrLf & vbCrLf
    Else
        cmd = "WScript.Sleep 50" & vbCrLf & _
                "StartApp" & vbCrLf & _
                "KillScriptFile" & vbCrLf & vbCrLf
    End If
    
    cmd = cmd & _
            "Sub StartApp" & vbCrLf & _
            "   On Error Resume Next" & vbCrLf & _
            "   Set objShell = WScript.CreateObject(""WScript.Shell"")" & vbCrLf & _
            "   'Triple Quotes for the filenames with spaces to be handled properly" & vbCrLf & _
            "   objShell.Run (""""""" & ShellCommand & """"""")" & vbCrLf & _
            "   If Err.Number <> 0 Then" & vbCrLf & _
            "      MsgBox ""Could not restart " & App.Title & "."" , vbInformation + vbOKOnly, """ & App.Title & """" & vbCrLf & _
            "   End If" & vbCrLf & _
            "End Sub" & vbCrLf & vbCrLf
            
    cmd = cmd & _
            "Sub KillScriptFile" & vbCrLf & _
            "   Set obj = CreateObject(""Scripting.FileSystemObject"")" & vbCrLf & _
            "   obj.DeleteFile (""" & ScriptFileName & """)" & vbCrLf & _
            "End Sub" & vbCrLf & vbCrLf
    
    cmd = cmd & _
            "Function CheckRunningApp(AppName)" & vbCrLf & _
            "   IsRunning = 0" & vbCrLf & _
            "   'The Local Computer Name is "".""" & vbCrLf & _
            "   Set objWMIService = GetObject(""winmgmts:\\.\root\cimv2"")" & vbCrLf & _
            "   sQuery = ""SELECT * FROM Win32_Process""" & vbCrLf & _
            "   Set objItems = objWMIService.ExecQuery(sQuery)" & vbCrLf & _
            "   'iterate all item(s)" & vbCrLf & _
            "   For Each objItem In objItems" & vbCrLf & _
            "       If UCase(objItem.Name)=UCase(AppName) Then" & vbCrLf & _
            "           IsRunning = 1" & vbCrLf & _
            "           Exit For" & vbCrLf & _
            "       End If" & vbCrLf & _
            "   Next" & vbCrLf & _
            "   CheckRunningApp = IsRunning" & vbCrLf & _
            "End Function" & vbCrLf

    BuildScript = cmd
End Function
