Attribute VB_Name = "modApp"
Option Explicit

Public MyApp As MyApp

'CSEH: ErrMsgBox
Public Sub Start()
        '<EhHeader>
        On Error GoTo Start_Err
        '</EhHeader>
Try:
        On Error GoTo Catch
        
100     Call Cairo.ImageList.AddImage("pic.png", "C:\Shared\Software Development\Icons\business-icons\png\books_256.png")
102     Call Cairo.ImageList.AddImage("menu.png", App.Path & "\Resources\MainMenu.png")
    
104     VCF.SetCustomConstructor New ObjectConstructor
    
    '<Various options to start application>
        ' Option 1: XAML
106     Set MyApp = New MyApp
    
        ' Option 2: Define StartupURI
        'MyApp.StartupURI = "ShellWindow"
        'MyApp.Run

        ' Option 3: define StartupObject as Argument of Run Method
        'MyApp.Run New ShellWindow
    '</Various options to start application>

    Exit Sub
    
Catch:
108     MsgBox Err.Description, , App.FileDescription
110     Err.Clear
        '<EhFooter>
        Exit Sub

Start_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Test1.modApp.Start " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub
