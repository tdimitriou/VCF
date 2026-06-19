VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellWindowViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get ButtonClick() As ICommand
    Dim Cmd As ActionCommand
    Set Cmd = New ActionCommand
    
    Cmd.Initialize VCF.NewFunction(Me, "ButtonClickMethod")
    
    Set ButtonClick = Cmd
    
End Property

Public Sub ButtonClickMethod()

    Dim ActiveThemeName As String
    ActiveThemeName = MyApp.Base.Resources("ThemesManager").ActiveThemeName
    
    MyApp.Base.Resources("ThemesManager").ActiveThemeName = IIf((ActiveThemeName = "Default"), "Dark", "Default")
    
    Dim DD As IWindow
    Set DD = ShellWindow
        
    'MsgBox DD.Base.DependencyProperties.GetValue("BorderStyle")
    'MsgBox DD.Base.Form.BorderStyle
    
'    MsgBox DD.Base.Style.Key
    'MsgBox MyApp.Base.Resources("ThemesManager").ActiveThemeName
End Sub

