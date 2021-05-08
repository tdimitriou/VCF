VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1

Private Win As Window

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub IWindow_InitializeComponent()
    Set Form = Win.Form
        
    Set Win.DataContext = New ShellWindowViewModel
    
    With New VCF.XAMLReader
        .LoadSuperclassData Me, New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\ShellWindow.xml")
    End With

End Sub

Private Sub Class_Initialize()
    VCF.NewWindow Me, Win
End Sub

Public Sub Show()
    Form.Show
End Sub
