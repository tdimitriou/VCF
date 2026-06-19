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

Private Win As Window
Attribute Win.VB_VarHelpID = -1
Private WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents NamedItems As ObservableDictionary
Attribute NamedItems.VB_VarHelpID = -1

Public Sub Show()
    Win.Show
End Sub

Private Sub Class_Initialize()
    NewWindow Me, Win
End Sub

'CSEH: ErrMsgBox
Private Sub InitializeComponent()
        '<EhHeader>
        On Error GoTo InitializeComponent_Err
        '</EhHeader>
        
        Win.DesignWidth = 1024
        Win.DesignHeight = 768
    
        Set Form = Win.Form
        Set NamedItems = Win.NamedChildren

        '<EhFooter>
        Exit Sub

InitializeComponent_Err:
        Debug.Print Err.Description & vbCrLf & _
               "in Test1.ShellWindow.InitializeComponent " & _
               "at line " & Erl
        Resume Next
        '</EhFooter>
End Sub

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub IWindow_InitializeComponent()
    InitializeComponent
End Sub

