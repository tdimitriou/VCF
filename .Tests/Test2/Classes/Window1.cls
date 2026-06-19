VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VCF.IWindow

Public Win As Window
Private WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1

Private Sub Class_Initialize()
    NewWindow Me, Win
End Sub

Private Sub Form_Click()
    MsgBox Application.Current.Base.Windows.Count
End Sub

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub IWindow_InitializeComponent()
    Set Form = Win.Form
End Sub
