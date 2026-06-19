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
Implements IDependencyObject

Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1

Private Win As Window
Private WithEvents WinSubclass As VCF.Window
Attribute WinSubclass.VB_VarHelpID = -1

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode > 100 Then MsgBox KeyCode
End Sub

Private Property Get IDependencyObject_Children() As Object
    Dim Base As IDependencyObject
    Set Base = Win
    Set IDependencyObject_Children = Base.Children
End Property

Private Property Get IDependencyObject_DependencyProperties() As VCF.DependencyProperties
    Dim Base As IDependencyObject
    Set Base = Win
    Set IDependencyObject_DependencyProperties = Base.DependencyProperties
End Property

Private Property Get IDependencyObject_Parent() As VCF.IDependencyObject
    Dim Base As IDependencyObject
    Set Base = Win
    Set IDependencyObject_Parent = Base.Parent
End Property

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
    On Error Resume Next
    Set WinSubclass = VCF.NewWindow(Me, Win)
    If Err Then Debug.Print Err.Description, Err.Source, Err.Number
End Sub

Public Sub Show()
    Form.Show
End Sub

Private Sub WinSubclass_WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal Subclass As vbRichClient5.cSubClass, CancelDefaultProc As Boolean)
    
    If Msg = 256 And wParam = 120 Then
        Result = Subclass.CallWindowProc(256, 121, lParam)
        CancelDefaultProc = True
    End If
    
End Sub
