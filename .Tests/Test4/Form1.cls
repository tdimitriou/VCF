VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1

Private MemDb As cMemDB
Private WithEvents LV As ListView
Attribute LV.VB_VarHelpID = -1
Private WithEvents Rs As cRecordset
Attribute Rs.VB_VarHelpID = -1

Private Win As Window

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub IWindow_InitializeComponent()
    Set Form = Win.Form

'104     CreateDb
'106     Set LV = New UnboundListView
'
'108     Form.WidgetRoot.Widgets.Add LV, "LV", 8, 8, 300, 300
'110     Set Rs = MemDb.GetTable("t1")
'112     LV.Base.ListCount = Rs.RecordCount
'114     LV.Base.ListIndex = 0: Rs.MoveFirst

        With New VCF.XAMLReader
            .LoadSuperclassData Me, New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\Numpad.xml")
        End With
        Debug.Print "IWindow_InitializeComponent"
End Sub

'CSEH: ErrMsgBox
Private Sub Class_Initialize()
        '<EhHeader>
        On Error GoTo Class_Initialize_Err
        '</EhHeader>
100     VCF.NewWindow Me, Win
    
102     'Set Form = Win.Form  ' Cairo.WidgetForms.Create
    
104     'CreateDb
106     'Set LV = New UnboundListView
    
108     'Form.WidgetRoot.Widgets.Add LV, "LV", 8, 8, 300, 300
110     'Set Rs = MemDb.GetTable("t1")
112     'LV.Base.ListCount = Rs.RecordCount
114     'LV.Base.ListIndex = 0: Rs.MoveFirst
        '<EhFooter>
        Exit Sub

Class_Initialize_Err:
        MsgBox Err.Description & vbCrLf & _
               "in Test3.Form1.Class_Initialize " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub CreateDb()
    Set MemDb = New_c.MemDb
    MemDb.Exec "create table t1(id integer not null primary key,name varchar(50) not null);"

    Dim i As Long
    Dim Sql As String
    
    For i = 0 To 9
        Sql = "insert into t1 values (" & i & ", 'Name " & i & "');"
        MemDb.Exec Sql
    Next
End Sub

Private Sub Form_KeyPress(KeyUnicode As Integer)
    Select Case KeyUnicode
    
        Case Asc("f"): Rs.MoveFirst
        Case Asc("n"): Rs.MoveNext
        Case Asc("p"): Rs.MovePrevious
        Case Asc("l"): Rs.MoveLast
    
        Case Asc("0"):
            Dim BaseUI As UIElementBase
            Dim IUIE As IUIElement
            Set IUIE = LV
            Set BaseUI = IUIE.Base
            MsgBox BaseUI.TryFindResource("Hello")
    End Select
End Sub

Private Sub LV_OwnerDrawItem(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
    CC.Rectangle 0, 0, CLng(dx), CLng(dy)
    CC.DrawText 0, 0, CLng(dx), CLng(dy), Rs.ValueMatrix(Index, 1)
End Sub

'#Region Rs Events
    Private Sub Rs_AddNew(ByVal NewRowIdxZeroBased As Long)
        LV.Base.ListCount = Rs.RecordCount
    End Sub
    
    Private Sub Rs_Delete(ByVal NewRowIdxZeroBased As Long)
        LV.Base.ListIndex = Rs.RecordCount
    End Sub
    
    Private Sub Rs_FieldChange(ByVal RowIdxZeroBased As Long, ByVal ColIdxZeroBased As Long)
        LV.Base.Widget.Refresh
    End Sub
    
    Private Sub Rs_Move(ByVal NewRowIdxZeroBased As Long)
        LV.Base.ListIndex = NewRowIdxZeroBased
        Debug.Print NewRowIdxZeroBased
    End Sub
    
    Private Sub Rs_QueryFinished()
        LV.Base.ListCount = Rs.RecordCount
        Debug.Print "QueryFinished"
    End Sub
    
    Private Sub Rs_Reset()
        LV.Base.ListCount = Rs.RecordCount
    End Sub
'#End Region

Public Sub Show()
    Form.Show
End Sub
