VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubSalesOrderViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_Lines As ObservableCollection
Private m_Title As String
Private m_Subtitle As String
Private m_LineSeq As Long

Public Property Get Title() As String
    Title = m_Title
End Property

Public Property Let Title(ByVal Value As String)
    If m_Title = Value Then Exit Property
    m_Title = Value
    m_Event.OnPropertyChanged "Title"
End Property

Public Property Get Subtitle() As String
    Subtitle = m_Subtitle
End Property

Public Property Let Subtitle(ByVal Value As String)
    If m_Subtitle = Value Then Exit Property
    m_Subtitle = Value
    m_Event.OnPropertyChanged "Subtitle"
End Property

Public Property Get Lines() As ObservableCollection
    Set Lines = m_Lines
End Property

Public Property Get TotalText() As String
    TotalText = "Lines: " & CStr(m_Lines.Count) & " (layout-only - no DB)"
End Property

Public Property Get AddSampleLineCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "AddSampleLineMethod")
    Set AddSampleLineCommand = Cmd
End Property

Public Property Get ClearLinesCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "ClearLinesMethod")
    Set ClearLinesCommand = Cmd
End Property

Public Property Get BackCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "BackMethod")
    Set BackCommand = Cmd
End Property

Public Sub AddSampleLineMethod(Optional Parameter As Variant)
    Dim Label As String
    m_LineSeq = m_LineSeq + 1
    If IsMissing(Parameter) Or IsEmpty(Parameter) Or IsNull(Parameter) Then
        Label = "Item"
    Else
        Label = CStr(Parameter)
        If Len(Label) = 0 Then Label = "Item"
    End If
    m_Lines.Add Label & " #" & CStr(m_LineSeq) & "  -  EUR " & Format$(1.5 * m_LineSeq, "0.00")
    m_Event.OnPropertyChanged "TotalText"
    Debug.Print "DeNovoSmoke Sales AddLine count=" & CStr(m_Lines.Count)
End Sub

Public Sub ClearLinesMethod()
    m_Lines.Clear
    m_LineSeq = 0
    m_Event.OnPropertyChanged "TotalText"
    Debug.Print "DeNovoSmoke Sales ClearLines"
End Sub

Public Sub BackMethod()
    On Error Resume Next
    If Not modApp.Shell Is Nothing Then modApp.Shell.ShowMainMenu
End Sub

Public Sub EnsureSeedLines()
    If m_Lines.Count > 0 Then Exit Sub
    AddSampleLineMethod "Coffee"
    AddSampleLineMethod "Pastry"
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    Set m_Lines = New ObservableCollection
    m_Title = "Sales Order"
    m_Subtitle = "Layout-only shell (milestone 3)"
    m_LineSeq = 0
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not m_Event Is Nothing Then m_Event.Unregister
End Sub
