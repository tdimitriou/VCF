VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubLoginViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_UserList As ObservableCollection
Private m_PasswordText As String

Public Property Get UserList() As ObservableCollection
    Set UserList = m_UserList
End Property

Public Property Get PasswordText() As String
    PasswordText = m_PasswordText
End Property

Public Property Let PasswordText(ByVal Value As String)
    If m_PasswordText = Value Then Exit Property
    m_PasswordText = Value
    m_Event.OnPropertyChanged "PasswordText"
End Property

Public Property Get LoginCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "LoginMethod")
    Set LoginCommand = Cmd
End Property

Public Property Get ClearInput() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "ClearInputMethod")
    Set ClearInput = Cmd
End Property

Public Sub LoginMethod()
    Debug.Print "DeNovoSmoke Login (stub) PasswordText=" & m_PasswordText
End Sub

Public Sub ClearInputMethod()
    PasswordText = ""
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    Set m_UserList = New ObservableCollection
    m_UserList.Add "Admin"
    m_UserList.Add "Cashier"
    m_UserList.Add "Manager"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    m_Event.Unregister
End Sub
