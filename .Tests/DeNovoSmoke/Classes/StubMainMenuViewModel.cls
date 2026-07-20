VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubMainMenuViewModel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent
Private m_Title As String
Private m_Subtitle As String
Private m_WelcomeMessage As String

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

Public Property Get WelcomeMessage() As String
    WelcomeMessage = m_WelcomeMessage
End Property

Public Property Let WelcomeMessage(ByVal Value As String)
    If m_WelcomeMessage = Value Then Exit Property
    m_WelcomeMessage = Value
    m_Event.OnPropertyChanged "WelcomeMessage"
End Property

Public Property Get SalesCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "SalesMethod")
    Set SalesCommand = Cmd
End Property

Public Property Get LogoutCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "LogoutMethod")
    Set LogoutCommand = Cmd
End Property

Public Sub SalesMethod()
    Debug.Print "DeNovoSmoke MainMenu Sales (stub - milestone 3)"
End Sub

Public Sub LogoutMethod()
    On Error Resume Next
    If Not modApp.Shell Is Nothing Then modApp.Shell.ShowLogin
End Sub

Public Sub ApplyLoggedInUser(ByVal UserName As String)
    If Len(UserName) = 0 Then UserName = "Harness User"
    WelcomeMessage = "Welcome, " & UserName
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    m_Title = "D e N o v o"
    m_Subtitle = "Main Menu"
    m_WelcomeMessage = "Welcome"
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If Not m_Event Is Nothing Then m_Event.Unregister
End Sub
