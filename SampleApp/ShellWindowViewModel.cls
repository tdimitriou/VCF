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

Implements INotifyPropertyChanged

Private m_Event As PropertyChangedEvent
Private m_InfoMessage As String
Private WithEvents Timer As cTimer
Attribute Timer.VB_VarHelpID = -1

Public Property Get InfoMessage() As String
    InfoMessage = m_InfoMessage
End Property

Public Property Let InfoMessage(ByVal Value As String)
    If m_InfoMessage = Value Then Exit Property
    
    m_InfoMessage = Value
    OnPropertyChanged "InfoMessage"
End Property

Public Property Get ButtonClick() As ICommand
    Dim Cmd As ActionCommand
    Set Cmd = New ActionCommand
    
    Cmd.Initialize VCF.NewFunction(Me, "ButtonClickMethod")
    
    Set ButtonClick = Cmd
    
End Property

Public Sub ButtonClickMethod()
    
    Timer.Enabled = False
    Me.InfoMessage = "You Clicked Me!"
    Timer.Enabled = True

End Sub

Public Property Get GridButtonClick() As ICommand
    Dim Cmd As ActionCommand
    Set Cmd = New ActionCommand
    
    Cmd.Initialize VCF.NewFunction(Me, "GridButtonClickMethod")
    
    Set GridButtonClick = Cmd
    
End Property

Public Sub GridButtonClickMethod(Optional Parameter)
    
    Dim Sender As VCF.Button
    Set Sender = Parameter
    Dim Label As TextBlock
    Set Label = Sender.Children(0)
    Dim Text As String
    Text = Label.Text
    
    Timer.Enabled = False
    Me.InfoMessage = "You Clicked Button " & Text
    Timer.Enabled = True

End Sub

Private Sub Class_Initialize()
    Set m_Event = New PropertyChangedEvent
    m_Event.Register Me
    Set Timer = New_c.Timer(3000, False)
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = m_Event
End Function

Private Sub OnPropertyChanged(ByVal PropertyName As String)
    m_Event.OnPropertyChanged PropertyName
End Sub

Private Sub Timer_Timer()
    Timer.Enabled = False
    InfoMessage = ""
End Sub
