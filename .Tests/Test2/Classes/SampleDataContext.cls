VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SampleDataContext"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent

Private m_Text As String
Private m_Info As String
Private m_Items As NestedDataContext
Private m_MyText As String

Private WithEvents Timer As cTimer
Attribute Timer.VB_VarHelpID = -1

'#Region Public Properties
    Public Property Get Items() As NestedDataContext
        Set Items = m_Items
    End Property
    
    Public Property Get MyList() As ObservableCollection
        Set MyList = New ObservableCollection
        
        Dim Index As Long
        For Index = 0 To 99
            MyList.Add "Item " & Index
        Next
        
    End Property
    Public Property Get MyText() As String
        MyText = m_MyText
    End Property
    
    Public Property Let MyText(ByVal Value As String)
        If m_MyText = Value Then Exit Property
            
        m_MyText = Value
        
        Call m_Event.OnPropertyChanged("MyText")
    End Property
    
    Public Property Get Text() As String
        Text = m_Text
    End Property
    
    Public Property Let Text(ByVal Value As String)
        If m_Text = Value Then Exit Property
            
        m_Text = Value
        
        Call m_Event.OnPropertyChanged("Text")
    End Property
    
    Public Property Get Info() As String
        Info = m_Info
    End Property
    
    Public Property Let Info(ByVal Value As String)
        If m_Info = Value Then Exit Property
            
        m_Info = Value
        
        Call m_Event.OnPropertyChanged("Info")
    End Property
    
    Public Property Get MyCommand() As ICommand
        Set MyCommand = New ShellExecCommand
    End Property
    
    Public Property Get SwapContent() As ICommand
        Dim MainWin As Object
        If Application.Current.Base.Windows.Count = 0 Then Exit Property
        Set MainWin = API.CUnk(Application.Current.Base.Windows.ItemByIndex(0))
        If Not TypeOf MainWin Is ShellWindow Then Exit Property
                
        Set SwapContent = modActionCommandCctor.NewActionCommand(VCF.NewFunction(MainWin, "SwapContent", VbMethod), True)
    End Property
'#End Region

'#Region INotifyPropertyChanged
    Private Function INotifyPropertyChanged_PropertyChangedEvent() As VCF.PropertyChangedEvent
        Set INotifyPropertyChanged_PropertyChangedEvent = PropertyChangedEvent
    End Function
    
    Public Function PropertyChangedEvent() As VCF.PropertyChangedEvent
        Set PropertyChangedEvent = m_Event
    End Function
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Items = New NestedDataContext
        Set m_Items.Child = New NestedDataContext
        Set m_Event = New PropertyChangedEvent
        m_Event.Register Me
        
        MyText = "Hello!!!"
                
        Set Timer = New_c.Timer(1000, True)
    End Sub
    
    Private Sub Class_Terminate()
        On Error Resume Next
        m_Event.Unregister
    End Sub
'#End Region

'#Region Timer Events
    Private Sub Timer_Timer()
        If Not IsNumeric(m_Text) Then m_Text = 1
        Me.Text = m_Text * 1.005
        
        m_Items.Child.List.Add "Hello"
        If m_Items.Child.Count = 10 Then
            Set m_Items = New NestedDataContext
            Set m_Items.Child = New NestedDataContext
            m_Event.OnPropertyChanged "Items"
        End If
        
    End Sub
'#End Region

