VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageString"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements INotifyPropertyChanged

Private m_Event As VCF.PropertyChangedEvent

Private m_Text As String
Private Key As String
Private EN As cCollection
Private EL As cCollection

Private Col As cCollection

'#Region Public Properties
    Public Property Get Text() As String
Attribute Text.VB_UserMemId = 0
        Text = Col(Key)
    End Property
    
    Public Property Get Language() As String
        Language = IIf((Col Is EN), "EN", "EL")
    End Property
    
    Public Property Let Language(ByVal Value As String)
                    
            
        Set Col = IIf((Value = "EN"), EN, EL)
        
        Call m_Event.OnPropertyChanged("Language")
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
        Set EL = New_c.Collection
        EL.Add "Δώρος", "name"
        EL.Add "Δημητρίου", "lastname"
        
        Set EN = New_c.Collection
        EN.Add "Doros", "name"
        EN.Add "Dimitriou", "lastname"
        
        Set Col = EN
        
        Set m_Event = New PropertyChangedEvent
        m_Event.Register Me
        
    End Sub
    
    Private Sub Class_Terminate()
        On Error Resume Next
        m_Event.Unregister
    End Sub
'#End Region
