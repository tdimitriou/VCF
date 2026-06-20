VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IApplication

Private WithEvents m_Base As VCF.Application
Attribute m_Base.VB_VarHelpID = -1

'#Region IApplication Inheritance
    Public Property Get Base() As VCF.Application
        Set Base = m_Base
    End Property
    
    Public Sub InitializeComponent()
        With New XAMLReader
            Call .LoadApp(Me, New_c.FSO.ReadByteContent(App.Path & "\Resources\XAML\MyApp.xml"))
        End With
    End Sub

    Public Property Get Resources() As ResourceDictionary
        Set Resources = m_Base.Resources
    End Property
    
    Public Sub Run(Optional Window As Object)
        m_Base.Run Window
    End Sub

    Public Property Let StartupURI(ByVal Value As String)
        m_Base.StartupURI = Value
    End Property
    
    Public Property Get StartupURI() As String
        StartupURI = m_Base.StartupURI
    End Property
'#End Region

'#Region IApplication Interface Implementation
    Private Property Get IApplication_Base() As VCF.Application
        Set IApplication_Base = Me.Base
    End Property

    Private Function IApplication_FindResource(ByVal Key As String) As Variant
        API.CopyVariable m_Base.FindResource(Key), IApplication_FindResource
    End Function

    Private Sub IApplication_InitializeComponent()
        InitializeComponent
    End Sub
    
    Private Property Get IApplication_Resources() As ResourceDictionary
        Set IApplication_Resources = Resources
    End Property
    
    Private Sub IApplication_Run(Optional Window As Object)
        Run Window
    End Sub

    Private Sub IApplication_SetBase(ByVal Baseclass As VCF.Application)
        Set m_Base = Baseclass
    End Sub

    Private Property Let IApplication_StartupURI(ByVal RHS As String)
        StartupURI = RHS
    End Property

    Private Property Get IApplication_StartupURI() As String
        IApplication_StartupURI = StartupURI
    End Property
    
    Private Function IApplication_TryFindResource(ByVal Key As String) As Variant
        API.CopyVariable m_Base.TryFindResource(Key), IApplication_TryFindResource
    End Function
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        VCF.Application.Create Me
    End Sub
'#End Region

Private Sub m_Base_Startup()
    '
End Sub

