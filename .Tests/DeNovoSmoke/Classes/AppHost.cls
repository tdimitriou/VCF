VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AppHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IApplication

Private WithEvents m_Base As VCF.Application
Attribute m_Base.VB_VarHelpID = -1

Public Property Get Base() As VCF.Application
    Set Base = m_Base
End Property

Public Sub InitializeApplication()
    Set modApp.XamlResources = modXamlResources.LoadXamlFolder
    m_Base.Resources.Add "XAML", modApp.XamlResources
    
    VCF.SetCustomConstructor New ObjectConstructor
    
    With New XAMLReader
        If modApp.XamlResources.ContainsKey("MyApp") Then
            Call .LoadApp(Me, AdaptMyAppXml(modApp.XamlResources("MyApp")))
        Else
            Err.Raise vbObjectError, , "MyApp.xml missing from Resources\XAML"
        End If
    End With
    
    m_Base.Resources.Add "LanguageManager", modApp.LanguageManager
    m_Base.Resources.Add "AppCommands", modApp.AppCommands
    m_Base.Resources.Add "AppProperties", modApp.AppProperties
End Sub

Public Property Get Resources() As ResourceDictionary
    Set Resources = m_Base.Resources
End Property

Public Sub Run(Optional Window As Object)
    m_Base.Run Window
End Sub

Private Property Get IApplication_Base() As VCF.Application
    Set IApplication_Base = Base
End Property

Private Function IApplication_FindResource(ByVal Key As String) As Variant
    API.CopyVariable m_Base.FindResource(Key), IApplication_FindResource
End Function

Private Sub IApplication_InitializeComponent()
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
    m_Base.StartupURI = RHS
End Property

Private Property Get IApplication_StartupURI() As String
    IApplication_StartupURI = m_Base.StartupURI
End Property

Private Function IApplication_TryFindResource(ByVal Key As String) As Variant
    API.CopyVariable m_Base.TryFindResource(Key), IApplication_TryFindResource
End Function

Private Sub Class_Initialize()
    VCF.Application.Create Me
End Sub

Private Function AdaptMyAppXml(ByVal Xml As String) As String
    AdaptMyAppXml = Replace(Xml, "x:Class=""MyApp""", "x:Class=""AppHost""", , , vbTextCompare)
End Function

Private Sub Class_Terminate()
    On Error Resume Next
    VCF.SetCustomConstructor Nothing
    VCF.ClearApplication
End Sub
