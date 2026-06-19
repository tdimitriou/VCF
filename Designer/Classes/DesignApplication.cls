VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignApplication"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VCF.IApplication

Private m_Base As VCF.Application

Private Sub Class_Initialize()
    On Error Resume Next
    
    ' Create Application base
    Set m_Base = New VCF.Application
    
    ' Set as current application - Application refers to StaticClasses.Application
    ' which is the ApplicationStatic instance from modStaticClasses
    ' We need to set it via VCF.StaticClasses.Application.Current
    Dim AppStatic As VCF.ApplicationStatic
    Set AppStatic = VCF.StaticClasses.Application
    Set AppStatic.Current = Me
    
    If Err Then Err.Clear
End Sub

Private Property Get IApplication_Base() As VCF.Application
    Set IApplication_Base = m_Base
End Property

Private Function IApplication_FindResource(ByVal Key As String) As Variant
    On Error Resume Next
    VCF.API.CopyVariable m_Base.FindResource(Key), IApplication_FindResource
    If Err Then Err.Clear
End Function

Private Sub IApplication_InitializeComponent()
    ' Design-time initialization - minimal
    On Error Resume Next
    If Err Then Err.Clear
End Sub

Private Property Get IApplication_Resources() As VCF.ObservableDictionary
    Set IApplication_Resources = m_Base.Resources
End Property

Private Sub IApplication_Run(Optional Window As Object)
    ' Don't run in design mode
    On Error Resume Next
    If Err Then Err.Clear
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
    On Error Resume Next
    VCF.API.CopyVariable m_Base.TryFindResource(Key), IApplication_TryFindResource
    If Err Then Err.Clear
End Function

Public Property Get Base() As VCF.Application
    Set Base = m_Base
End Property

