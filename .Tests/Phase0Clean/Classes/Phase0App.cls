VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IApplication

Private m_Base As VCF.Application

Public Property Get Base() As VCF.Application
    Set Base = m_Base
End Property

Private Property Get IApplication_Base() As VCF.Application
    Set IApplication_Base = Base
End Property

Private Function IApplication_FindResource(ByVal Key As String) As Variant
    If Not m_Base Is Nothing Then API.CopyVariable m_Base.FindResource(Key), IApplication_FindResource
End Function

Private Sub IApplication_InitializeComponent()
End Sub

Private Property Get IApplication_Resources() As ResourceDictionary
    If Not m_Base Is Nothing Then Set IApplication_Resources = m_Base.Resources
End Property

Private Sub IApplication_Run(Optional Window As Object)
End Sub

Private Sub IApplication_SetBase(ByVal Baseclass As VCF.Application)
    Set m_Base = Baseclass
End Sub

Private Property Let IApplication_StartupURI(ByVal RHS As String)
    If Not m_Base Is Nothing Then m_Base.StartupURI = RHS
End Property

Private Property Get IApplication_StartupURI() As String
    If Not m_Base Is Nothing Then IApplication_StartupURI = m_Base.StartupURI
End Property

Private Function IApplication_TryFindResource(ByVal Key As String) As Variant
    If Not m_Base Is Nothing Then API.CopyVariable m_Base.TryFindResource(Key), IApplication_TryFindResource
End Function

Private Sub Class_Initialize()
    VCF.Application.Create Me
End Sub

' No ClearApplication here - suite calls VCF.ClearApplication after Window.Unload.