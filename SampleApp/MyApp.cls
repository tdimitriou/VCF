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

Private WithEvents Base As VCF.Application
Attribute Base.VB_VarHelpID = -1

Private Sub Base_Startup()
    '
End Sub

Private Sub Class_Initialize()
    VCF.Application.Create Me
End Sub

Private Property Get IApplication_Base() As VCF.Application
    Set IApplication_Base = Base
End Property

Private Function IApplication_FindResource(ByVal Key As String) As Variant
    API.CopyVariable Base.FindResource(Key), IApplication_FindResource
End Function

Private Sub IApplication_InitializeComponent()
    '
End Sub

Private Property Get IApplication_Resources() As VCF.ObservableDictionary
    Set IApplication_Resources = Base.Resources
End Property

Private Sub IApplication_Run(Optional Window As Object)
    Base.Run Window
End Sub

Private Sub IApplication_SetBase(ByVal Baseclass As VCF.Application)
    Set Base = Baseclass
End Sub

Private Property Let IApplication_StartupURI(ByVal RHS As String)
    Base.StartupURI = RHS
End Property

Private Property Get IApplication_StartupURI() As String
    IApplication_StartupURI = Base.StartupURI
End Property

Private Function IApplication_TryFindResource(ByVal Key As String) As Variant
    API.CopyVariable Base.TryFindResource(Key), IApplication_TryFindResource
End Function
