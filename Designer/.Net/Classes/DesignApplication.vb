Imports System
Imports System.Runtime.InteropServices

Public Class DesignApplication
    Implements VCF.IApplication

    Private m_Base As VCF.Application

    Public Sub New()
        ' Create Application base
        m_Base = New VCF.Application()

        ' Set as current application
        Dim appStatic As VCF.ApplicationStatic = VCF.StaticClasses.Application
        appStatic.Current = Me
    End Sub

    Public ReadOnly Property Base() As VCF.Application
        Get
            Return m_Base
        End Get
    End Property

    Public ReadOnly Property IApplication_Base() As VCF.Application Implements VCF.IApplication.Base
        Get
            Return m_Base
        End Get
    End Property

    Public Function IApplication_FindResource(ByVal Key As String) As Object Implements VCF.IApplication.FindResource
        Return m_Base.FindResource(Key)
    End Function

    Public Sub IApplication_InitializeComponent() Implements VCF.IApplication.InitializeComponent
        ' Design-time initialization - minimal
    End Sub

    Public ReadOnly Property IApplication_Resources() As VCF.ObservableDictionary Implements VCF.IApplication.Resources
        Get
            Return m_Base.Resources
        End Get
    End Property

    Public Sub IApplication_Run(Optional ByVal Window As Object = Nothing) Implements VCF.IApplication.Run
        ' Don't run in design mode
    End Sub

    Public Sub IApplication_SetBase(ByVal Baseclass As VCF.Application) Implements VCF.IApplication.SetBase
        m_Base = Baseclass
    End Sub

    Public Property IApplication_StartupURI() As String Implements VCF.IApplication.StartupURI
        Get
            Return m_Base.StartupURI
        End Get
        Set(ByVal value As String)
            m_Base.StartupURI = value
        End Set
    End Property

    Public Function IApplication_TryFindResource(ByVal Key As String) As Object Implements VCF.IApplication.TryFindResource
        Return m_Base.TryFindResource(Key)
    End Function
End Class

