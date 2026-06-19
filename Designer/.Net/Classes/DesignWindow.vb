Imports System
Imports System.Runtime.InteropServices

Public Class DesignWindow
    Implements VCF.IWindow

    Private m_Win As VCF.Window

    Public Sub New()
        ' Use Constructor to create Window
        Dim constructor As VCF.Constructor = New VCF.Constructor()
        m_Win = constructor.NewWindow(Me)
    End Sub

    Public ReadOnly Property Base() As VCF.Window
        Get
            Return m_Win
        End Get
    End Property

    Public ReadOnly Property IWindow_Base() As VCF.Window Implements VCF.IWindow.Base
        Get
            Return m_Win
        End Get
    End Property

    Public Sub IWindow_InitializeComponent() Implements VCF.IWindow.InitializeComponent
        ' Design-time initialization
        m_Win.DesignWidth = 800
        m_Win.DesignHeight = 600
    End Sub
End Class

