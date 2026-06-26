VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ShellWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IWindow

Private Win As Window
Attribute Win.VB_VarHelpID = -1
Private WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1

Private m_Splash As HarnessScreen
Private m_Login As HarnessScreen
Private m_Nav As HarnessNavigation
Private WithEvents m_SplashTimer As cTimer
Attribute m_SplashTimer.VB_VarHelpID = -1

Public Sub Show()
    If m_Splash Is Nothing Then InitializeViews
    If Win.Children.Count = 0 Then ShowSplash
    Win.Show
End Sub

Public Sub InitializeViews()
    Set m_Nav = New HarnessNavigation
    
    Set m_Splash = New HarnessScreen
    m_Splash.Configure "Screens\Splash\SplashView", False
    m_Splash.LoadView
    
    Set m_Login = New HarnessScreen
    m_Login.Configure "Screens\Login\LoginView", True
    m_Login.LoadView
End Sub

Public Sub ShowSplash()
    Win.Children.Add m_Splash
    m_Splash.Visibility = VisibilityVisible
    Win.RelayoutChildren
    m_Splash.OnActivated
    
    Set m_SplashTimer = New_c.Timer(2500, True)
End Sub

Public Sub ShowLogin()
    On Error Resume Next
    Set m_SplashTimer = Nothing
    m_Nav.SetActiveView Win, m_Login
End Sub

Private Sub Class_Initialize()
    NewWindow Me, Win
    Win.DesignWidth = 1024
    Win.DesignHeight = 768
    Win.DependencyProperties.SetValue "BorderStyle", 0
    Set Form = Win.Form
    Form.Width = 1024
    Form.Height = 768
End Sub

Private Sub IWindow_InitializeComponent()
End Sub

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub m_SplashTimer_Timer()
    ShowLogin
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyL And (Shift And vbShiftMask) <> 0 Then ShowLogin
End Sub
