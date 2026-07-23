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
Private m_MainMenu As HarnessScreen
Private m_Sales As HarnessScreen
Private m_Nav As HarnessNavigation
Private WithEvents m_SplashTimer As cTimer
Attribute m_SplashTimer.VB_VarHelpID = -1

Public Sub Show()
    If m_Splash Is Nothing Then InitializeViews
    If Win.Children.Count = 0 Then ShowSplash
    Win.Show
    Win.RelayoutChildren
    Win.Form.WidgetRoot.Refresh
End Sub

Public Sub InitializeViews()
    Set m_Nav = New HarnessNavigation
    
    Set m_Splash = New HarnessScreen
    m_Splash.Configure "Screens\Splash\SplashView", False
    modHarnessLoadBench.TimedLoadView m_Splash, "P7d-LOAD-SPLASH"
    
    Set m_Login = New HarnessScreen
    m_Login.Configure LoginViewResourceKey(), True
    If modHarnessConfig.EAGER_LOGIN_LOAD Then
        modHarnessLoadBench.TimedLoadView m_Login, "P7d-LOAD-LOGIN"
    End If
    
    Set m_MainMenu = New HarnessScreen
    m_MainMenu.Configure "Screens\MainMenu\MainMenuView", False
    If modHarnessConfig.EAGER_MAINMENU_LOAD Then
        modHarnessLoadBench.TimedLoadView m_MainMenu, "P7d-LOAD-MAINMENU"
    End If
    
    Set m_Sales = New HarnessScreen
    m_Sales.Configure "Screens\SalesOrder\SalesOrderView", False
    If modHarnessConfig.EAGER_SALES_LOAD Then
        modHarnessLoadBench.TimedLoadView m_Sales, "P7d-LOAD-SALES"
    End If
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
    EnsureLoginLoaded
    m_Nav.SetActiveView Win, m_Login
End Sub

Public Sub ShowMainMenu()
    On Error Resume Next
    Set m_SplashTimer = Nothing
    EnsureLoginLoaded
    EnsureMainMenuLoaded
    ApplyMainMenuUserContext
    m_Nav.SetActiveView Win, m_MainMenu
End Sub

Public Sub ShowSalesOrder()
    On Error Resume Next
    Set m_SplashTimer = Nothing
    EnsureMainMenuLoaded
    EnsureSalesLoaded
    m_Nav.SetActiveView Win, m_Sales
End Sub

Private Sub EnsureLoginLoaded()
    If m_Login Is Nothing Then
        Set m_Login = New HarnessScreen
        m_Login.Configure LoginViewResourceKey(), True
    End If
    If Not m_Login.IsViewLoaded Then
        modHarnessLoadBench.TimedLoadView m_Login, "P7d-LOAD-LOGIN"
    End If
End Sub

Private Sub EnsureMainMenuLoaded()
    If m_MainMenu Is Nothing Then
        Set m_MainMenu = New HarnessScreen
        m_MainMenu.Configure "Screens\MainMenu\MainMenuView", False
    End If
    If Not m_MainMenu.IsViewLoaded Then
        modHarnessLoadBench.TimedLoadView m_MainMenu, "P7d-LOAD-MAINMENU"
    End If
End Sub

Private Sub EnsureSalesLoaded()
    If m_Sales Is Nothing Then
        Set m_Sales = New HarnessScreen
        m_Sales.Configure "Screens\SalesOrder\SalesOrderView", False
    End If
    If Not m_Sales.IsViewLoaded Then
        modHarnessLoadBench.TimedLoadView m_Sales, "P7d-LOAD-SALES"
    End If
End Sub

Private Sub ApplyMainMenuUserContext()
    Dim Vm As StubMainMenuViewModel
    Dim UserName As String
    Dim Named As Object
    Dim LV As VCF.ListView
    
    On Error Resume Next
    If m_MainMenu Is Nothing Then Exit Sub
    If Not TypeOf m_MainMenu.DataContext Is StubMainMenuViewModel Then Exit Sub
    Set Vm = m_MainMenu.DataContext
    
    UserName = ""
    If Not m_Login Is Nothing Then
        Set Named = VCF.NamingManager.GetNamedChildren(m_Login)
        If Not Named Is Nothing Then
            If Named.ContainsKey("MyList") Then
                Set LV = Named.Item("MyList")
                If Not LV Is Nothing Then
                    If Not IsEmpty(LV.SelectedItem) Then UserName = CStr(LV.SelectedItem)
                End If
            End If
        End If
    End If
    If Len(UserName) = 0 And Not modApp.ActiveLoginViewModel Is Nothing Then
        If Not modApp.ActiveLoginViewModel.UserList Is Nothing Then
            If modApp.ActiveLoginViewModel.UserList.Count > 0 Then
                UserName = CStr(modApp.ActiveLoginViewModel.UserList.Item(0))
            End If
        End If
    End If
    If Len(UserName) = 0 And Not modApp.AppProperties Is Nothing Then
        UserName = modApp.AppProperties.CurrentUser
    End If
    Vm.ApplyLoggedInUser UserName
End Sub

Public Sub StopTimers()
    On Error Resume Next
    Set m_SplashTimer = Nothing
End Sub

Private Sub Class_Initialize()
    NewWindow Me, Win
    Win.Width = 1024
    Win.Height = 768
    Set Form = Win.Form
    Form.Width = 1024
    Form.Height = 768
End Sub

Private Sub IWindow_InitializeComponent()
    If modHarnessConfig.USE_SIZABLE_SHELL_BORDER Then
        Win.DependencyProperties.SetValue "BorderStyle", 2
    Else
        Win.DependencyProperties.SetValue "BorderStyle", 0
    End If
End Sub

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = Win
End Property

Private Sub m_SplashTimer_Timer()
    ShowLogin
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If (Shift And vbShiftMask) = 0 Then Exit Sub
    Select Case KeyCode
        Case vbKeyL: ShowLogin
        Case vbKeyM: ShowMainMenu
        Case vbKeyS: ShowSalesOrder
        Case vbKeyB: modBorderTest.ShowBorderTestWindow
        Case vbKey1: modHarnessResize.ApplyPreset1024 Win, Form
        Case vbKey2: modHarnessResize.ApplyPreset800 Win, Form
        Case vbKey3: modHarnessResize.ApplyPresetWidescreen Win, Form
    End Select
End Sub

Private Sub Form_QueryUnload(Cancel As Boolean, UnloadMode As QueryUnloadConstants)
    ' RemoveAll during Exit may raise QueryUnload — guard avoids re-entrant Shutdown.
    If UnloadMode = vbFormControlMenu Then modHarnessAppManager.Shutdown
End Sub
