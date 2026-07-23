Attribute VB_Name = "modHarnessConfig"
Option Explicit

' Migrated Grid login (no canvas scale). False = legacy absolute LoginView.xml.
Public Const USE_WPF_LOGIN_LAYOUT As Boolean = True

' Shift+B border test — log BorderStyle to Immediate window (modBorderChromeDiag).
Public Const ENABLE_BORDER_CHROME_DIAG As Boolean = True

' True = BorderStyle 2 (title bar + drag resize). False = borderless POS-like shell (use Shift+1/2/3 to resize).
Public Const USE_SIZABLE_SHELL_BORDER As Boolean = False

' False = production-like: Login XAML loads on first ShowLogin (Phase 7f). True = m1 eager preload (A/B only).
Public Const EAGER_LOGIN_LOAD As Boolean = False

' False = MainMenu loads on first ShowMainMenu (milestone 2). True = eager preload.
Public Const EAGER_MAINMENU_LOAD As Boolean = False

' False = SalesOrder loads on first ShowSalesOrder (milestone 3). True = eager preload.
Public Const EAGER_SALES_LOAD As Boolean = False

' Log [P7d-LOAD-*] ms to Immediate window for Splash/Login/MainMenu/Sales/Bordered loads.
Public Const ENABLE_LOAD_BENCH As Boolean = True

Public Function LoginViewResourceKey() As String
    If USE_WPF_LOGIN_LAYOUT Then
        LoginViewResourceKey = "Migrated\Login\LoginViewWpf"
    Else
        LoginViewResourceKey = "Screens\Login\LoginView"
    End If
End Function

Public Function SalesOrderViewResourceKey() As String
    SalesOrderViewResourceKey = "Migrated\SalesOrder\SalesOrderView"
End Function

Public Function MainMenuViewResourceKey() As String
    MainMenuViewResourceKey = "Migrated\MainMenu\MainMenuView"
End Function
