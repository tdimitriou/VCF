Attribute VB_Name = "modHarnessConfig"
Option Explicit

' WPF Grid login (LoginViewWpf.xml) — resize-friendly inner panel; legacy LoginView.xml when False.
Public Const USE_WPF_LOGIN_LAYOUT As Boolean = True

Public Function LoginViewResourceKey() As String
    If USE_WPF_LOGIN_LAYOUT Then
        LoginViewResourceKey = "Screens\Login\LoginViewWpf"
    Else
        LoginViewResourceKey = "Screens\Login\LoginView"
    End If
End Function
