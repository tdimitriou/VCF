VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectConstructor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IObjectConstructor

Public Function CreateInstance(ByVal Classname As String) As Object
    Select Case Classname
        Case "SplashView"
            Dim Splash As HarnessScreen
            Set Splash = New HarnessScreen
            Splash.Configure "Screens\Splash\SplashView", False
            Set CreateInstance = Splash
        Case "LoginView"
            Dim Login As HarnessScreen
            Set Login = New HarnessScreen
            Login.Configure "Screens\Login\LoginView", True
            Set CreateInstance = Login
        Case "SplashViewModel"
            Set CreateInstance = New StubSplashViewModel
        Case "LoginViewModel"
            Set CreateInstance = New StubLoginViewModel
        Case "CaseConverter"
            Set CreateInstance = New StubCaseConverter
        Case Else
            Set CreateInstance = TryCreateResObject(Classname)
    End Select
End Function

Private Function TryCreateResObject(ByVal Classname As String) As Object
    On Error Resume Next
    
    Dim Key As String
    If LCase$(Left$(Classname, 4)) = "res." Then
        Key = Mid$(Classname, 5)
    ElseIf LCase$(Left$(Classname, 4)) = "res:" Then
        Key = Mid$(Classname, 5)
    Else
        Exit Function
    End If
    
    If Len(Key) = 0 Then Exit Function
    If Not modApp.XamlResources.ContainsKey(Key) Then Exit Function
    
    With New XAMLReader
        Set TryCreateResObject = .Load(modApp.XamlResources(Key))
    End With
End Function

Private Function IObjectConstructor_CreateInstance(ByVal Classname As String) As Object
    Set IObjectConstructor_CreateInstance = CreateInstance(Classname)
End Function
