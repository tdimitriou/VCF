VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubAppCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Property Get KeyClick() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "KeyClickMethod")
    Set KeyClick = Cmd
End Property

Public Property Get ClockInOutCommand() As ICommand
    Set ClockInOutCommand = NoOpCommand()
End Property

Public Property Get RestartCommand() As ICommand
    Set RestartCommand = NoOpCommand()
End Property

Public Property Get ShutdownCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "ShutdownMethod")
    Set ShutdownCommand = Cmd
End Property

Public Sub KeyClickMethod(Parameter)
    On Error Resume Next
    Dim VM As StubLoginViewModel
    Set VM = modApp.ActiveLoginViewModel
    If VM Is Nothing Then Exit Sub
    VM.PasswordText = VM.PasswordText & CStr(Parameter)
End Sub

Public Sub ShutdownMethod()
    On Error Resume Next
    modHarnessAppManager.Shutdown
End Sub

Private Function NoOpCommand() As ICommand
    Dim Cmd As HarnessActionCommand
    Set Cmd = New HarnessActionCommand
    Cmd.Initialize VCF.NewFunction(Me, "NoOpMethod")
    Set NoOpCommand = Cmd
End Function

Public Sub NoOpMethod()
    Debug.Print "DeNovoSmoke command (no-op)"
End Sub
