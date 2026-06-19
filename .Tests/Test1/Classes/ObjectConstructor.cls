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
        Case "SampleDataContext":       Set CreateInstance = New SampleDataContext
        Case "TestValueConverter":      Set CreateInstance = New TestValueConverter
        Case "TestCommandObj":          Set CreateInstance = New TestCommandObj
        Case "ShellExecCommand":        Set CreateInstance = New ShellExecCommand
        Case "LanguageManager":         Set CreateInstance = New LanguageManager
        Case "Numpad":                  Set CreateInstance = New Numpad
        Case "ShellWindow":             Set CreateInstance = New ShellWindow
        Case "LanguageString":          Set CreateInstance = New LanguageString
        
    End Select

End Function

Private Function IObjectConstructor_CreateInstance(ByVal Classname As String) As Object
    Set IObjectConstructor_CreateInstance = CreateInstance(Classname)
End Function
