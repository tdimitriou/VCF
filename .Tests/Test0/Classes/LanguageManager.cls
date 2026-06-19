VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LanguageManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMarkupExtension

Public Name As String
Public Lastname As String

Private Sub Class_Initialize()
    Name = ""
    Lastname = ""
End Sub

Private Function IMarkupExtension_ProvideValue() As Variant
    IMarkupExtension_ProvideValue = ProvideValue
End Function

Public Function ProvideValue() As Variant
    ProvideValue = "Hello " & Name & " " & Lastname & "!"
End Function
