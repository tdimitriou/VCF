VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Parse(ByVal strText As String) As String
    Parse = Replace(strText, "\n", vbCrLf)
End Function

Public Function Format(ByVal FormatString As String, ParamArray Args() As Variant) As String
   On Error Resume Next
   Dim i As Integer
    
    If IsMissing(Args) Then
        Format = FormatString
        Exit Function
    End If
        
    For i = 0 To UBound(Args)
        FormatString = Replace$(FormatString, "{" & i & "}", Args(i))
    Next
    Format = FormatString
End Function
