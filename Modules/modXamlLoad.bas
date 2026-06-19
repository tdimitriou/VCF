Attribute VB_Name = "modXamlLoad"
Option Explicit

' vbObjectError Or &H2000 — XAML load failures (catch via Err.Number)
Public Const vbXamlLoadError As Long = vbObjectError Or &H2000&

Public Sub RaiseXamlLoad( _
    ByVal Message As String, _
    Optional ByVal ElementName As String = vbNullString, _
    Optional ByVal PropertyName As String = vbNullString, _
    Optional ByVal Line As Long = 0, _
    Optional ByVal Column As Long = 0, _
    Optional ByVal InnerCode As Long = 0)

    Dim Ex As XamlLoadException
    Set Ex = New XamlLoadException
    Ex.Initialize Message, ElementName, PropertyName, Line, Column, InnerCode
    Ex.Raise
End Sub
