VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLStyleReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub LoadStyle(ByVal ResourceDictionary As ObservableDictionary, Node As cElement)
    If Node.BaseName <> "Style" Then Exit Sub
    
    Dim Style As Style
    Set Style = GetStyle(ResourceDictionary)
    
    If Style Is Nothing Then
        Set Style = New Style
        ResourceDictionary.Add "Style", Style
    End If
    
    
End Sub

Private Function GetStyle(ByVal ResourceDictionary As ObservableDictionary) As Style
    On Error Resume Next
    
    Static Style As Style
    Static IsSet As Boolean
    
    If Not IsSet Then Set Style = ResourceDictionary("Style")

    Set GetStyle = Style
    
End Function

Private Function GetTargetType(Attributes As cAttributes) As String
    On Error Resume Next
    
    GetTargetType = Attributes("TargetType").Text
    
    If Len(GetTargetType) = 0 Then Exit Function
    If InStr(1, GetTargetType, ".") <= 0 Then GetTargetType = "VCF." & GetTargetType
End Function
