VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Col As cSortedDictionary

Public Function AddStyleDictionary(ByVal TargetType As String) As cSortedDictionary
    If Len(TargetType) = 0 Then
        Err.Raise vbObjectError + 10002, "No Target name defined"
        Exit Function
    End If
    
    If InStr(1, TargetType, ".") = 0 Then TargetType = "VCF." & TargetType
    
    If m_Col.Exists(TargetType) Then
        Err.Raise vbObjectError + 10001, "Target Type already exists"
        Exit Function
    End If
    
    Dim NewTarget As cSortedDictionary
    Set NewTarget = New_c.SortedDictionary
    
    m_Col.Add TargetType, NewTarget
    
    Set AddStyleDictionary = NewTarget
End Function

Public Property Get StyleDictionary(ByVal TargetType As String) As cSortedDictionary
Attribute StyleDictionary.VB_UserMemId = 0
    If Len(TargetType) = 0 Then
        Err.Raise vbObjectError + 10002, "No Target name defined"
        Exit Property
    End If
    
    If InStr(1, TargetType, ".") <= 0 Then TargetType = "VCF." & TargetType
    
    If Not m_Col.Exists(TargetType) Then
        Err.Raise vbObjectError + 10001, "Target Type not found"
        Exit Property
    End If

    Set StyleDictionary = m_Col(TargetType)
End Property

Public Property Set StyleDictionary(ByVal TargetType As String, ByVal Value As cSortedDictionary)
    If Len(TargetType) = 0 Then
        Err.Raise vbObjectError + 10002, "No Target name defined"
        Exit Property
    End If
    
    If InStr(1, TargetType, ".") <= 0 Then TargetType = "VCF." & TargetType
    
    If Not m_Col.Exists(TargetType) Then
        m_Col.Add TargetType, Value
        Exit Property
    End If

    Set m_Col(TargetType) = Value
End Property

Public Property Get Count() As Long
    Count = m_Col.Count
End Property

Public Function StyleDictionaryExists(ByVal TargetType As String) As Boolean
    If InStr(1, TargetType, ".") <= 0 Then TargetType = "VCF." & TargetType
    StyleDictionaryExists = m_Col.Exists(TargetType)
End Function

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = m_Col.Enumerator
End Function

Private Sub Class_Initialize()
    Set m_Col = New_c.SortedDictionary
End Sub
