VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "List"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements vbRichClient5.IEnumerable

Private m_List As cArrayList

'#Region Class Events
Private Sub Class_Initialize()
    Set m_List = New_c.ArrayList(vbVariant)
End Sub
'#End Region

'#Region Public Methods
Public Sub Add(Value)
    m_List.Add Value
End Sub

Public Sub AddRange(List)
    On Error Resume Next
    
    Dim Value
        
    If Not IsObject(List) Then
        m_List.AddElements List
        Exit Sub
    End If
    
    If TypeOf List Is VCF.List Then
        m_List.AddElements List.ToArray
        Exit Sub
    End If
    
    If TypeOf List Is cArrayList Then
        Dim Data
        Call List.CopyToArray(Data, 0)
        m_List.AddElements Data
        Exit Sub
    End If
    
    For Each Value In List
        m_List.Add Value
    Next
End Sub

Public Sub Insert(ByVal Index As Long, Value)
    If Index < 0 Then
        Err.Raise 9
        Exit Sub
    End If
    
    If Index > m_List.Count Then
        Err.Raise 9
        Exit Sub
    End If
    
    m_List.Add Value, Index
End Sub

Public Sub InsertRange(ByVal Index As Long, List As Object)
    Dim Idx
    Dim Value
    
    If Index < 0 Then
        Err.Raise 9     'Subscript out of range
        Exit Sub
    End If
    
    If Index > m_List.Count Then
        Err.Raise 9     'Subscript out of range
        Exit Sub
    End If
        
    Idx = Index
    For Each Value In List
        m_List.Add Value, Idx
        Idx = Idx + 1
    Next
End Sub

Public Sub Remove(Value)
    Dim Index As Long
    
    If Not Exists(Value, Index) Then Exit Sub
    
    Call m_List.Remove(Index)
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    If Index < 0 Then
        Err.Raise 9 'Subscript out of range
        Exit Sub
    End If
    
    If Count < 0 Then
        Err.Raise 9 'Subscript out of range
        Exit Sub
    End If
    
    If Index + Count > m_List.Count Then
        Err.Raise 9 'Subscript out of range
        Exit Sub
    End If
        
    m_List.RemoveElements Index, Count
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    If Index < 0 Or Index > m_List.Count - 1 Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Sub
    End If
    
    Call m_List.Remove(Index)
End Sub

Public Function ToArray() As Variant
    Call m_List.CopyToArray(ToArray, 0)
End Function

Public Sub Clear()
    Set m_List = New_c.ArrayList(vbVariant)
End Sub

Public Function Contains(Value) As Boolean
    Contains = Exists(Value)
End Function

Public Function IndexOf(Value) As Long
    Call Exists(Value, IndexOf)
End Function

Public Function GetRange(ByVal Index As Long, ByVal Count As Long) As VCF.List
    If Index < 0 Then
        Err.Raise 9 'Subscript out of range
        Exit Function
    End If
    
    If Count < 0 Then
        Err.Raise 9 'Subscript out of range
        Exit Function
    End If
    
    If Index + Count > m_List.Count Then
        Err.Raise 9 'Subscript out of range
        Exit Function
    End If

    Set GetRange = New List
    
    Dim Idx As Long
    For Idx = Index To Index + Count - 1
        GetRange.Add m_List(Idx)
    Next
End Function

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = New_c.EnumerateOn(Me, m_List.Count, 0)
End Function
'#End Region

'#Region Public Properties
Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
    API.CopyVariable m_List(Index), Item
End Property

Public Property Let Item(ByVal Index As Long, Value As Variant)
        
    If Index < 0 Or Index > m_List.Count - 1 Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Property
    End If
    
    m_List(Index) = Value
End Property

Public Property Set Item(ByVal Index As Long, Value As Variant)
        
    If Index < 0 Or Index > m_List.Count - 1 Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Property
    End If
    
    Set m_List(Index) = Value
End Property
'#End Region

'#Region Private Members
Private Function Exists(Value, Optional ByRef Index As Long) As Boolean
    Dim i As Long
    
    Index = -1
    
    If IsObject(Value) Then
        For i = 0 To m_List.Count - 1
            If m_List(i) Is Value Then
                Index = i
                Exists = True
                Exit Function
            End If
        Next
    Else
        For i = 0 To m_List.Count - 1
            If m_List(i) = Value Then
                Index = i
                Exists = True
                Exit Function
            End If
        Next
    End If
End Function
'#End Region

'#Region IEnumerable Interface
Private Function IEnumerable_NextItem(Idx As Long) As Variant
    If IsObject(m_List(Idx)) Then
        Set IEnumerable_NextItem = m_List(Idx)
    Else
        IEnumerable_NextItem = m_List(Idx)
    End If
End Function
'#End Region

