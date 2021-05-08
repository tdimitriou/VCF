VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObservableDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)

Private m_List As cSortedDictionary

'#Region Class Events
Private Sub Class_Initialize()
    Set m_List = New_c.SortedDictionary(TextCompare, True)
End Sub
'#End Region

'#Region Public Methods
Public Sub Add(Key, Value)
    m_List.Add Key, Value
    
    Call OnCollectionChanged(CollectionChangedActionAdd, NewList(Value), m_List.Count - 1, Nothing, -1)
End Sub

Public Sub Remove(Value)
    Dim Index As Long
        
    If Not Exists(Value, Index) Then Exit Sub
    
    Call m_List.RemoveByIndex(Index)
    
    Call OnCollectionChanged(CollectionChangedActionRemove, Nothing, -1, NewList(Value), Index)
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    Dim OldValue As Variant
        
    If Index < 0 Or Index > m_List.Count - 1 Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Sub
    End If
        
    If IsObject(m_List(Index)) Then
        Set OldValue = m_List.ItemByIndex(Index)
    Else
        OldValue = m_List.ItemByIndex(Index)
    End If
    
    Call m_List.RemoveByIndex(Index)
    
    Call OnCollectionChanged(CollectionChangedActionRemove, Nothing, -1, NewList(OldValue), Index)
End Sub

Public Sub RemoveKey(ByVal Key)
    Dim OldValue As Variant
    Dim Index As Long
        
    If Not Exists(Key) Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
    
        Exit Sub
    End If
    
    Index = m_List.IndexByKey(Key)
    
    If IsObject(m_List(Key)) Then
        Set OldValue = m_List(Key)
    Else
        OldValue = m_List(Index)
    End If
    
    Call m_List.Remove(Key)
    
    Call OnCollectionChanged(CollectionChangedActionRemove, Nothing, -1, NewList(OldValue), Index)
End Sub

Public Sub Clear()
    Dim OldList As List
    Dim Idx As Long
    
    Set OldList = New List
    
    For Idx = 0 To m_List.Count - 1
        OldList.Add m_List.ItemByIndex(Idx)
    Next
    
    Set m_List = New_c.SortedDictionary(TextCompare)
    
    Call OnCollectionChanged(CollectionChangedActionReset, Nothing, -1, OldList, 0)
End Sub

Public Function Contains(Value) As Boolean
    Contains = Exists(Value)
End Function

Public Function ContainsKey(Key) As Boolean
    ContainsKey = m_List.Exists(Key)
End Function

Public Function IndexOf(Value) As Long
    Call Exists(Value, IndexOf)
End Function

Public Function IndexOfKey(Key) As Long
    IndexOfKey = m_List.IndexByKey(Key)
End Function

Public Function KeyOfIndex(ByVal Index As Long)
    API.CopyVariable m_List.KeyByIndex(Index), KeyOfIndex
End Function

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = m_List.Enumerator
End Function
'#End Region

'#Region Public Properties
Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get Item(Key) As Variant
Attribute Item.VB_UserMemId = 0
    If IsObject(m_List(Key)) Then
        Set Item = m_List(Key)
    Else
        Item = m_List(Key)
    End If
End Property

Public Property Set Item(Key, Value)
   
    If Not m_List.Exists(Key) Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Property
    End If
    
    Dim OldItems As List
    Dim Index As Long
    
    Index = m_List.IndexByKey(Key)
    
    If IsObject(Value) Then
        If m_List(Key) Is Value Then Exit Property
        Set OldItems = NewList(m_List(Key))
        Set m_List(Key) = Value
        Call OnCollectionChanged(CollectionChangedActionReplace, NewList(Value), Index, OldItems, Index)
    Else
        If m_List(Key) = Value Then Exit Property
        Set OldItems = NewList(m_List(Key))
        m_List(Key) = Value
        Call OnCollectionChanged(CollectionChangedActionReplace, NewList(Value), Index, OldItems, Index)
    End If
End Property

Public Property Get ItemAt(ByVal Index As Long) As Variant
    If IsObject(m_List.ItemByIndex(Index)) Then
        Set ItemAt = m_List.ItemByIndex(Index)
    Else
        ItemAt = m_List.ItemByIndex(Index)
    End If
End Property

Public Property Set ItemAt(ByVal Index As Long, Value As Variant)
        
    If Index < 0 Or Index > m_List.Count - 1 Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Property
    End If
    
    Dim OldItems As List
    
    If IsObject(Value) Then
        If m_List(Index) Is Value Then Exit Property
        Set OldItems = NewList(m_List.ItemByIndex(Index))
        Set m_List.ItemByIndex(Index) = Value
        Call OnCollectionChanged(CollectionChangedActionReplace, NewList(Value), Index, OldItems, Index)
    Else
        If m_List.ItemByIndex(Index) = Value Then Exit Property
        Set OldItems = NewList(m_List.ItemByIndex(Index))
        m_List.ItemByIndex(Index) = Value
        Call OnCollectionChanged(CollectionChangedActionReplace, NewList(Value), Index, OldItems, Index)
    End If
End Property
'#End Region

'#Region Private Members
Private Function Exists(Value, Optional ByRef Index As Long) As Boolean
    Exists = m_List.ItemExists(Value, Index)
End Function
'#End Region

'#Region Friend Members
'CSEH: ErrMsgBox
Friend Sub OnCollectionChanged(ByVal Action As CollectionChangedAction, _
                                                ByVal NewItems As List, _
                                                ByVal NewStartingIndex As Long, _
                                                ByVal OldItems As List, _
                                                ByVal OldStartingIndex As Long)
        '<EhHeader>
        On Error GoTo OnCollectionChanged_Err
        '</EhHeader>
    
100     RaiseEvent CollectionChanged(Me, NewCollectionChangedEventArgs(Action, NewItems, NewStartingIndex, OldItems, OldStartingIndex))

        '<EhFooter>
        Exit Sub

OnCollectionChanged_Err:
        MsgBox Err.Description & vbCrLf & _
               "in RC5App.ObservableCollection.OnCollectionChanged " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub
'#End Region
