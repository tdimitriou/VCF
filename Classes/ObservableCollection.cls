VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObservableCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements vbRichClient5.IEnumerable
Implements INotifyCollectionChanged

Public Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)

Private m_List              As cArrayList
Private m_Ptr               As Long

Private m_CollectionChangedEvent As CollectionChangedEvent

'#Region Class Events
Private Sub Class_Initialize()
    Set m_List = New_c.ArrayList(vbVariant)
    Set m_CollectionChangedEvent = New CollectionChangedEvent
    m_CollectionChangedEvent.Register Me
    m_Ptr = ObjPtr(Me)
End Sub

Private Sub Class_Terminate()
    modStaticClasses.CollectionViewSource.DestroyDefaultView m_Ptr
    m_CollectionChangedEvent.Unregister
End Sub
'#End Region

'#Region Public Methods
Public Sub Add(Value)
    m_List.Add Value
    
    Call OnCollectionChanged(CollectionChangedActionAdd, NewList(Value), m_List.Count - 1, Nothing, -1)
End Sub

Public Sub AddRange(List)
    Dim Value
    Dim NewStartingIndex As Long
    
    NewStartingIndex = m_List.Count '- 1

    If Not IsObject(List) Then
        
        m_List.AddElements List
    
    ElseIf TypeOf List Is VCF.List Then
        
        m_List.AddElements List.ToArray
    
    ElseIf TypeOf List Is cArrayList Then
        
        Dim Data
        Call List.CopyToArray(Data, 0)
        m_List.AddElements Data
    
    Else
        
        For Each Value In List
            m_List.Add Value
        Next
    
    End If
    

'================================================
    
    Dim NewItems As VCF.List
    
    If Not IsObject(List) Then
        
        Set NewItems = New VCF.List
        NewItems.AddRange List
    
    Else
        
        If TypeOf List Is VCF.List Then
            
            Set NewItems = List
        
        Else
            
            Set NewItems = New VCF.List
            NewItems.AddRange List
        
        End If
    
    End If
        
    Call OnCollectionChanged(CollectionChangedActionAdd, NewItems, NewStartingIndex, Nothing, -1)
End Sub

Public Sub Insert(Index, Value)
    If Index < 0 Then
        Err.Raise 9
        Exit Sub
    End If
    
    If Index > m_List.Count Then
        Err.Raise 9
        Exit Sub
    End If
    
    m_List.Add Value, Index
    
    Call OnCollectionChanged(CollectionChangedActionAdd, NewList(Value), Index, Nothing, -1)
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

    Dim NewItems As VCF.List
    If TypeOf List Is VCF.List Then
        Set NewItems = List
    Else
        Set NewItems = New VCF.List
        NewItems.AddRange List
    End If
    
    Call OnCollectionChanged(CollectionChangedActionAdd, NewItems, Index, Nothing, -1)
End Sub

Public Sub Remove(Value)
    Dim Index As Long
    
    If Not Exists(Value, Index) Then Exit Sub
    
    Call m_List.Remove(Index)
    
    Call OnCollectionChanged(CollectionChangedActionRemove, Nothing, -1, NewList(Value), Index)
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Dim Idx As Long
    Dim OldList As VCF.List
    
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
    
    Set OldList = New VCF.List
    For Idx = Index To Index + Count - 1
        OldList.Add m_List(Idx)
    Next
    
    m_List.RemoveElements Index, Count

    Call OnCollectionChanged(CollectionChangedActionRemove, Nothing, -1, OldList, Index)
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    Dim OldValue As Variant
        
    If Index < 0 Or Index > m_List.Count - 1 Then
        ' Raise "Subscript out of range" error
        Err.Raise 9
        Exit Sub
    End If
        
    If IsObject(m_List(Index)) Then
        Set OldValue = m_List(Index)
    Else
        OldValue = m_List(Index)
    End If
    
    Call m_List.Remove(Index)
    
    Call OnCollectionChanged(CollectionChangedActionRemove, Nothing, -1, NewList(OldValue), Index)
End Sub

Public Sub Clear()
    Dim OldList As List
    Dim Idx As Long
    
    Set OldList = New List
    
    For Idx = 0 To m_List.Count - 1
        OldList.Add m_List(Idx)
    Next
    
    Set m_List = New_c.ArrayList(vbVariant)
    
    Call OnCollectionChanged(CollectionChangedActionReset, Nothing, -1, OldList, 0)
End Sub

Public Function CollectionChangedEvent() As CollectionChangedEvent
    Set CollectionChangedEvent = m_CollectionChangedEvent
End Function

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

Public Function GetHashCode() As Long
    GetHashCode = m_Ptr
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
                
        If Object.Equals(m_List(Index), Value) Then Exit Property
        
        Dim OldItems As List
        Set OldItems = NewList(m_List(Index))
        
        m_List(Index) = Value
        
        Call OnCollectionChanged(CollectionChangedActionReplace, NewList(Value), Index, OldItems, Index)
    End Property
    
    Public Property Set Item(ByVal Index As Long, Value As Variant)
            
        If Index < 0 Or Index > m_List.Count - 1 Then
            ' Raise "Subscript out of range" error
            Err.Raise 9
            Exit Property
        End If
                
        If Object.Equals(m_List(Index), Value) Then Exit Property
        
        Dim OldItems As List
        Set OldItems = NewList(m_List(Index))
        
        Set m_List(Index) = Value
        
        Call OnCollectionChanged(CollectionChangedActionReplace, NewList(Value), Index, OldItems, Index)
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

'#Region Friend Members
    Friend Sub OnCollectionChanged(ByVal Action As CollectionChangedAction, _
                                                    ByVal NewItems As List, _
                                                    ByVal NewStartingIndex As Long, _
                                                    ByVal OldItems As List, _
                                                    ByVal OldStartingIndex As Long)
        
        Dim e As CollectionChangedEventArgs
        Dim DefaultView As ListCollectionView
        
        Set e = NewCollectionChangedEventArgs(Action, NewItems, NewStartingIndex, OldItems, OldStartingIndex)
        
        ' Notify the DefaultView (if any)
        Set DefaultView = modStaticClasses.CollectionViewSource.GetView(m_Ptr)
        If Not DefaultView Is Nothing Then DefaultView.OnSourceCollectionChanged Me, e
        
        RaiseEvent CollectionChanged(Me, e)
    
        m_CollectionChangedEvent.OnCollectionChanged e
    End Sub
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

'#Region INotifyCollectionChanged Interface
    Private Function INotifyCollectionChanged_CollectionChangedEvent() As CollectionChangedEvent
        Set INotifyCollectionChanged_CollectionChangedEvent = Me.CollectionChangedEvent
    End Function
'#End Region

