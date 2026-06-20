VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIElementCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)

Private WithEvents m_List As ObservableCollection
Attribute m_List.VB_VarHelpID = -1
Private m_ParentPtr As Long

'#Region Class Events
Private Sub Class_Initialize()
    Set m_List = New ObservableCollection
End Sub
'#End Region

'#Region Public Methods
Public Sub Add(Value As Object)
    m_List.Add Value
End Sub

Public Sub Insert(ByVal Index As Long, Value As Object)
    m_List.Insert Index, Value
End Sub

Public Sub AddRange(List)
    m_List.AddRange List
End Sub

Public Sub Remove(Value As Object)
    Call m_List.Remove(Value)
End Sub

Public Sub RemoveRange(ByVal Index As Long, ByVal Count As Long)
    Call m_List.RemoveRange(Index, Count)
End Sub

Public Sub RemoveAt(ByVal Index As Long)
    Call m_List.RemoveAt(Index)
End Sub

Public Sub Clear()
    m_List.Clear
End Sub

Friend Sub AddSilent(Value As Object)
    Dim UIElement As IUIElement

    If m_ParentPtr <> 0 Then
        Set UIElement = Value
        Set UIElement.Parent = Parent
    End If
    m_List.AddSilent Value
End Sub

Friend Sub ClearSilent()
    Dim UIElement As IUIElement
    Dim i As Long

    For i = 0 To m_List.Count - 1
        Set UIElement = m_List(i)
        Set UIElement.Parent = Nothing
    Next i
    m_List.ClearSilent
End Sub

Public Function Contains(Value As Object) As Boolean
    Contains = m_List.Contains(Value)
End Function

Public Function IndexOf(Value As Object) As Long
    IndexOf = m_List.IndexOf(Value)
End Function

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = m_List.GetEnumerator
End Function
'#End Region

'#Region Public Properties
Public Property Get Count() As Long
    Count = m_List.Count
End Property

Public Property Get Item(ByVal Index As Long) As Object
Attribute Item.VB_UserMemId = 0
    Set Item = m_List(Index)
End Property

Public Property Set Item(ByVal Index As Long, Value As Object)
    Set m_List(Index) = Value
End Property

Public Property Get Parent() As Object
    If m_ParentPtr = 0 Then Exit Property
    Set Parent = API.ObjFromPtr(m_ParentPtr)
End Property
'#End Region

'#Region Friend Methods
Public Sub Initialize(Parent As Object)
    If Parent Is Nothing Then
        Err.Raise 91
        Exit Sub
    End If
    
    m_ParentPtr = ObjPtr(Parent)
End Sub
'#End Region

'#Region ObservableCollection Events
Private Sub m_List_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
    Dim UIElement As IUIElement
    Dim Index As Long
    
    
    Select Case e.Action
        Case CollectionChangedActionAdd
            If m_ParentPtr <> 0 Then
                If Not e.NewItems Is Nothing Then
                    For Each UIElement In e.NewItems
                        Set UIElement.Parent = Parent
                    Next
                End If
            End If
        Case CollectionChangedActionRemove
            If Not e.OldItems Is Nothing Then
                For Each UIElement In e.OldItems
                    Set UIElement.Parent = Nothing
                Next
            End If
        Case CollectionChangedActionReplace
            If Not e.OldItems Is Nothing Then
                For Each UIElement In e.OldItems
                    Set UIElement.Parent = Nothing
                Next
            End If
            
            If m_ParentPtr <> 0 Then
                If Not e.NewItems Is Nothing Then
                    For Each UIElement In e.NewItems
                        Set UIElement.Parent = Parent
                    Next
                End If
            End If
        Case CollectionChangedActionMove
            '
        Case CollectionChangedActionReset
            If Not e.OldItems Is Nothing Then
                For Each UIElement In e.OldItems
                    Set UIElement.Parent = Nothing
                Next
            End If
    End Select
            
    RaiseEvent CollectionChanged(Me, e)
    
    If m_ParentPtr <> 0 Then
        Dim ParentWin As Window
        Set ParentWin = GetParentWindow(Parent)
        If Not ParentWin Is Nothing Then ParentWin.OnChildElementsChanged Parent, e
    End If
    
End Sub
'#End Region

Private Function GetParentWindow(ByVal Obj As IUIElement) As Window
    On Error Resume Next
    
    If Obj Is Nothing Then Exit Function

    If TypeOf Obj Is Window Then
        Set GetParentWindow = Obj
    Else
        Set GetParentWindow = GetParentWindow(Obj.Parent)
    End If
End Function
