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
            For Each UIElement In e.NewItems
                Set UIElement.Parent = Parent
            Next
        Case CollectionChangedActionRemove
            For Each UIElement In e.OldItems
                Set UIElement.Parent = Nothing
            Next
        Case CollectionChangedActionReplace
            For Each UIElement In e.OldItems
                Set UIElement.Parent = Nothing
            Next
            
            For Each UIElement In e.NewItems
                Set UIElement.Parent = Parent
            Next
        Case CollectionChangedActionMove
            '
        Case CollectionChangedActionReset
            For Each UIElement In e.OldItems
                Set UIElement.Parent = Nothing
            Next
    End Select
            
    RaiseEvent CollectionChanged(Me, e)
    
    Dim ParentWin As Window
    Set ParentWin = GetParentWindow(Parent)
    
    If Not ParentWin Is Nothing Then ParentWin.OnChildElementsChanged Parent, e
    
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
