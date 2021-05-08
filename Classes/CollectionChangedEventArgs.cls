VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionChangedEventArgs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Enum CollectionChangedAction
    CollectionChangedActionAdd = 0
    CollectionChangedActionRemove = 1
    CollectionChangedActionReplace = 2
    CollectionChangedActionMove = 3
    CollectionChangedActionReset = 4
End Enum

Private m_Action As CollectionChangedAction
Private m_NewItems As List
Private m_NewStartingIndex As Long
Private m_OldItems As List
Private m_OldStartingIndex As Long

' Gets the action that caused the event.
Public Property Get Action() As CollectionChangedAction
    Action = m_Action
End Property

' Gets the list of new items involved in the change.
Public Property Get NewItems() As List
    Set NewItems = m_NewItems
End Property

' Gets the index at which the change occurred.
Public Property Get NewStartingIndex() As Long
    NewStartingIndex = m_NewStartingIndex
End Property

' Gets the list of items affected by a Replace, Remove, or Move action.
Public Property Get OldItems() As List
    Set OldItems = m_OldItems
End Property

' Gets the index at which a Move, Remove, or Replace action occurred.
Public Property Get OldStartingIndex() As Long
    OldStartingIndex = m_OldStartingIndex
End Property

Friend Sub Initialize(ByVal Action As CollectionChangedAction, _
                        ByVal NewItems As List, _
                        ByVal NewStartingIndex As Long, _
                        ByVal OldItems As List, _
                        ByVal OldStartingIndex As Long)
    m_Action = Action
    Set m_NewItems = NewItems
    m_NewStartingIndex = NewStartingIndex
    Set m_OldItems = OldItems
    m_OldStartingIndex = OldStartingIndex
End Sub
