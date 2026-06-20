Attribute VB_Name = "modCollectionNotifications"
Option Explicit

' Reusable single-item List buffers for collection change notifications.
' Avoids allocating a new List on every Add/Remove/Replace/Move when not reentrant.

Private m_ScratchNewItems As List
Private m_ScratchOldItems As List
Private m_InNotify As Long

Private Sub EnsureScratchLists()
    If m_ScratchNewItems Is Nothing Then Set m_ScratchNewItems = New List
    If m_ScratchOldItems Is Nothing Then Set m_ScratchOldItems = New List
End Sub

Public Sub NotifyEnter()
    m_InNotify = m_InNotify + 1
End Sub

Public Sub NotifyLeave()
    If m_InNotify > 0 Then m_InNotify = m_InNotify - 1
End Sub

Public Function ScratchNewForItem(ByVal Value As Variant) As List
    Dim L As List
    
    If m_InNotify > 0 Then
        Set L = New List
    Else
        EnsureScratchLists
        Set L = m_ScratchNewItems
    End If
    
    L.SetSingleItem Value
    Set ScratchNewForItem = L
End Function

Public Function ScratchOldForItem(ByVal Value As Variant) As List
    Dim L As List
    
    If m_InNotify > 0 Then
        Set L = New List
    Else
        EnsureScratchLists
        Set L = m_ScratchOldItems
    End If
    
    L.SetSingleItem Value
    Set ScratchOldForItem = L
End Function
