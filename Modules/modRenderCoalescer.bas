Attribute VB_Name = "modRenderCoalescer"
Option Explicit

Private m_UpdateDepth As Long
Private m_Pending As cSortedDictionary
Private m_PendingReady As Boolean
Private m_LastFlushCount As Long

Public Sub BeginRenderUpdate()
    m_UpdateDepth = m_UpdateDepth + 1
End Sub

Public Sub EndRenderUpdate()
    If m_UpdateDepth <= 0 Then
        Err.Raise 5, "modRenderCoalescer", "EndRenderUpdate without matching BeginRenderUpdate"
    End If

    m_UpdateDepth = m_UpdateDepth - 1
    If m_UpdateDepth = 0 Then FlushPendingRefreshes
End Sub

Public Sub RequestWidgetRefresh(ByVal W As cWidgetBase)
    Dim Key As String

    If W Is Nothing Then Exit Sub

    If m_UpdateDepth > 0 Then
        EnsurePending
        Key = CStr(ObjPtr(W))
        If Not m_Pending.Exists(Key) Then m_Pending.Add Key, W
        Exit Sub
    End If

    m_LastFlushCount = 1
    If Not W.LockRefresh Then W.Refresh
End Sub

Public Function PendingRefreshCount() As Long
    If Not m_PendingReady Then
        PendingRefreshCount = 0
    Else
        PendingRefreshCount = m_Pending.Count
    End If
End Function

Public Property Get LastFlushRefreshCount() As Long
    LastFlushRefreshCount = m_LastFlushCount
End Property

Private Sub EnsurePending()
    If m_PendingReady Then Exit Sub
    Set m_Pending = New_c.SortedDictionary
    m_PendingReady = True
End Sub

Private Sub FlushPendingRefreshes()
    Dim Pending() As cWidgetBase
    Dim Count As Long
    Dim i As Long
    Dim W As cWidgetBase

    If Not m_PendingReady Then Exit Sub
    Count = m_Pending.Count
    If Count = 0 Then Exit Sub

    m_LastFlushCount = Count
    ReDim Pending(0 To Count - 1)

    For i = 0 To Count - 1
        Set Pending(i) = m_Pending.ItemByIndex(i)
    Next

    m_Pending.RemoveAll

    For i = 0 To Count - 1
        Set W = Pending(i)
        If Not W Is Nothing Then
            On Error Resume Next
            If Not W.LockRefresh Then W.Refresh
            Err.Clear
        End If
    Next
End Sub
