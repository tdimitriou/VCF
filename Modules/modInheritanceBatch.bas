Attribute VB_Name = "modInheritanceBatch"
Option Explicit

' Phase 8a — suppress PassPropertyValue fan-out during XAML load / style apply,
' then optionally propagate once from the batch root (DataContext coalesce).

Private m_Depth As Long
Private m_Dirty As Boolean
Private m_RootPtr As Long

Public Sub BeginInheritanceUpdate(Optional ByVal Root As Object = Nothing)
    m_Depth = m_Depth + 1
    If m_Depth = 1 Then
        m_Dirty = False
        If Root Is Nothing Then
            m_RootPtr = 0
        Else
            m_RootPtr = ObjPtr(Root)
        End If
    End If
End Sub

Public Sub EndInheritanceUpdate()
    Dim Root As Object
    Dim Dirty As Boolean
    Dim RootPtr As Long

    If m_Depth <= 0 Then
        Err.Raise 5, "modInheritanceBatch", "EndInheritanceUpdate without matching BeginInheritanceUpdate"
    End If

    m_Depth = m_Depth - 1
    If m_Depth > 0 Then Exit Sub

    Dirty = m_Dirty
    RootPtr = m_RootPtr
    m_Dirty = False
    m_RootPtr = 0

    If Not Dirty Then Exit Sub
    If RootPtr = 0 Then Exit Sub

    On Error Resume Next
    Set Root = API.ObjFromPtr(RootPtr)
    If Root Is Nothing Then Exit Sub
    If Not TypeOf Root Is IDependencyObject Then Exit Sub

    DependencyPropertiesStatic.PropagateInheritableFrom Root
End Sub

Public Function IsInheritanceBatchActive() As Boolean
    IsInheritanceBatchActive = (m_Depth > 0)
End Function

Public Sub MarkInheritanceDirty()
    If m_Depth > 0 Then m_Dirty = True
End Sub

Public Sub SetInheritanceBatchRoot(ByVal Root As Object)
    ' Only the outermost batch owns the propagate root. Nested XAMLReader.Load
    ' (res: fragments) must not steal it — otherwise End notifies the wrong tree
    ' and screen-level DataContext bindings/commands never attach.
    If m_Depth <> 1 Then Exit Sub
    If Root Is Nothing Then
        m_RootPtr = 0
    Else
        m_RootPtr = ObjPtr(Root)
    End If
End Sub
