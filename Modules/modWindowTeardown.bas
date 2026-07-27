Attribute VB_Name = "modWindowTeardown"
Option Explicit

' Teardown helpers for Window.Unload vs compiled DLL.
' File log: %TEMP%\VCF_Unload.log (Debug.Print from DLL is often invisible).

Private m_Hold As Collection
Private m_HoldPermanent As Collection
Private Const LOG_NAME As String = "VCF_Unload.log"

Public Sub LogStep(ByVal Msg As String)
    Dim F As Integer
    Dim Path As String
    On Error Resume Next
    Path = Environ$("TEMP") & "\" & LOG_NAME
    F = FreeFile
    Open Path For Append As #F
    Print #F, Format$(Now, "hh:nn:ss") & "." & Right$("000" & (Timer * 1000#) Mod 1000, 3) & " " & Msg
    Close #F
End Sub

Public Sub PrepareTree(ByVal Root As Object)
    Dim Ctrl As IControl
    Dim Kids As UIElementCollection
    Dim i As Long
    Dim Child As Object
    Dim Snapshot As Collection

    On Error Resume Next
    If Root Is Nothing Then Exit Sub

    If TypeOf Root Is IControl Then
        Set Ctrl = Root
        Set Kids = Ctrl.Children
        If Not Kids Is Nothing Then
            Set Snapshot = New Collection
            For i = 0 To Kids.Count - 1
                Snapshot.Add Kids(i)
            Next
            For i = 1 To Snapshot.Count
                Set Child = Snapshot(i)
                PrepareTree Child
                Set Child = Nothing
            Next
        End If
    End If

    PrepareNode Root
End Sub

Public Sub PrepareWindowChildren(ByVal Win As Window)
    Dim i As Long
    Dim Snapshot As Collection
    Dim Child As Object

    On Error Resume Next
    If Win Is Nothing Then Exit Sub
    If Win.Children Is Nothing Then Exit Sub

    Set Snapshot = New Collection
    For i = 0 To Win.Children.Count - 1
        Snapshot.Add Win.Children(i)
    Next
    LogStep "PrepareWindowChildren count=" & CStr(Snapshot.Count)
    For i = 1 To Snapshot.Count
        Set Child = Snapshot(i)
        LogStep "PrepareTree " & TypeName(Child)
        PrepareTree Child
        Set Child = Nothing
    Next
    LogStep "PrepareWindowChildren done"
End Sub

Private Sub PrepareNode(ByVal Obj As Object)
    Dim Btn As Button
    Dim IC As ItemsControl
    Dim Ug As UniformGrid
    Dim Sp As StackPanel
    Dim Tb As TextBlock

    On Error Resume Next
    If Obj Is Nothing Then Exit Sub

    Select Case TypeName(Obj)
        Case "Button"
            Set Btn = Obj
            Btn.PrepareForUnload
        Case "ItemsControl"
            Set IC = Obj
            IC.PrepareForUnload
        Case "UniformGrid"
            Set Ug = Obj
            Ug.PrepareForUnload
        Case "StackPanel"
            Set Sp = Obj
            Sp.PrepareForUnload
        Case "TextBlock"
            Set Tb = Obj
            Tb.PrepareForUnload
    End Select
End Sub

' Move child refs to hold without Release (ClearSilent after Hold).
Public Sub Hold(ByVal Obj As Object)
    Dim El As IUIElement
    Dim Ctrl As IControl
    Dim Kids As UIElementCollection

    On Error Resume Next
    If Obj Is Nothing Then Exit Sub
    If TypeOf Obj Is IUIElement Then
        Set El = Obj
        Set El.Parent = Nothing
    End If
    If TypeOf Obj Is IControl Then
        Set Ctrl = Obj
        Set Kids = Ctrl.Children
        If Not Kids Is Nothing Then Kids.SeverParent
    End If
    If m_Hold Is Nothing Then Set m_Hold = New Collection
    m_Hold.Add Obj
End Sub

' Emptied ItemsHost panels — Release after Form.Unload hangs (see Disarm log).
Public Sub HoldPermanent(ByVal Obj As Object)
    Dim El As IUIElement
    Dim Ctrl As IControl
    Dim Kids As UIElementCollection

    On Error Resume Next
    If Obj Is Nothing Then Exit Sub
    If TypeOf Obj Is IUIElement Then
        Set El = Obj
        Set El.Parent = Nothing
    End If
    If TypeOf Obj Is IControl Then
        Set Ctrl = Obj
        Set Kids = Ctrl.Children
        If Not Kids Is Nothing Then Kids.SeverParent
    End If
    If m_HoldPermanent Is Nothing Then Set m_HoldPermanent = New Collection
    m_HoldPermanent.Add Obj
    LogStep "HoldPermanent " & TypeName(Obj) & " total=" & CStr(m_HoldPermanent.Count)
End Sub

Public Sub TransferAll(ByVal Kids As UIElementCollection)
    Dim i As Long
    Dim n As Long
    Dim Child As Object

    On Error Resume Next
    If Kids Is Nothing Then Exit Sub
    n = Kids.Count
    LogStep "TransferAll count=" & CStr(n)
    For i = 0 To n - 1
        Set Child = Kids(i)
        Hold Child
        Set Child = Nothing
    Next
    Kids.AbandonAll
    LogStep "TransferAll done held=" & CStr(HeldCount)
End Sub

Public Property Get HeldCount() As Long
    If m_Hold Is Nothing Then
        HeldCount = 0
    Else
        HeldCount = m_Hold.Count
    End If
End Property

' Release held trees. count=1 (e.g. TextBlock) is safe after Prepare.
' count>=2 with ItemsControls hangs on bulk Set m_Hold = Nothing vs compiled DLL.
' Mid-Unload defers; DrainHoldForce (before MsgBox) releases one-by-one after Disarm.
Public Sub DrainHold()
    Dim n As Long
    On Error Resume Next
    n = HeldCount
    LogStep "DrainHold begin count=" & CStr(n)
    If n <= 0 Then
        LogStep "DrainHold done empty"
        Exit Sub
    End If
    If n = 1 Then
        Set m_Hold = Nothing
        LogStep "DrainHold done small"
        Exit Sub
    End If
    LogStep "DrainHold deferred count=" & CStr(n)
End Sub

' One-by-one Release after ItemsControl.DisarmForRelease. Call before process End.
Public Sub DrainHoldForce()
    Dim Obj As Object
    Dim IC As ItemsControl
    Dim n As Long

    On Error Resume Next
    n = HeldCount
    LogStep "DrainHoldForce begin count=" & CStr(n)

    Do While HeldCount > 0
        Set Obj = m_Hold(1)
        m_Hold.Remove 1
        LogStep "DrainHoldForce release " & TypeName(Obj) & " remaining=" & CStr(HeldCount)

        If TypeOf Obj Is ItemsControl Then
            Set IC = Obj
            IC.DisarmForRelease
            LogStep "DrainHoldForce disarmed ItemsControl"
            Set IC = Nothing
        Else
            ' StackPanel etc. from B-NAV — park only (Prepare already ran before Form.Unload).
            LogStep "DrainHoldForce park " & TypeName(Obj)
            HoldPermanent Obj
        End If

        Set Obj = Nothing
        LogStep "DrainHoldForce released ok"
    Loop

    Set m_Hold = Nothing
    LogStep "DrainHoldForce done"
End Sub

' Release HoldPermanent one-by-one. Call only while Form still alive.
Public Sub DrainHoldPermanent()
    Dim Obj As Object
    Dim n As Long

    On Error Resume Next
    If m_HoldPermanent Is Nothing Then
        LogStep "DrainHoldPermanent empty"
        Exit Sub
    End If
    n = m_HoldPermanent.Count
    LogStep "DrainHoldPermanent begin count=" & CStr(n)

    Do While m_HoldPermanent.Count > 0
        Set Obj = m_HoldPermanent(1)
        m_HoldPermanent.Remove 1
        LogStep "DrainHoldPermanent release " & TypeName(Obj) & " remaining=" & CStr(m_HoldPermanent.Count)
        Set Obj = Nothing
        LogStep "DrainHoldPermanent released ok"
    Loop

    Set m_HoldPermanent = Nothing
    LogStep "DrainHoldPermanent done"
End Sub
