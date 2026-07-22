Attribute VB_Name = "modBindingSourceResolve"
Option Explicit

' Resolve RelativeSource / ElementName for BindingsManager and Phase0.

Public Function ResolveRelativeSource(ByVal Target As Object, ByVal RS As RelativeSource) As Object
    Dim Full As Object

    On Error GoTo Handler

    If Target Is Nothing Then Exit Function
    If RS Is Nothing Then Exit Function

    ' Widen IDependencyObject / other iface ptrs to coclass so CallByName
    ' can see TemplatedParent / Parent (VB6 interface narrowing).
    Set Full = API.CObj(Target)
    If Full Is Nothing Then Set Full = Target

    Select Case RS.Mode
        Case RelativeSourceSelf
            Set ResolveRelativeSource = Full
        Case RelativeSourceTemplatedParent
            Set ResolveRelativeSource = FindTemplatedParent(Full)
        Case RelativeSourceFindAncestor
            Set ResolveRelativeSource = FindAncestor(Full, RS.AncestorType, RS.AncestorLevel)
    End Select
    Exit Function

Handler:
    Set ResolveRelativeSource = Nothing
End Function

Public Function ResolveElementName(ByVal Target As Object, ByVal ElementName As String, Optional ByVal Root As IControl = Nothing) As Object
    Dim Scope As IControl
    Dim Named As ObservableDictionary
    Dim NM As NamingManager

    On Error GoTo Handler

    If Len(ElementName) = 0 Then Exit Function

    Set Scope = Root
    If Scope Is Nothing Then Set Scope = FindNamescopeRoot(Target)
    If Scope Is Nothing Then Exit Function

    Set NM = New NamingManager
    Set Named = NM.GetNamedChildren(Scope)
    If Named Is Nothing Then Exit Function
    If Not Named.ContainsKey(ElementName) Then Exit Function
    Set ResolveElementName = Named(ElementName)
    Exit Function

Handler:
    Set ResolveElementName = Nothing
End Function

Public Function FindTemplatedParent(ByVal Start As Object) As Object
    Dim Cur As Object
    Dim NextParent As Object
    Dim TP As Object

    On Error Resume Next

    Set Cur = Start
    Do While Not Cur Is Nothing
        Set TP = Nothing
        Err.Clear
        Set TP = CallByName(Cur, "TemplatedParent", VbGet)
        If Err.Number = 0 Then
            If Not TP Is Nothing Then
                Set FindTemplatedParent = TP
                Err.Clear
                Exit Function
            End If
        End If
        Err.Clear

        Set NextParent = Nothing
        Set NextParent = CallByName(Cur, "Parent", VbGet)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        Set Cur = NextParent
    Loop
    Err.Clear
End Function

Public Function FindNamescopeRoot(ByVal Start As Object) As IControl
    Dim Cur As Object
    Dim NextParent As Object

    On Error Resume Next

    If Start Is Nothing Then Exit Function
    Set Cur = Start
    Do
        Set NextParent = Nothing
        Err.Clear
        Set NextParent = CallByName(Cur, "Parent", VbGet)
        If Err.Number <> 0 Or NextParent Is Nothing Then
            Err.Clear
            If TypeOf Cur Is IControl Then Set FindNamescopeRoot = Cur
            Exit Function
        End If
        Set Cur = NextParent
    Loop
End Function

Public Function FindAncestor(ByVal Start As Object, ByVal AncestorType As String, ByVal AncestorLevel As Long) As Object
    Dim Cur As Object
    Dim NextParent As Object
    Dim Level As Long
    Dim Need As Long
    Dim Tn As String

    On Error Resume Next

    If Start Is Nothing Then Exit Function
    Need = AncestorLevel
    If Need < 1 Then Need = 1

    Set Cur = CallByName(Start, "Parent", VbGet)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Function
    End If

    Do While Not Cur Is Nothing
        Tn = TypeName(Cur)
        If Len(AncestorType) = 0 Or StrComp(Tn, AncestorType, vbTextCompare) = 0 Then
            Level = Level + 1
            If Level >= Need Then
                Set FindAncestor = Cur
                Exit Function
            End If
        End If

        Set NextParent = Nothing
        Err.Clear
        Set NextParent = CallByName(Cur, "Parent", VbGet)
        If Err.Number <> 0 Then
            Err.Clear
            Exit Function
        End If
        Set Cur = NextParent
    Loop
    Err.Clear
End Function
