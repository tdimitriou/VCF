Attribute VB_Name = "modControlTemplateEngine"
Option Explicit

Public Sub ApplyControlTemplate(ByVal Style As Style, ByVal Target As Object)
    Dim Tmpl As ControlTemplate

    On Error GoTo Handler

100 If Style Is Nothing Then Exit Sub
102 If Target Is Nothing Then Exit Sub

110 Set Tmpl = Style.Template
112 If Tmpl Is Nothing Then
114     If TypeName(Target) = "Button" Then
116         Dim BtnClear As Button
118         Set BtnClear = Target
120         BtnClear.ClearTemplateChrome
        End If
122     Exit Sub
    End If

130 If Len(Tmpl.TargetType) > 0 Then
132     If TypeName(Target) <> Tmpl.TargetType Then Exit Sub
    End If

140 If Tmpl.Children.Count = 0 Then Exit Sub

150 Select Case TypeName(Target)
        Case "Button"
160         ApplyButtonTemplate Target, Tmpl
    End Select

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modControlTemplateEngine", "ApplyControlTemplate"
End Sub

' Lookless (P6g-P6j): clone template visual tree under Button.
' Supports flat Border+CP siblings and deeper Border->Panel->ContentPresenter.
Private Sub ApplyButtonTemplate(ByVal Btn As Button, ByVal Tmpl As ControlTemplate)
    Dim i As Long
    Dim Node As Object
    Dim B As Border
    Dim Clone As Border
    Dim CPSrc As ContentPresenter
    Dim CPClone As ContentPresenter
    Dim Rad As VCF.CornerRadius
    Dim BackColor As Variant
    Dim Tn As String
    Dim NestUnderChrome As Boolean

    On Error GoTo Handler

200 For i = 0 To Tmpl.Children.Count - 1
210     Set Node = Tmpl.Children(i)
212     If Node Is Nothing Then GoTo NextNode
220     Tn = TypeName(Node)
230     If StrComp(Tn, "Border", vbTextCompare) = 0 Then
240         If B Is Nothing Then
                On Error Resume Next
250             Set B = Node
                If Err.Number <> 0 Then
                    Err.Clear
                    Set B = Nothing
                End If
                On Error GoTo Handler
            End If
260     ElseIf StrComp(Tn, "ContentPresenter", vbTextCompare) = 0 Then
270         If CPSrc Is Nothing Then
                On Error Resume Next
280             Set CPSrc = Node
                If Err.Number <> 0 Then
                    Err.Clear
                    Set CPSrc = Nothing
                End If
                On Error GoTo Handler
            End If
        End If
NextNode:
290 Next

300 If B Is Nothing Then
302     Btn.ClearTemplateChrome
304     Exit Sub
    End If

    ' Deeper tree: ContentPresenter under Border.Child / panel children.
310 If CPSrc Is Nothing Then Set CPSrc = FindContentPresenterInSubtree(B)

320 Call API.CopyVariable(B.DependencyProperties.GetValue("CornerRadius"), Rad)
330 If Rad.TopLeft > 0# Then Btn.CornerRadius = Rad.TopLeft

340 If B.DependencyProperties.Exists("BackColor") Then
350     Call API.CopyVariable(B.DependencyProperties.GetValue("BackColor"), BackColor)
360     If Not IsEmpty(BackColor) And Not IsNull(BackColor) Then
370         Btn.DependencyProperties.SetCurrentValue "BackColor", BackColor
        End If
    End If

380 Set CPClone = Nothing
390 NestUnderChrome = True
400 Set Clone = CloneBorderSubtree(B, Btn, CPClone, NestUnderChrome)
410 If Clone Is Nothing Then
412     Btn.ClearTemplateChrome
414     Exit Sub
    End If

420 Btn.AttachTemplateChrome Clone
430 Call ApplyCloneCornerRadius(Clone, Rad)

440 If CPClone Is Nothing Then
450     If Not CPSrc Is Nothing Or Tmpl.HasContentAlignmentMarker Then
460         Set CPClone = New ContentPresenter
462         Set CPClone.TemplatedParent = Btn
470         If Not CPSrc Is Nothing Then
480             CPClone.HorizontalContentAlignment = CPSrc.HorizontalContentAlignment
490             CPClone.VerticalContentAlignment = CPSrc.VerticalContentAlignment
500         ElseIf Tmpl.HasContentAlignmentMarker Then
510             CPClone.HorizontalContentAlignment = Tmpl.ContentHorizontalAlignment
520             CPClone.VerticalContentAlignment = Tmpl.ContentVerticalAlignment
            End If
530         NestUnderChrome = True
        End If
    End If

540 If Not CPClone Is Nothing Then
550     Btn.AttachTemplatePresenter CPClone, NestUnderChrome
    End If

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modControlTemplateEngine", "ApplyButtonTemplate"
End Sub

Private Function FindContentPresenterInSubtree(ByVal Node As Object) As ContentPresenter
    Dim Tn As String
    Dim B As Border
    Dim Child As Object
    Dim Found As ContentPresenter
    Dim Kids As Object

    On Error Resume Next

    If Node Is Nothing Then Exit Function
    Tn = TypeName(Node)
    If StrComp(Tn, "ContentPresenter", vbTextCompare) = 0 Then
        Set FindContentPresenterInSubtree = Node
        Exit Function
    End If

    If StrComp(Tn, "Border", vbTextCompare) = 0 Then
        Set B = Node
        If Not B.Child Is Nothing Then
            Set Found = FindContentPresenterInSubtree(B.Child)
            If Not Found Is Nothing Then
                Set FindContentPresenterInSubtree = Found
                Exit Function
            End If
        End If
    End If

    Set Kids = Nothing
    Set Kids = Node.Children
    If Kids Is Nothing Then Exit Function
    For Each Child In Kids
        Set Found = FindContentPresenterInSubtree(Child)
        If Not Found Is Nothing Then
            Set FindContentPresenterInSubtree = Found
            Exit Function
        End If
    Next
End Function

' Clone Border chrome + optional Child/Children subtree. OutCP receives live ContentPresenter.
Private Function CloneBorderSubtree(ByVal Src As Border, ByVal Btn As Button, ByRef OutCP As ContentPresenter, ByRef NestUnderChrome As Boolean) As Border
    Dim Clone As Border
    Dim ChildClone As Object

    On Error GoTo Handler

    Set Clone = New Border
    Clone.Widget.BackColor = Btn.Widget.BackColor
    Clone.BorderColor = Src.BorderColor
    Set Clone.TemplatedParent = Btn

    If Not Src.Child Is Nothing Then
        Set ChildClone = CloneTemplateNode(Src.Child, Btn, OutCP)
        If Not ChildClone Is Nothing Then
            If TypeOf ChildClone Is IUIElement Then
                Set Clone.Child = ChildClone
                ' CP already placed in subtree ? do not overwrite Border.Child in AttachTemplatePresenter.
                If Not OutCP Is Nothing Then NestUnderChrome = False
            End If
        End If
    End If

    Set CloneBorderSubtree = Clone
    Exit Function

Handler:
    Set CloneBorderSubtree = Nothing
End Function

Private Function CloneTemplateNode(ByVal Src As Object, ByVal Btn As Button, ByRef OutCP As ContentPresenter) As Object
    Dim Tn As String
    Dim CPSrc As ContentPresenter
    Dim CPClone As ContentPresenter
    Dim GridClone As Grid
    Dim StackClone As StackPanel
    Dim Child As Object
    Dim ChildClone As Object

    On Error GoTo Handler

    If Src Is Nothing Then Exit Function
    Tn = TypeName(Src)

    If StrComp(Tn, "ContentPresenter", vbTextCompare) = 0 Then
        Set CPSrc = Src
        Set CPClone = New ContentPresenter
        CPClone.HorizontalContentAlignment = CPSrc.HorizontalContentAlignment
        CPClone.VerticalContentAlignment = CPSrc.VerticalContentAlignment
        Set CPClone.TemplatedParent = Btn
        Set OutCP = CPClone
        Set CloneTemplateNode = CPClone
        Exit Function
    End If

    If StrComp(Tn, "Grid", vbTextCompare) = 0 Then
        Set GridClone = New Grid
        Set GridClone.TemplatedParent = Btn
        For Each Child In Src.Children
            Set ChildClone = CloneTemplateNode(Child, Btn, OutCP)
            If Not ChildClone Is Nothing Then
                If TypeOf ChildClone Is IUIElement Then GridClone.Children.Add ChildClone
            End If
        Next
        Set CloneTemplateNode = GridClone
        Exit Function
    End If

    If StrComp(Tn, "StackPanel", vbTextCompare) = 0 Then
        Set StackClone = New StackPanel
        Set StackClone.TemplatedParent = Btn
        For Each Child In Src.Children
            Set ChildClone = CloneTemplateNode(Child, Btn, OutCP)
            If Not ChildClone Is Nothing Then
                If TypeOf ChildClone Is IUIElement Then StackClone.Children.Add ChildClone
            End If
        Next
        Set CloneTemplateNode = StackClone
        Exit Function
    End If

    If StrComp(Tn, "Border", vbTextCompare) = 0 Then
        Dim NestedBorder As Border
        Dim NestedClone As Border
        Dim NestedNest As Boolean
        Set NestedBorder = Src
        NestedNest = True
        Set NestedClone = CloneBorderSubtree(NestedBorder, Btn, OutCP, NestedNest)
        Set CloneTemplateNode = NestedClone
        Exit Function
    End If

    Exit Function

Handler:
    Set CloneTemplateNode = Nothing
End Function

Private Sub ApplyCloneCornerRadius(ByVal Clone As Border, ByRef Src As VCF.CornerRadius)
    Dim OutRad As VCF.CornerRadius

    On Error GoTo Handler

600 OutRad.TopLeft = Src.TopLeft
610 OutRad.TopRight = Src.TopRight
620 OutRad.BottomLeft = Src.BottomLeft
630 OutRad.BottomRight = Src.BottomRight
640 Clone.DependencyProperties.SetValue "CornerRadius", OutRad

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modControlTemplateEngine", "ApplyCloneCornerRadius"
End Sub
