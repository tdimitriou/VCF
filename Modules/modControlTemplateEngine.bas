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

' Lookless (P6g/P6h): clone first template Border into Button.Children;
' clone ContentPresenter as paint-only TemplateBinding content slot.
' Alignment marker still pushed via StyleManager.PushTemplateContentAlignment.
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

310 Call API.CopyVariable(B.DependencyProperties.GetValue("CornerRadius"), Rad)
320 If Rad.TopLeft > 0# Then Btn.CornerRadius = Rad.TopLeft

    ' Border does not register BackColor as a DP (widget BackColor only).
    ' Unconditional GetValue("BackColor") raised 424 and aborted Style apply
    ' before PushTemplateContentAlignment - P6f HAlign stayed at default.
330 If B.DependencyProperties.Exists("BackColor") Then
340     Call API.CopyVariable(B.DependencyProperties.GetValue("BackColor"), BackColor)
350     If Not IsEmpty(BackColor) And Not IsNull(BackColor) Then
360         Btn.DependencyProperties.SetCurrentValue "BackColor", BackColor
        End If
    End If

    ' Clone - never attach the template-bag instance (shared Style/template).
370 Set Clone = New Border
380 Clone.Widget.BackColor = Btn.Widget.BackColor
390 Clone.BorderColor = B.BorderColor
400 Btn.AttachTemplateChrome Clone
    ' CornerRadius after parenting; rebuild UDT fields so SetValue sticks.
410 Call ApplyCloneCornerRadius(Clone, Rad)

    ' Live content slot (paint-only; no widget) when template has CP and/or align marker.
420 If Not CPSrc Is Nothing Or Tmpl.HasContentAlignmentMarker Then
430     Set CPClone = New ContentPresenter
440     If Not CPSrc Is Nothing Then
450         CPClone.HorizontalContentAlignment = CPSrc.HorizontalContentAlignment
460         CPClone.VerticalContentAlignment = CPSrc.VerticalContentAlignment
470     ElseIf Tmpl.HasContentAlignmentMarker Then
480         CPClone.HorizontalContentAlignment = Tmpl.ContentHorizontalAlignment
490         CPClone.VerticalContentAlignment = Tmpl.ContentVerticalAlignment
        End If
500     Btn.AttachTemplatePresenter CPClone
    End If

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modControlTemplateEngine", "ApplyButtonTemplate"
End Sub

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
