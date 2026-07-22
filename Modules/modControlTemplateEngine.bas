Attribute VB_Name = "modControlTemplateEngine"
Option Explicit

Public Sub ApplyControlTemplate(ByVal Style As Style, ByVal Target As Object)
    Dim Tmpl As ControlTemplate

    On Error GoTo Handler

100 If Style Is Nothing Then Exit Sub
102 If Target Is Nothing Then Exit Sub

110 Set Tmpl = Style.Template
112 If Tmpl Is Nothing Then Exit Sub

120 If Len(Tmpl.TargetType) > 0 Then
122     If TypeName(Target) <> Tmpl.TargetType Then Exit Sub
    End If

130 If Tmpl.Children.Count = 0 Then Exit Sub

140 Select Case TypeName(Target)
        Case "Button"
150         ApplyButtonTemplate Target, Tmpl
    End Select

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modControlTemplateEngine", "ApplyControlTemplate"
End Sub

' Lookless-prep (P6f-TBIND): Border chrome only here.
' ContentPresenter-slot alignment is applied via StyleManager.PushTemplateContentAlignment.
Private Sub ApplyButtonTemplate(ByVal Btn As Button, ByVal Tmpl As ControlTemplate)
    Dim i As Long
    Dim Node As Object
    Dim B As Border
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
        End If
NextNode:
260 Next

270 If B Is Nothing Then Exit Sub

280 Call API.CopyVariable(B.DependencyProperties.GetValue("CornerRadius"), Rad)
290 If Rad.TopLeft > 0# Then Btn.CornerRadius = Rad.TopLeft

    ' Border does not register BackColor as a DP (widget BackColor only).
    ' Unconditional GetValue("BackColor") raised 424 and aborted Style apply
    ' before PushTemplateContentAlignment — P6f HAlign stayed at default.
300 If B.DependencyProperties.Exists("BackColor") Then
310     Call API.CopyVariable(B.DependencyProperties.GetValue("BackColor"), BackColor)
320     If Not IsEmpty(BackColor) And Not IsNull(BackColor) Then
330         Btn.DependencyProperties.SetCurrentValue "BackColor", BackColor
        End If
    End If

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modControlTemplateEngine", "ApplyButtonTemplate"
End Sub
