Attribute VB_Name = "modControlTemplateEngine"
Option Explicit

Public Sub ApplyControlTemplate(ByVal Style As Style, ByVal Target As Object)
    Dim Tmpl As ControlTemplate

    If Style Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    Set Tmpl = Style.Template
    If Tmpl Is Nothing Then Exit Sub

    If Len(Tmpl.TargetType) > 0 Then
        If TypeName(Target) <> Tmpl.TargetType Then Exit Sub
    End If

    If Tmpl.Children.Count = 0 Then Exit Sub

    Select Case TypeName(Target)
        Case "Button"
            ApplyButtonTemplate Target, Tmpl.Children(0)
    End Select
End Sub

Private Sub ApplyButtonTemplate(ByVal Btn As Button, ByVal Root As Object)
    If Not TypeOf Root Is Border Then Exit Sub

    Dim B As Border
    Dim Rad As VCF.CornerRadius
    Dim BackColor As Variant

    Set B = Root

    Call API.CopyVariable(B.DependencyProperties.GetValue("CornerRadius"), Rad)
    If Rad.TopLeft > 0# Then Btn.CornerRadius = Rad.TopLeft

    Call API.CopyVariable(B.DependencyProperties.GetValue("BackColor"), BackColor)
    If Not IsEmpty(BackColor) And Not IsNull(BackColor) Then
        Btn.DependencyProperties.SetCurrentValue "BackColor", BackColor
    End If
End Sub
