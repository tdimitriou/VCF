Attribute VB_Name = "modHarnessResize"
Option Explicit

' Borderless shell (BorderStyle=0) has no resize grips — use preset sizes (Shift+1/2/3)
' or set USE_SIZABLE_SHELL_BORDER=True for drag-resize with title bar (dev only).
'
' Framework (3.6.x+): Window fills visible UserControl views to the client; UserControl
' scales absolute Margin/Width children from Width/Height design canvas → widget size.
' Harness only changes Form client size + RelayoutChildren.

Private Const SHELL_DESIGN_W As Long = 1024
Private Const SHELL_DESIGN_H As Long = 768

Public Sub ApplyShellClientSize(ByVal Win As VCF.Window, ByVal Form As cWidgetForm, _
    ByVal ClientW As Long, ByVal ClientH As Long, ByVal Label As String)

    On Error GoTo Fail

    SetFormClientSize Form, ClientW, ClientH

    Form.WidgetRoot.Refresh
    Win.RelayoutChildren
    Form.WidgetRoot.Refresh

    Debug.Print "[HARNESS-RESIZE] " & Label & " -> " & ClientW & "x" & ClientH & _
                " client Scale=" & Form.ScaleWidth & "x" & Form.ScaleHeight & _
                " design=" & SHELL_DESIGN_W & "x" & SHELL_DESIGN_H

    Exit Sub

Fail:

    Debug.Print "[HARNESS-RESIZE] FAIL " & Label & ": " & Err.Description

End Sub

Private Sub SetFormClientSize(ByVal Form As cWidgetForm, ByVal TargetW As Long, ByVal TargetH As Long)
    Form.Width = TargetW
    Form.Height = TargetH

    If Not modHarnessConfig.USE_SIZABLE_SHELL_BORDER Then Exit Sub

    ' BorderStyle=2: Form.Width/Height is outer frame; ScaleWidth/Height is client.
    Dim DeltaW As Long
    Dim DeltaH As Long

    DeltaW = TargetW - CLng(Form.ScaleWidth)
    DeltaH = TargetH - CLng(Form.ScaleHeight)
    If DeltaW <> 0 Or DeltaH <> 0 Then
        Form.Width = Form.Width + DeltaW
        Form.Height = Form.Height + DeltaH
    End If
End Sub

Public Sub ApplyPreset1024(ByVal Win As VCF.Window, ByVal Form As cWidgetForm)
    ApplyShellClientSize Win, Form, SHELL_DESIGN_W, SHELL_DESIGN_H, "Shift+1 baseline"
End Sub

Public Sub ApplyPreset800(ByVal Win As VCF.Window, ByVal Form As cWidgetForm)
    ApplyShellClientSize Win, Form, 800, 600, "Shift+2 compact"
End Sub

Public Sub ApplyPresetWidescreen(ByVal Win As VCF.Window, ByVal Form As cWidgetForm)
    ApplyShellClientSize Win, Form, 1366, SHELL_DESIGN_H, "Shift+3 widescreen"
End Sub
