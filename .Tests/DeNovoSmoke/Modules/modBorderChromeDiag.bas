Attribute VB_Name = "modBorderChromeDiag"
Option Explicit

Private m_StartTick As Single
Private m_Enabled As Boolean

Public Sub Reset()
    m_StartTick = Timer
    m_Enabled = modHarnessConfig.ENABLE_BORDER_CHROME_DIAG
    If m_Enabled Then
        Debug.Print "=== BORDER-DIAG begin " & Format$(Now, "hh:nn:ss") & " ==="
    End If
End Sub

Public Sub LogStage(ByVal Stage As String, ByVal Win As VCF.Window, ByVal Form As cWidgetForm)
    If Not m_Enabled Then Exit Sub
    
    Dim DpBorder As Long
    Dim FormBorder As Long
    Dim Visible As String
    
    On Error Resume Next
    DpBorder = CLng(Win.DependencyProperties.GetValue("BorderStyle"))
    If Form Is Nothing Then
        FormBorder = -1
        Visible = "n/a"
    Else
        FormBorder = Form.BorderStyle
        Visible = CStr(Form.Visible)
    End If
    
    Debug.Print "[BORDER-DIAG] +" & ElapsedMs & "ms  " & Stage & _
                "  DP=" & DpBorder & " (" & BorderStyleLabel(DpBorder) & ")" & _
                "  Form=" & FormBorder & " (" & BorderStyleLabel(FormBorder) & ")" & _
                "  Visible=" & Visible
End Sub

Public Sub LogChange(ByVal Stage As String, ByVal Previous As Long, ByVal Current As Long)
    If Not m_Enabled Then Exit Sub
    Debug.Print "[BORDER-DIAG] *** CHANGE +" & ElapsedMs & "ms  " & Stage & _
                "  Form.BorderStyle " & Previous & " -> " & Current & _
                " (" & BorderStyleLabel(Previous) & " -> " & BorderStyleLabel(Current) & ") ***"
End Sub

Public Sub LogNote(ByVal Message As String)
    If Not m_Enabled Then Exit Sub
    Debug.Print "[BORDER-DIAG] +" & ElapsedMs & "ms  " & Message
End Sub

Public Sub LogSummary(ByVal FormBorderBeforeShow As Long, ByVal FormBorderAtFirstPoll As Long, ByVal ChangeCountAfterShow As Long)
    If Not m_Enabled Then Exit Sub
    
    Debug.Print "[BORDER-DIAG] --- summary ---"
    Debug.Print "[BORDER-DIAG] Form.BorderStyle before Win.Show: " & FormBorderBeforeShow & " (" & BorderStyleLabel(FormBorderBeforeShow) & ")"
    Debug.Print "[BORDER-DIAG] Form.BorderStyle at first poll:   " & FormBorderAtFirstPoll & " (" & BorderStyleLabel(FormBorderAtFirstPoll) & ")"
    Debug.Print "[BORDER-DIAG] Changes detected after Show:     " & ChangeCountAfterShow
    
    If FormBorderBeforeShow = 2 And ChangeCountAfterShow = 0 Then
        Debug.Print "[BORDER-DIAG] VERDICT: Bordered (2) before show; no late chrome change - OK for first-frame border"
    ElseIf FormBorderBeforeShow <> 2 And FormBorderAtFirstPoll = 2 Then
        Debug.Print "[BORDER-DIAG] VERDICT: Border applied AFTER show became visible - likely borderless-first flash"
    ElseIf ChangeCountAfterShow > 0 Then
        Debug.Print "[BORDER-DIAG] VERDICT: BorderStyle changed after Show - review CHANGE lines above"
    Else
        Debug.Print "[BORDER-DIAG] VERDICT: Inconclusive - check stage lines above"
    End If
    
    Debug.Print "=== BORDER-DIAG end ==="
End Sub

Private Function ElapsedMs() As String
    ElapsedMs = CStr(CLng((Timer - m_StartTick) * 1000#))
End Function

Private Function BorderStyleLabel(ByVal Value As Long) As String
    Select Case Value
        Case 0:  BorderStyleLabel = "borderless"
        Case 1:  BorderStyleLabel = "fixed single"
        Case 2:  BorderStyleLabel = "sizable"
        Case 3:  BorderStyleLabel = "fixed dialog"
        Case 4:  BorderStyleLabel = "fixed tool"
        Case 5:  BorderStyleLabel = "sizable tool"
        Case -1: BorderStyleLabel = "?"
        Case Else: BorderStyleLabel = "style " & Value
    End Select
End Function
