Attribute VB_Name = "modHarnessLoadBench"
Option Explicit

' Phase 7f — timed POS fixture loads for Immediate-window P7d-LOAD-* gates.
' Enable via modHarnessConfig.ENABLE_LOAD_BENCH.

Public Sub TimedLoadView(ByVal Screen As HarnessScreen, ByVal GateId As String)
    Dim Started As Double
    Dim ElapsedMs As Long

    If Screen Is Nothing Then Exit Sub

    If Not modHarnessConfig.ENABLE_LOAD_BENCH Then
        Screen.LoadView
        Exit Sub
    End If

    Started = Timer
    Screen.LoadView
    ElapsedMs = ElapsedMsSince(Started)

    Debug.Print "[" & GateId & "] " & ElapsedMs & " ms  key=" & Screen.XamlResourceKey
End Sub

Public Sub LogTimedStage(ByVal GateId As String, ByVal Stage As String, ByVal ElapsedMs As Long)
    If Not modHarnessConfig.ENABLE_LOAD_BENCH Then Exit Sub
    Debug.Print "[" & GateId & "] " & Stage & "  " & ElapsedMs & " ms"
End Sub

Public Function ElapsedMsSince(ByVal Started As Double) As Long
    Dim Delta As Double
    Delta = Timer - Started
    If Delta < 0 Then Delta = Delta + 86400# ' midnight wrap
    ElapsedMsSince = CLng(Delta * 1000#)
End Function
