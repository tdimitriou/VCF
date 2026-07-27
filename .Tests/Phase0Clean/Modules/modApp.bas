Attribute VB_Name = "modApp"
Option Explicit

Public Sub Main()
    VCF.ClearCustomConstructor
    VCF.StrictXamlLoad = True
    modPhase0CleanBench.RunAll

    ' Idle so Main does not return until Stop. DrainHoldForce (after MsgBox) should
    ' have parked nested graphs and Released empty shells so Stop is safer.
    Debug.Print "Phase0Clean: idle (DoEvents) — press Run/End to stop"
    Do
        DoEvents
    Loop
End Sub
