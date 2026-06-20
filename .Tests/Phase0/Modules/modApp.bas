Attribute VB_Name = "modApp"
Option Explicit

Public Sub Main()
    VCF.ClearCustomConstructor
    VCF.StrictXamlLoad = True
    modPhase0Bench.RunAll
End Sub
