Attribute VB_Name = "modApp"
Option Explicit

Public Sub Main()
    VCF.SetCustomConstructor Nothing
    VCF.StrictXamlLoad = True
    modPhase0Bench.RunAll
End Sub
