VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 1  'Persistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MyAsyncTask"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IBackgroundTask

Private Function IBackgroundTask_Execute(ByVal Bridge As VCF.InternalWorker, Args() As Variant) As Variant
On Error GoTo Catch

    Dim TotalCycles As Long
    TotalCycles = Args(0)
    
    Dim i As Long
    For i = 0 To TotalCycles - 1
        
        Bridge.CheckAndWaitIfPaused
        
        If Bridge.CancellationPending Then
            GoTo Finally
        End If
        
        New_c.SleepEx 50
        
        Dim Percent As Long
        Percent = CLng(((i + 1) / TotalCycles) * 100)
        Bridge.ReportProgress Percent, "Step " & i
        
        If i Mod 5 = 0 Then
            Bridge.RaiseEventAsync "ChunkCompleted", i, "Check OK"
        End If
        
    Next i

Finally:
    
    IBackgroundTask_Execute = i
    Exit Function
    
Catch:
    
    i = -1
    GoTo Finally
End Function
