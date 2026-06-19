VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "InternalWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event WorkerEvent(ByVal Name As String, ByVal Args As Variant)

Private m_pStopSignal As Long
Private m_pPauseSignal As Long ' NEW ADDITION: Stores the incoming Pause Signal Pointer

Public Function DoWork(Task As IBackgroundTask, _
                        ByRef Args() As Variant, _
                        ByVal pStopSignal As Long, _
                        ByVal pPauseSignal As Long) As Variant
    
    On Error Resume Next
    
    m_pStopSignal = pStopSignal
    m_pPauseSignal = pPauseSignal ' NEW ADDITION: Receives the Pause pointer from the UI Thread
            
    If Not Task Is Nothing Then
        ' Passing "Me" (this InternalWorker instance) to the Job execution block,
        ' allowing ExecuteJob to invoke the new CheckAndWaitIfPaused method
        AssignVariable Task.Execute(Me, Args), DoWork
    End If

End Function

Private Sub AssignVariable(Source, Destination)
    
    If IsObject(Source) Then
        Set Destination = Source
    Else
        Destination = Source
    End If

End Sub

' --- NEW METHOD: The clean one-liner loop placeholder inside Job.ExecuteJob ---
Public Sub CheckAndWaitIfPaused()
    ' Thread-safe read of the Pause pointer via modInternalSignals
    If Not modInternalSignals.ReadPauseSignal(m_pPauseSignal) Then Exit Sub
    
    ' Infinite Alertable Sleep with 0% CPU overhead.
    ' If woken up via an APC call, it re-evaluates the thread-safe state flags.
    Do While modInternalSignals.ReadPauseSignal(m_pPauseSignal) And _
             Not modInternalSignals.ReadStopSignal(m_pStopSignal)
             
        New_c.SleepEx -1, True
    Loop
End Sub

Public Property Get CancellationPending() As Boolean
    CancellationPending = modInternalSignals.ReadStopSignal(m_pStopSignal)
End Property

Public Sub ReportProgress(ByVal Percent As Long, Optional ByVal UserState As Variant)
    RaiseEvent WorkerEvent("ReportProgress", Array(Percent, UserState))
End Sub

Public Sub RaiseEventAsync(ByVal EventName As String, ParamArray EventArgs() As Variant)
    Dim Args() As Variant: Args = EventArgs
    RaiseEvent WorkerEvent(EventName, Args)
End Sub


