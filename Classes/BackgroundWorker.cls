VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BackgroundWorker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' High-level Events for the UI Thread
Public Event ProgressChanged(ByVal Percent As Long, ByVal UserState As Variant)
Public Event WorkerEvent(ByVal Name As String, ByRef Args() As Variant)
Public Event RunWorkerCompleted(ByVal Cancelled As Boolean, Result As Variant, ErrorInfo As ErrorInfo)

' Core vbRichClient Thread Handler
Private WithEvents m_TH As cThreadHandler
Attribute m_TH.VB_VarHelpID = -1

' Shared Memory & Control Flags
Private m_SlotIndex As Long
Private m_pStopSignal As Long
Private m_pPauseSignal As Long
Private m_IsBusy As Boolean
Private m_ThreadKey As String

' .NET Framework Feature Properties
Private m_WorkerSupportsCancellation As Boolean
Private m_WorkerReportsProgress As Boolean

' --- Initialization ---
Private Sub Class_Initialize()
    ' Defaulting behavior to True for improved out-of-the-box usability in VB6
    m_WorkerSupportsCancellation = True
    m_WorkerReportsProgress = True
    m_SlotIndex = -1
End Sub

' --- Public API Properties ---

' Returns whether the background worker is currently running a task
Public Property Get IsBusy() As Boolean
    IsBusy = m_IsBusy
End Property

' Returns whether the background worker is currently paused
Public Property Get IsPaused() As Boolean
    If m_IsBusy And m_pPauseSignal <> 0 Then
        IsPaused = modInternalSignals.ReadPauseSignal(m_pPauseSignal)
    End If
End Property

' Gets or sets whether the worker supports cooperative cancellation
Public Property Get WorkerSupportsCancellation() As Boolean
    WorkerSupportsCancellation = m_WorkerSupportsCancellation
End Property
Public Property Let WorkerSupportsCancellation(ByVal Value As Boolean)
    m_WorkerSupportsCancellation = Value
End Property

' Gets or sets whether the worker can report progress updates to the UI
Public Property Get WorkerReportsProgress() As Boolean
    WorkerReportsProgress = m_WorkerReportsProgress
End Property
Public Property Let WorkerReportsProgress(ByVal Value As Boolean)
    m_WorkerReportsProgress = Value
End Property

' --- Public Control Methods ---

' Starts execution of a background task asynchronously
Public Sub RunWorkerAsync(ByVal Task As IBackgroundTask, ParamArray Args() As Variant)

    If m_IsBusy Then Err.Raise 5, , "The worker is already executing a task."
    If Task Is Nothing Then Err.Raise 5, , "Task object cannot be Nothing."
    
    m_IsBusy = True
    
    ' 1. Allocate a Thread-Safe slot in the Dynamic Array (Passing Stop and Pause pointers)
    On Error Resume Next
    m_SlotIndex = modInternalSignals.AllocateSignalSlot(m_pStopSignal, m_pPauseSignal)
    
    If Err.Number <> 0 Or m_SlotIndex = -1 Then
        ' Reset internal state to keep the component usable for future calls
        m_IsBusy = False
        m_SlotIndex = -1
        m_pStopSignal = 0
        m_pPauseSignal = 0
        
        ' Raise a clean, informative error back to the Caller Application
        Dim ErrDesc As String: ErrDesc = Err.Description
        On Error GoTo 0 ' Reset error routing
        
        If Len(ErrDesc) = 0 Then ErrDesc = "Maximum concurrent background workers limit reached (200)."
        Err.Raise 5, "BackgroundWorker.RunWorkerAsync", ErrDesc
        Exit Sub
    End If
    On Error GoTo 0 ' Restore standard error routing
    ' --------------------------------------------------------
        
    ' 2. Generate unique apartment thread key using ObjPtr
    m_ThreadKey = CStr(ObjPtr(Me))
    
    ' 3. Spin up the background thread side-by-side using RegFree activation
    Set m_TH = New_c.RegFree.ThreadObjectCreate(m_ThreadKey, modAPI.DllPath, "InternalWorker")
    
    ' 4. Offload the execution to the InternalWorker asynchronously
    Dim InternalArgs() As Variant: InternalArgs = Args
    m_TH.CallAsync "DoWork", Task, InternalArgs, m_pStopSignal, m_pPauseSignal

End Sub

' Signals the background thread to pause execution
Public Sub Pause()
    If Not m_IsBusy Or m_SlotIndex = -1 Then Exit Sub
    If Me.IsPaused Then Exit Sub
    
    modInternalSignals.SetPauseSignal m_SlotIndex, 1
End Sub

' Signals the background thread to resume execution and interrupts its alertable sleep state
Public Sub ResumeWorker()
    If Not m_IsBusy Or m_SlotIndex = -1 Then Exit Sub
    If Not Me.IsPaused Then Exit Sub
    
    ' 1. Lower the pause signal flag in shared memory
    modInternalSignals.SetPauseSignal m_SlotIndex, 0
    
    ' 2. Issue an asynchronous kernel APC wakeup call to instantly break out of SleepEx
    Call QueueUserAPC(AddressOf EmptyAPCCallback, m_TH.ThreadHandle, 0)
End Sub

' Requests cancellation of a pending background operation
Public Sub CancelAsync()
    ' Enforce .NET rules: Raise a runtime error if cancellation is invoked but disabled
    If Not m_WorkerSupportsCancellation Then
        Err.Raise 5, , "This BackgroundWorker instance is not configured to support cancellation."
    End If

    If m_IsBusy And m_SlotIndex <> -1 Then
        ' 1. Set the Stop Signal flag in the global shared array
        modInternalSignals.SetStopSignal m_SlotIndex
        
        ' 2. If the thread is currently paused, wake it up immediately to digest the Stop Signal
        If Me.IsPaused Then
            Call QueueUserAPC(AddressOf EmptyAPCCallback, m_TH.ThreadHandle, 0)
        End If
    End If
End Sub

' --- Internal Event Handlers ---

' Task Completion Handler triggered automatically by vbRichClient
Private Sub m_TH_MethodFinished(MethodName As String, Result As Variant, ErrString As String, ErrSource As String, ByVal ErrNumber As Long)
    Dim Cancelled As Boolean
    Dim ErrorInfo As ErrorInfo

    ' Check if the cancellation sequence was successfully triggered and finalized
    Cancelled = modInternalSignals.ReadStopSignal(m_pStopSignal)


    Set ErrorInfo = New ErrorInfo
    ' If a background crash occurred, encapsulate it into our custom container
    If ErrNumber <> 0 Then
        ErrorInfo.Populate ErrNumber, ErrString, ErrSource
    End If
    
    ' Thread-safe resource deallocation and memory cleanup
    modInternalSignals.ReleaseSignalSlot m_SlotIndex
    m_SlotIndex = -1
    m_pStopSignal = 0
    m_pPauseSignal = 0
    m_IsBusy = False
    Set m_TH = Nothing

    ' Dispatch final notification event back to the calling UI Thread
    RaiseEvent RunWorkerCompleted(Cancelled, Result, ErrorInfo)
End Sub

' Generic Cross-Thread Proxy routing Events and Progress notifications to the UI Apartment
Private Sub m_TH_ThreadEvent(MethodName As String, EventName As String, ByVal ParamCount As Long, P1 As Variant, P2 As Variant, P3 As Variant, P4 As Variant, P5 As Variant, P6 As Variant, P7 As Variant, P8 As Variant)
    If EventName = "WorkerEvent" Then
        Dim InternalEventName As String: InternalEventName = P1
        Dim Args() As Variant: Args = P2
        
        ' Handle standard Progress Reporting via proxy filtering
        If InternalEventName = "ReportProgress" Then
            If m_WorkerReportsProgress Then
                RaiseEvent ProgressChanged(Args(0), Args(1))
            End If
        Else
            ' Route all other developer-defined custom async events directly
            RaiseEvent WorkerEvent(InternalEventName, Args)
        End If
    End If
End Sub

' Class Destructor providing thread hardening features against memory leaks and orphan processes
Private Sub Class_Terminate()
    ' Bail out immediately if no asynchronous operation is active
    If Not m_IsBusy Or m_TH Is Nothing Then Exit Sub
    
    ' 1. Force the stop signal flag to True if not explicitly flagged already
    If Not modInternalSignals.ReadStopSignal(m_pStopSignal) Then
        modInternalSignals.SetStopSignal m_SlotIndex
    End If
    
    ' 2. Issue an immediate wakeup call if the class gets un-instantiated while paused
    If Me.IsPaused Then
        modInternalSignals.SetPauseSignal m_SlotIndex, 0
        Call QueueUserAPC(AddressOf EmptyAPCCallback, m_TH.ThreadHandle, 0)
    End If
    
    ' 3. Grant the thread a 3-second grace period for a cooperative, graceful shutdown
    Dim bCleanExit As Boolean
    bCleanExit = m_TH.WaitForEmptyJobQueue(3)
    
    ' 4. Hard Fallback: Force absolute termination if the thread becomes deadlocked
    If Not bCleanExit Then
        m_TH.TimeoutSecondsToHardTerminate = 1
        m_TH.CancelExecution
    End If
    
    ' 5. Safely recycle the memory slot index if MethodFinished was bypassed by hard termination
    If m_SlotIndex <> -1 Then
        modInternalSignals.ReleaseSignalSlot m_SlotIndex
    End If
End Sub


