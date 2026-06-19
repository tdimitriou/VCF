# AsyncKit - Advanced Enterprise Multithreading Component for VB6

## Overview
AsyncKit is a lightweight, high-performance, and "fool-proof" asynchronous task execution framework designed for VB6. It emulates the design pattern of the .NET BackgroundWorker while strictly adhering to the Single-Threaded Apartment (STA) memory isolation model enforced by the VB6 Runtime and the vbRichClient library.

By combining low-level Win32 Kernel APIs (QueueUserAPC, atomic memory operations) with vbRichClient’s robust threading infrastructure, AsyncKit provides safe execution, progress reporting, dynamic custom event routing, cooperative cancellation, computational data returns, and true 0% CPU overhead pause/resume capabilities without freezing the main User Interface (UI).

---

## Architectural Layout & Threading Model

The framework isolates complex Win32 handling into a standalone ActiveX DLL (AsyncKit.dll), providing client applications with a clean, high-level API.

### Execution Flowchart

 [ UI THREAD (Main Process) ]              [ WORKER THREAD (Isolated STA) ]
 ────────────────────────────              ────────────────────────────────
     BackgroundWorker
            │
            ├─► 1. RunWorkerAsync(Task) 
            │      │
            │      ├─► Allocates fixed shared memory slots for Stop/Pause flags.
            │      │   [Safeguard protects against slot exhaustion (Max 200)]
            │      │
            │      └─► Launches cThreadHandler.CallAsync("DoWork")
            │                                     │
            │   =================== COM SERIALIZATION ===================
            │                                     ▼
            │                               InternalWorker
            │                                     │
            │                                     ├─► 2. Binds Me as "Bridge"
            │                                     │
            │                                     └─► 3. Executes TargetTask.Execute()
            │                                                │
            │   ◄───────── ThreadEvent ──────────────────────┤ (Heavy loop begins)
            │         (Marshalling UI Proxy)                 │
            │                                                ├─► CheckAndWaitIfPaused
            │                                                │   (SleepEx -1, Alertable)
            │                                                │
            ├─► 4. Pause() ──────────────────────────────────┤ Updates shared flag
            │                                                │
            ├─► 5. ResumeWorker() ──► QueueUserAPC ──────────┼─► Interrupts SleepEx
            │                                                │
            ├─► 6. CancelAsync() ───► QueueUserAPC ──────────┼─► Interrupts SleepEx
            │                                                │   & Signals Termination
            │                                                │
            │   ◄───────── MethodFinished ───────────────────┴─► 7. Clean Exit
            ▼
   RunWorkerCompleted

### Component Breakdown
1. BackgroundWorker (Public, NotPersistable): Resides on the UI Thread. Handles component lifecycle, safety validation, dynamic slot safeguard, component termination, and routes multi-threaded notifications to the UI.
2. IBackgroundTask (Public, NotCreatable): A stateless abstract interface implemented by client applications to write custom background routines. (Instancing must be set to 2 - PublicNotCreatable).
3. InternalWorker (Public, Persistable): Instantiated automatically inside the background thread by the framework. Acts as an isolated execution host, implicit assignment boundary for diverse types, and cross-thread event proxy.
4. ErrorInfo (Public, NotPersistable): Encapsulates full COM exception details (Number, Description, Source) intercepted from the background thread crash. The 'Number' property is designated as the (Default) property, allowing native implicit Boolean evaluations in standard VB6 conditions.
5. modInternalSignals (Private Module): Pre-allocates a fixed thread-safe global shared memory region buffer for exactly 200 concurrent slots via Win32 atomic operations (InterlockedExchange, InterlockedExchangeAdd). Pointers are guaranteed to remain immovable in RAM, preventing multi-worker race conditions.
6. modAPI (Private Module): Hosts deep Win32 Kernel declarations and calculates side-by-side (RegFree) DLL registration paths.

---

## Prerequisites & Installation

1. Compiling the binary requires a valid reference to the vbRichClient library (RC5 or RC6) in your development environment.
2. Compile the ActiveX DLL project naming the binary AsyncKit.dll.
3. Procedure Attribute Mapping: Ensure that the 'Number' property in ErrorInfo.cls is set to (Default) via the VB6 'Tools ➔ Procedure Attributes' Advanced menu.
4. Registration-Free Deployment: Client projects can consume this library side-by-side (RegFree) without complex registry setups, as the component internally resolves its physical location at runtime via GetModuleHandleEx.

---

## Developer Implementation Workflow

To offload a long-running procedure to a background thread using AsyncKit, developers need to follow this minimal workflow configuration:

### 1. Minimal Background Task Implementation (MyAsyncTask.cls)
Create a Class Module, set its Persistable property to True (Mandatory), and implement the interface as a Function:

```vb
Option Explicit

Implements IBackgroundTask

Private Function IBackgroundTask_Execute(ByVal Bridge As AsyncKit.InternalWorker, Args() As Variant) As Variant
    Dim i As Long
    For i = 1 To 100
        Bridge.CheckAndWaitIfPaused
        If Bridge.CancellationPending Then Exit Function
        
        New_c.SleepEx 50 ' Simulating workload
        Bridge.ReportProgress i, "Processing step " & i
    Next i
    
    ' Return data payload back to the UI thread (Can be String, Array, Recordset, Collection, etc.)
    IBackgroundTask_Execute = "Task Execution Data Payload"
End Function

' Required properties for Persistable setting (Keep Empty)
Private Sub Class_WriteProperties(PropBag As PropertyBag): End Sub
Private Sub Class_ReadProperties(PropBag As PropertyBag): End Sub
Private Sub Class_InitProperties(): End Sub
```

### 2. Minimal Caller Implementation (Form1.frm)
Wire up the WithEvents component listener and call the asynchronous engine:

```vb
Option Explicit

Private WithEvents Worker As AsyncKit.BackgroundWorker

Private Sub Form_Load()
    Set Worker = New AsyncKit.BackgroundWorker
End Sub

Private Sub cmdStart_Click()
    Dim HeavyTask As New MyAsyncTask
    Worker.RunWorkerAsync HeavyTask
End Sub

Private Sub Worker_ProgressChanged(ByVal Percent As Long, ByVal UserState As Variant)
    lblStatus.Caption = CStr(UserState) & " (" & Percent & "%)"
End Sub

Private Sub Worker_RunWorkerCompleted(ByVal Cancelled As Boolean, ByVal Result As Variant, ByVal ErrorInfo As AsyncKit.ErrorInfo)
    ' Leveraging the Default Property behavior of the ErrorInfo object
    If ErrorInfo Then
        MsgBox "Failed: " & ErrorInfo.Description & " (" & ErrorInfo.Number & ")", vbCritical, ErrorInfo.Source
    ElseIf Cancelled Then
        MsgBox "Cancelled by user."
    Else
        ' Successfully retrieve the computed background data natively mapped inside the UI context
        MsgBox "Done! Result: " & CStr(Result), vbInformation, "Success"
    End If
End Sub
```

---

## Integration in this repo

**LinphoneLib** (`pos-v1/Devices/CallerID/LinphoneSuite/LinphoneLib`) embeds the same **BackgroundWorker** sources under **`Packages/BackgroundWorker/`** (no **AsyncKit.dll** at deploy; commit `0f413e8`). **`LinphoneClient`** + **`LinphoneWorker`** implement **`IBackgroundTask`**. **Demac.CallerID** uses **`LinphoneClient`** from built-in **`SipLine`**. POS UI: **`pos-v1/UI/Source/Controls/CallerList.ctl`** — see **`Devices/CallerID/POS_CALLERID_MIGRATION_PLAN.md`**.

Cross-stack fine-tuning and verification notes: `pos-v1/Devices/CallerID/FUTURE_FINE_TUNING_NOTES.md`.

---

## Technical Highlights & Safety Design

* Fixed Buffer Immovable Memory Architecture: In standard multithreaded environments, resizing dynamic shared lookup tables triggers pointer movement in RAM, resulting in dangling pointers and background race conditions. AsyncKit locks allocations inside a constant 200-slot buffer. Pointers obtained via VarPtr remain strictly stationary throughout the entire process runtime lifecycle, delivering absolute alignment for simultaneous threads.

* Safe Asynchronous Resuming: Rather than relying on raw thread hijacking hooks which crash the VB6 runtime thread-local storage engine (MSVMVM60.dll), AsyncKit targets the thread via QueueUserAPC passing a mathematically sterile, empty callback stub. This signals SleepEx to instantly return control back to the loop natively without altering critical processor registers.

* Anti-Zombie Thread Hardening: If an application form is forcefully destroyed while a task is sleeping or in a frozen state, the custom Class_Terminate destructor intercepts the reference drop. It clears the memory stack slots, un-pauses the underlying loop, applies a cooperative grace timeout exit sequence, and automatically executes a hard fallback CancelExecution sweep to prevent orphan zombie processing threads.
