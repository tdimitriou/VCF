VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   9735
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14460
   LinkTopic       =   "Form1"
   ScaleHeight     =   9735
   ScaleWidth      =   14460
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdResume 
      Caption         =   "Resume"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton cmdPause 
      Caption         =   "Pause"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.ListBox lstLog 
      Height          =   9030
      Left            =   1440
      TabIndex        =   2
      Top             =   120
      Width           =   12855
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Enabled         =   0   'False
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   1920
      Width           =   1215
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "Start"
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1215
   End
   Begin VB.Label lblStatus 
      Caption         =   "Ready"
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   9360
      Width           =   14175
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' Declarations section with WithEvents listener for the BackgroundWorker
Private WithEvents Worker As BackgroundWorker
Attribute Worker.VB_VarHelpID = -1

Private Sub Form_Load()
    ' Instantiate the BackgroundWorker component
    Set Worker = New BackgroundWorker
    
    ' Set initial UI state for control buttons
    cmdCancel.Enabled = False
    cmdPause.Enabled = False
    cmdResume.Enabled = False
End Sub

Private Sub cmdStart_Click()
    ' Clear and reset UI controls for the new asynchronous execution
    lstLog.Clear
    lblStatus.Caption = "Start..."
    
    ' Update button states using pure event-driven flow
    cmdStart.Enabled = False
    cmdCancel.Enabled = True
    cmdPause.Enabled = True    ' Enable pause since the worker is now running
    cmdResume.Enabled = False  ' Keep resume disabled until an actual pause occurs
    
    ' Trigger the background thread execution asynchronously
    Worker.RunWorkerAsync New MyAsyncTask, 200
End Sub

Private Sub cmdPause_Click()
    ' Check failsafe properties before execution
    If Not Worker.IsBusy Then Exit Sub
    If Worker.IsPaused Then Exit Sub
    
    ' Signal the shared memory to freeze the thread loop
    Worker.Pause
    
    ' Flip button states instantaneously on user click
    cmdPause.Enabled = False
    cmdResume.Enabled = True
    lblStatus.Caption = "Task Paused. Waiting for resume..."
End Sub

Private Sub cmdResume_Click()
    ' Check failsafe properties before execution
    If Not Worker.IsBusy Then Exit Sub
    If Not Worker.IsPaused Then Exit Sub
    
    ' Lower the pause flag and fire the Kernel APC wakeup signal
    Worker.ResumeWorker
    
    ' Flip button states instantaneously on user click
    cmdPause.Enabled = True
    cmdResume.Enabled = False
    lblStatus.Caption = "Resuming task execution..."
End Sub

Private Sub cmdCancel_Click()
    lblStatus.Caption = "Canceling. Please wait..."
    
    ' Request a safe, cooperative cancellation
    Worker.CancelAsync
    
    ' Disable pause/resume buttons immediately during the cancellation transition phase
    cmdPause.Enabled = False
    cmdResume.Enabled = False
End Sub

' --- Thread Notification Event Handlers ---

Private Sub Worker_ProgressChanged(ByVal Percent As Long, ByVal UserState As Variant)
    ' UI Safe update invoked by the underlying thread proxy routing
    lblStatus.Caption = CStr(UserState) & " (" & Percent & "%)"
End Sub

Private Sub Worker_RunWorkerCompleted(ByVal Cancelled As Boolean, Result As Variant, ErrorInfo As VCF.ErrorInfo)
    ' Task finalized: Restore core control buttons to their idle state
    cmdStart.Enabled = True
    cmdCancel.Enabled = False
    cmdPause.Enabled = False   ' Disable since the execution is over
    cmdResume.Enabled = False  ' Disable since the execution is over
    
    Dim Message As String
    ' Trap unhandled worker assembly crashes
    If ErrorInfo Then
        lblStatus.Caption = "Task Failed."
        Message = "Background task crashed: Error (" & ErrorInfo.Number & ") " & ErrorInfo.Description
        If Len(ErrorInfo.Source) Then Message = Message & " in " & ErrorInfo.Source
        MsgBox Message, vbCritical, "Error"
        Exit Sub
    End If
    
    Message = "RunWorkerCompleted, Result: " & Result
    lstLog.AddItem "[" & Time & "] " & Message
    lstLog.TopIndex = lstLog.NewIndex ' Dynamic Auto-scroll feature
    
    ' Evaluate whether the execution completed cleanly or via user interruption
    If Cancelled Then
        lblStatus.Caption = "Task cancelled by user."
        MsgBox "The background task was safely cancelled!", vbInformation
    Else
        lblStatus.Caption = "Task completed!"
        MsgBox "The background task completed successfully!", vbInformation
    End If

End Sub

Private Sub Worker_WorkerEvent(ByVal Name As String, ByRef Args() As Variant)
    ' Process developer-defined custom asynchronous events
    Select Case Name
        Case "ChunkCompleted"
            Dim FileIndex As Long: FileIndex = Args(0)
            Dim Message As String: Message = Args(1)
            
            lstLog.AddItem "[" & Time & "] " & Message
            lstLog.TopIndex = lstLog.NewIndex ' Dynamic Auto-scroll feature
    End Select
End Sub

