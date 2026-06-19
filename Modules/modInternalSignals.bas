Attribute VB_Name = "modInternalSignals"
Option Explicit

' --- Atomic Win32 Memory Operations ---
Private Declare Function InterlockedExchangeAdd Lib "kernel32" (ByRef Target As Long, ByVal Value As Long) As Long
Private Declare Function InterlockedExchange Lib "kernel32" (ByRef Target As Long, ByVal NewValue As Long) As Long

' --- RESTORED Win32 API Declaration for QueueUserAPC ---
' Declared as Public so it is globally accessible to the BackgroundWorker class
Public Declare Function QueueUserAPC Lib "kernel32" ( _
    ByVal pfnAPC As Long, _
    ByVal hThread As Long, _
    ByVal dwData As Long _
) As Long

' FIXED: Fixed buffer array configuration to guarantee absolute pointers in RAM
Private Const MAX_CONCURRENT_WORKERS As Long = 200

Private m_StopSignals(0 To MAX_CONCURRENT_WORKERS - 1) As Long
Private m_PauseSignals(0 To MAX_CONCURRENT_WORKERS - 1) As Long
Private m_InUse(0 To MAX_CONCURRENT_WORKERS - 1) As Boolean

' FIXED: Allocates immovable slots safely across multiple concurrent threads
Public Function AllocateSignalSlot(ByRef OutStopPtr As Long, ByRef OutPausePtr As Long) As Long
    Dim i As Long
    For i = 0 To MAX_CONCURRENT_WORKERS - 1
        If Not m_InUse(i) Then
            m_InUse(i) = True
            m_StopSignals(i) = 0
            m_PauseSignals(i) = 0
            
            ' Return solid memory addresses using VarPtr
            OutStopPtr = VarPtr(m_StopSignals(i))
            OutPausePtr = VarPtr(m_PauseSignals(i))
            
            AllocateSignalSlot = i
            Exit Function
        End If
    Next i
    
    Err.Raise 5, , "VCF Critical Error: Maximum concurrent background workers limit reached (" & MAX_CONCURRENT_WORKERS & ")."
End Function

Public Sub ReleaseSignalSlot(ByVal SlotIndex As Long)
    If SlotIndex >= 0 And SlotIndex < MAX_CONCURRENT_WORKERS Then
        m_InUse(SlotIndex) = False
        m_StopSignals(SlotIndex) = 0
        m_PauseSignals(SlotIndex) = 0
    End If
End Sub

' --- METHODS FOR STOP SIGNAL ---
Public Sub SetStopSignal(ByVal SlotIndex As Long)
    If SlotIndex >= 0 And SlotIndex < MAX_CONCURRENT_WORKERS Then
        InterlockedExchange m_StopSignals(SlotIndex), 1
    End If
End Sub

Public Function ReadStopSignal(ByVal pStopSignal As Long) As Boolean
    If pStopSignal = 0 Then Exit Function
    If InterlockedExchangeAdd(ByVal pStopSignal, 0) <> 0 Then
        ReadStopSignal = True
    End If
End Function

' --- METHODS FOR PAUSE SIGNAL ---
Public Sub SetPauseSignal(ByVal SlotIndex As Long, ByVal Value As Long)
    If SlotIndex >= 0 And SlotIndex < MAX_CONCURRENT_WORKERS Then
        InterlockedExchange m_PauseSignals(SlotIndex), Value
    End If
End Sub

Public Function ReadPauseSignal(ByVal pPauseSignal As Long) As Boolean
    If pPauseSignal = 0 Then Exit Function
    If InterlockedExchangeAdd(ByVal pPauseSignal, 0) <> 0 Then
        ReadPauseSignal = True
    End If
End Function

' --- RESTORED: The Sterile APC Callback Stub ---
' Located in a standard module to be fully compatible with the AddressOf operator.
' Executed asynchronously by the OS kernel to safely interrupt alertable SleepEx states.
Public Sub EmptyAPCCallback(ByVal dwParam As Long)
    ' Intentionally left blank by design.
    ' Do NOT inject VB6 UI routines or message loops here.
End Sub


