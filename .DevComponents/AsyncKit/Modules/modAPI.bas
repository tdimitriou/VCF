Attribute VB_Name = "modAPI"
Option Explicit

Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" ( _
    ByVal hModule As Long, _
    ByVal lpFileName As String, _
    ByVal nSize As Long) As Long

Private Declare Function GetModuleHandleEx Lib "kernel32" Alias "GetModuleHandleExA" ( _
    ByVal dwFlags As Long, _
    ByVal lpModuleName As Long, _
    ByRef phModule As Long) As Long

Private Const GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS As Long = &H4
Private Const GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT As Long = &H2
Private Const MAX_PATH As Long = 260

Private Declare Function GetCurrentThreadId Lib "kernel32" () As Long

Public Function DllPath() As String
    Dim hModule As Long
    Dim buf As String * MAX_PATH
    Dim nLen As Long
    
    If GetModuleHandleEx(GET_MODULE_HANDLE_EX_FLAG_FROM_ADDRESS Or _
                         GET_MODULE_HANDLE_EX_FLAG_UNCHANGED_REFCOUNT, _
                         AddressOf modAPI.DllPath, hModule) = 0 Then
        Exit Function
    End If
    
    nLen = GetModuleFileName(hModule, buf, MAX_PATH)
    If nLen <= 0 Then Exit Function
    
    Dim p As String
    p = Left$(buf, nLen)
    If Len(p) = 0 Then Exit Function
    If Len(Dir$(p, vbNormal)) = 0 Then Exit Function
    
    DllPath = p
End Function


