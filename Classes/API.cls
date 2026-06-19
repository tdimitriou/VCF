VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "API"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal length As Long)
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocusAPI Lib "user32" Alias "GetFocus" () As Long

Private Const VT_BYREF As Long = &H4000&

Public Function ObjFromPtr(ByVal Ptr As Long) As Object
    Dim Obj As Object
    
    CopyMemory Obj, Ptr, 4&
    Set ObjFromPtr = Obj
    CopyMemory Obj, 0&, 4&
End Function

Public Sub CopyVariable(Source, ByRef Destination)
    If IsObject(Source) Then
        Set Destination = Source
    Else
        Destination = Source
    End If
End Sub

Public Function CObj(ByVal Object As Object) As Object
    Dim Unk As IUnknown
    
    Set Unk = CUnk(Object)
    Set CObj = Unk
End Function

Public Function CUnk(ByVal Object As IUnknown) As IUnknown
    Set CUnk = Object
End Function

Public Function SetFocus(ByVal Handle As Long) As Long
    SetFocus = SetFocusAPI(Handle)
End Function

Public Function GetFocus() As Long
    GetFocus = GetFocusAPI
End Function

Public Function ArrayExists(ByRef Arr As Variant) As Boolean
  ArrayExists = (pArrPtr(Arr) <> 0)
End Function

' When declared in this way, the passed array is wrapped in a Variant/ByRef. It is not copied.
' Returns *SAFEARRAY, not **SAFEARRAY
Public Function pArrPtr(ByRef Arr As Variant) As Long
    'VarType lies to you, hiding important differences. Manual VarType here.
    Dim vt As Integer
    CopyMemory ByVal VarPtr(vt), ByVal VarPtr(Arr), Len(vt)
    
    If (vt And vbArray) <> vbArray Then
        'Err.Raise 5, , "Variant must contain an array"
        pArrPtr = 0
        Exit Function
    End If
    
    'see https://msdn.microsoft.com/en-us/library/windows/desktop/ms221627%28v=vs.85%29.aspx
    If (vt And VT_BYREF) = VT_BYREF Then
        'By-ref variant array. Contains **pparray at offset 8
        CopyMemory ByVal VarPtr(pArrPtr), ByVal VarPtr(Arr) + 8, Len(pArrPtr)  'pArrPtr = arr->pparray;
        CopyMemory ByVal VarPtr(pArrPtr), ByVal pArrPtr, Len(pArrPtr)          'pArrPtr = *pArrPtr;
    Else
        'Non-by-ref variant array. Contains *parray at offset 8
        CopyMemory ByVal VarPtr(pArrPtr), ByVal VarPtr(Arr) + 8, Len(pArrPtr)  'pArrPtr = arr->parray;
    End If
End Function

Function ArrayFind(SourceArray, Value, Optional ByRef Index As Long) As Boolean
    ArrayFind = ArrayWrapper(SourceArray).Find(Value, Index)
End Function


