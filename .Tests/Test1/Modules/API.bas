Attribute VB_Name = "API"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function ObjFromPtr(ByVal Ptr As Long) As Object
    Dim Obj As Object
    CopyMemory Obj, Ptr, 4&
    Set ObjFromPtr = Obj
    CopyMemory Obj, 0&, 4&
End Function


