Attribute VB_Name = "modAPI"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Public Function ObjFromPtr(ByVal Ptr As Long) As Object
    Dim Obj As Object
    CopyMemory Obj, Ptr, 4&
    Set ObjFromPtr = Obj
    CopyMemory Obj, 0&, 4&
End Function

Public Sub CopyVariable(Src, ByRef Dst)
    If IsObject(Src) Then
        Set Dst = Src
    Else
        Dst = Src
    End If
End Sub

Public Function CObj(ByVal Obj As Object) As Object
    Dim Unk As IUnknown
    
    Set Unk = CUnk(Obj)
    Set CObj = Unk
End Function

Public Function CUnk(ByVal Obj As IUnknown) As IUnknown
    Set CUnk = Obj
End Function

