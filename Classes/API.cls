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

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hWnd As Long) As Long
Private Declare Function GetFocusAPI Lib "user32" Alias "GetFocus" () As Long

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


