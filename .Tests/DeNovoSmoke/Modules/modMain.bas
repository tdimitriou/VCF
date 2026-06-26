Attribute VB_Name = "modMain"
Option Explicit

Private Declare Function LoadLibraryA Lib "kernel32.dll" (ByVal lpLibFileName As String) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function InitCommonControlsEx Lib "comctl32.dll" (iccex As InitCommonControlsExStruct) As Boolean
Private Declare Sub InitCommonControls Lib "comctl32.dll" ()

Private Type InitCommonControlsExStruct
    lngSize As Long
    lngICC As Long
End Type

Public Sub Main()
    Dim iccex As InitCommonControlsExStruct
    Dim hMod As Long
    
    Const ICC_STANDARD_CLASSES As Long = &H4000&
    
    With iccex
        .lngSize = LenB(iccex)
        .lngICC = ICC_STANDARD_CLASSES
    End With
    
    On Error Resume Next
    hMod = LoadLibraryA("shell32.dll")
    InitCommonControlsEx iccex
    If Err Then
        InitCommonControls
        Err.Clear
    End If
    On Error GoTo 0
    
    modApp.Start
    
    If hMod Then FreeLibrary hMod
End Sub
