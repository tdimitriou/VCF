VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "INIParser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long

Public Function ReadSetting(ByVal Filename As String, _
                            ByVal Section As String, _
                            ByVal Key As String, _
                            Optional ByVal Default As String) As String

    Dim Temp As String * 2048
    Dim Length As Integer

    Temp = Space$(2048)
    Length = GetPrivateProfileString(Section, Key, Default, Temp, 2047, Filename)
    ReadSetting = VBA.Left$(Temp, Length)

End Function

Public Function WriteSetting(ByVal Filename As String, _
                                ByVal Section As String, _
                                ByVal Key As String, _
                                ByVal Value As String) As Boolean

    Dim Index As Long

    For Index = 1 To Len(Value)
        
        If Mid$(Value, Index, 1) = vbCr Or Mid$(Value, Index, 1) = vbLf Then Mid$(Value, Index, 1) = ""
    
    Next

    WriteSetting = (WritePrivateProfileString(Section, Key, Value, Filename) <> 0)
    
End Function
