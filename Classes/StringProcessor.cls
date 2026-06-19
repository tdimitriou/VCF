VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StringProcessor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Parse(ByVal Text As String) As String

    ' Escape Backslash ('\') character
    Text = Replace$(Text, "\\", Chr(17))    ' Chr(17) = DeviceControl1 Character - Unused - Non Printable
    
    Text = Replace(Text, "\n", vbCrLf)      ' New Line
    Text = Replace(Text, "\t", vbTab)       ' Horizontal Tab
    
    ' Restore Backslash ('\') character
    Parse = Replace$(Text, Chr(17), "\")
    
End Function

Public Function Format(ByVal FormatString As String, ParamArray Args() As Variant) As String
   On Error Resume Next
   Dim i As Integer
    
    If IsMissing(Args) Then
        Format = FormatString
        Exit Function
    End If
        
    For i = 0 To UBound(Args)
        FormatString = Replace$(FormatString, "{" & i & "}", Args(i))
    Next
    Format = FormatString
End Function

Function Split(ByVal Value As String, _
                        Optional Delimiter, _
                        Optional Limit As Long = -1, _
                        Optional Compare As VbCompareMethod = vbBinaryCompare, _
                        Optional ByVal TrimParts As Boolean = False)
    Dim Parts
    Parts = VBA.Strings.Split(Value, Delimiter, Limit, Compare)
    
    If TrimParts Then
        Dim Index As Long
        For Index = LBound(Parts) To UBound(Parts)
        
            Parts(Index) = Trim$(Parts(Index))
        
        Next
    End If
    
    Split = Parts

End Function


