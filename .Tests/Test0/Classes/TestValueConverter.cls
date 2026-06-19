VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TestValueConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValueConverter

'CSEH: ErrMsgBox
Private Function IValueConverter_Convert(Value As Variant, Optional Parameter As Variant) As Variant
    If Not IsNumeric(Value) Then Exit Function
    
    IValueConverter_Convert = 2 * Value
End Function

Private Function IValueConverter_ConvertBack(Value As Variant, Optional Parameter As Variant) As Variant
    '
End Function
