VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StubCaseConverter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IValueConverter

Private Function IValueConverter_Convert(Value As Variant, Optional Parameter As Variant) As Variant
    Select Case LCase$(CStr(Parameter))
        Case "upper"
            IValueConverter_Convert = VCF.StringConversion.ToUpperCase(Value)
        Case "lower"
            IValueConverter_Convert = VCF.StringConversion.ToLowerCase(Value)
        Case "proper", "title"
            IValueConverter_Convert = VCF.StringConversion.ToProperCase(Value)
        Case Else
            IValueConverter_Convert = Value
    End Select
End Function

Private Function IValueConverter_ConvertBack(Value As Variant, Optional Parameter As Variant) As Variant
End Function
