Attribute VB_Name = "modStyleTriggerEngine"
Option Explicit

Public Sub ApplyActiveTriggers(ByVal Style As Style, ByVal Target As Object)
    On Error GoTo Handler

100 If Style Is Nothing Then Exit Sub
102 If Target Is Nothing Then Exit Sub

110 If Not Style.BasedOn Is Nothing Then ApplyActiveTriggers Style.BasedOn, Target
120 ApplyTriggersOnStyle Style, Target

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modStyleTriggerEngine", "ApplyActiveTriggers"
End Sub

Private Sub ApplyTriggersOnStyle(ByVal Style As Style, ByVal Target As Object)
    Dim i As Long
    Dim Trig As PropertyTrigger

    On Error GoTo Handler

200 For i = 0 To Style.TriggerCount - 1
210     Set Trig = Style.TriggerAt(i)
220     If Not Trig Is Nothing Then
230         If IsPropertyTriggerActive(Target, Trig) Then ApplyTriggerSetters Target, Trig
        End If
240 Next

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modStyleTriggerEngine", "ApplyTriggersOnStyle"
End Sub

Public Function IsPropertyTriggerActive(ByVal Target As Object, ByVal Trig As PropertyTrigger) As Boolean
    On Error GoTo Handler

300 If Trig Is Nothing Then Exit Function
310 IsPropertyTriggerActive = TriggerValuesEqual(ReadTriggerPropertyValue(Target, Trig.PropertyName), Trig.TriggerValue)

    Exit Function

Handler:
    modStyleApplyLog.LogErrorAndReraise "modStyleTriggerEngine", "IsPropertyTriggerActive"
End Function

' Soft property probe — intentional Resume Next.
Public Function ReadTriggerPropertyValue(ByVal Target As Object, ByVal PropertyName As String) As Variant
    On Error Resume Next

    ReadTriggerPropertyValue = CallByName(Target, PropertyName, VbGet)
    If Err.Number = 0 Then Exit Function
    Err.Clear

    If TypeOf Target Is IControl Then
        ReadTriggerPropertyValue = CallByName(Target.Widget, PropertyName, VbGet)
    End If
End Function

Private Function TriggerValuesEqual(ByVal Actual As Variant, ByVal ExpectedSpec As String) As Boolean
    Dim Expected As Variant

    On Error GoTo Handler

400 Select Case LCase$(Trim$(ExpectedSpec))
        Case "true"
            Expected = True
        Case "false"
            Expected = False
        Case Else
            If IsNumeric(ExpectedSpec) Then
                Expected = Val(ExpectedSpec)
            Else
                Expected = ExpectedSpec
            End If
    End Select

410 If VarType(Actual) = vbBoolean Or VarType(Expected) = vbBoolean Then
        TriggerValuesEqual = (CBool(Actual) = CBool(Expected))
    ElseIf IsNumeric(Actual) And IsNumeric(Expected) Then
        TriggerValuesEqual = (CDbl(Actual) = CDbl(Expected))
    Else
        TriggerValuesEqual = (StrComp(CStr(Actual), CStr(Expected), vbTextCompare) = 0)
    End If

    Exit Function

Handler:
    modStyleApplyLog.LogErrorAndReraise "modStyleTriggerEngine", "TriggerValuesEqual"
End Function

Private Sub ApplyTriggerSetters(ByVal Target As Object, ByVal Trig As PropertyTrigger)
    Dim i As Long

    On Error GoTo Handler

500 For i = 0 To Trig.SetterCount - 1
510     ApplySingleSetter Target, Trig.SetterKeyAt(i), Trig.SetterValueAt(i)
520 Next

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modStyleTriggerEngine", "ApplyTriggerSetters"
End Sub

Private Sub ApplySingleSetter(ByVal Target As Object, ByVal PropertyName As String, ByVal RawValue As Variant)
    Dim Dep As IDependencyObject
    Dim Value As Variant
    Dim Prop As DependencyProperty

    On Error GoTo Handler

600 If TypeOf Target Is IDependencyObject Then Set Dep = Target

610 Value = RawValue
620 With New MarkupExtensions
630     API.CopyVariable .GetMarkupValue(Value, Dep, PropertyName), Value
    End With

640 If Not Dep Is Nothing Then
650     If Dep.DependencyProperties.Exists(PropertyName) Then
660         Set Prop = Dep.DependencyProperties.GetProperty(PropertyName)
670         With New XAMLDependencyPropertyManager
680             Dep.DependencyProperties.SetCurrentValue PropertyName, .GetPropertyValueFromString(Prop, Value)
            End With
            Exit Sub
        End If
    End If

    ' Soft widget/property probe
    On Error Resume Next
690 CallByName Target, PropertyName, VbLet, Value
    If Err.Number = 0 Then Exit Sub
    Err.Clear

700 If TypeOf Target Is IControl Then CallByName Target.Widget, PropertyName, VbLet, Value
    Err.Clear
    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "modStyleTriggerEngine", "ApplySingleSetter"
End Sub
