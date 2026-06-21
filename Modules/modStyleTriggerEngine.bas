Attribute VB_Name = "modStyleTriggerEngine"
Option Explicit

Public Sub ApplyActiveTriggers(ByVal Style As Style, ByVal Target As Object)
    If Style Is Nothing Then Exit Sub
    If Target Is Nothing Then Exit Sub

    If Not Style.BasedOn Is Nothing Then ApplyActiveTriggers Style.BasedOn, Target
    ApplyTriggersOnStyle Style, Target
End Sub

Private Sub ApplyTriggersOnStyle(ByVal Style As Style, ByVal Target As Object)
    Dim i As Long
    Dim Trig As PropertyTrigger

    For i = 0 To Style.TriggerCount - 1
        Set Trig = Style.TriggerAt(i)
        If Not Trig Is Nothing Then
            If IsPropertyTriggerActive(Target, Trig) Then ApplyTriggerSetters Target, Trig
        End If
    Next
End Sub

Public Function IsPropertyTriggerActive(ByVal Target As Object, ByVal Trig As PropertyTrigger) As Boolean
    If Trig Is Nothing Then Exit Function
    IsPropertyTriggerActive = TriggerValuesEqual(ReadTriggerPropertyValue(Target, Trig.PropertyName), Trig.TriggerValue)
End Function

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

    Select Case LCase$(Trim$(ExpectedSpec))
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

    If VarType(Actual) = vbBoolean Or VarType(Expected) = vbBoolean Then
        TriggerValuesEqual = (CBool(Actual) = CBool(Expected))
    ElseIf IsNumeric(Actual) And IsNumeric(Expected) Then
        TriggerValuesEqual = (CDbl(Actual) = CDbl(Expected))
    Else
        TriggerValuesEqual = (StrComp(CStr(Actual), CStr(Expected), vbTextCompare) = 0)
    End If
End Function

Private Sub ApplyTriggerSetters(ByVal Target As Object, ByVal Trig As PropertyTrigger)
    Dim i As Long

    For i = 0 To Trig.SetterCount - 1
        ApplySingleSetter Target, Trig.SetterKeyAt(i), Trig.SetterValueAt(i)
    Next
End Sub

Private Sub ApplySingleSetter(ByVal Target As Object, ByVal PropertyName As String, ByVal RawValue As Variant)
    Dim Dep As IDependencyObject
    Dim Value As Variant
    Dim Prop As DependencyProperty

    If TypeOf Target Is IDependencyObject Then Set Dep = Target

    Value = RawValue
    With New MarkupExtensions
        API.CopyVariable .GetMarkupValue(Value, Dep, PropertyName), Value
    End With

    If Not Dep Is Nothing Then
        If Dep.DependencyProperties.Exists(PropertyName) Then
            Set Prop = Dep.DependencyProperties.GetProperty(PropertyName)
            With New XAMLDependencyPropertyManager
                Dep.DependencyProperties.SetCurrentValue PropertyName, .GetPropertyValueFromString(Prop, Value)
            End With
            Exit Sub
        End If
    End If

    On Error Resume Next
    CallByName Target, PropertyName, VbLet, Value
    If Err.Number = 0 Then Exit Sub
    Err.Clear

    If TypeOf Target Is IControl Then CallByName Target.Widget, PropertyName, VbLet, Value
End Sub
