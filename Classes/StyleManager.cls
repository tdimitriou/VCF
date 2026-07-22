VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StyleManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ApplyStyle(ByVal Style As Style, ByVal Target As Object)
    Dim InheritanceBatchStarted As Boolean
    Dim FailNum As Long
    Dim FailSrc As String
    Dim FailDesc As String
    Dim FailErl As Long

    On Error GoTo Handler

100 If Style Is Nothing Then Err.Raise 91, "StyleManager.ApplyStyle", "Style is Nothing"
102 If Target Is Nothing Then Err.Raise 91, "StyleManager.ApplyStyle", "Target is Nothing"
104 If TypeName(Target) <> Style.TargetType Then
        Err.Raise vbObjectError, TypeName(Me), "Target Type mismatch"
    End If

110 modInheritanceBatch.BeginInheritanceUpdate Target
    InheritanceBatchStarted = True

120 SetObjectProperties Target, Style
130 modStyleTriggerEngine.ApplyActiveTriggers Style, Target

140 modInheritanceBatch.EndInheritanceUpdate
    InheritanceBatchStarted = False

150 modControlTemplateEngine.ApplyControlTemplate Style, Target
160 PushTemplateContentAlignment Style, Target

170 modRenderCoalescer.RequestWidgetRefresh Target.Widget

    Exit Sub

Handler:
    FailNum = Err.Number
    FailSrc = Err.Source
    FailDesc = Err.Description
    FailErl = Erl
    modStyleApplyLog.LogError "StyleManager", "ApplyStyle", FailNum, FailDesc, FailErl
    If InheritanceBatchStarted Then
        On Error Resume Next
        modInheritanceBatch.EndInheritanceUpdate
        On Error GoTo 0
    End If
    If Len(FailSrc) = 0 Then
        Err.Raise FailNum, "StyleManager.ApplyStyle", FailDesc
    Else
        Err.Raise FailNum, FailSrc, FailDesc
    End If
End Sub

' Lookless-prep: ControlTemplate alignment slot ? target DPs via SetCurrentValue
Private Sub PushTemplateContentAlignment(ByVal St As Style, ByVal Target As Object)
    Dim Tmpl As ControlTemplate
    Dim Dep As IDependencyObject
    Dim HAlign As Long
    Dim VAlign As Long

    On Error GoTo Handler

200 If St Is Nothing Then Exit Sub
202 If Target Is Nothing Then Exit Sub

210 Set Tmpl = St.Template
212 If Tmpl Is Nothing Then Exit Sub
214 If Not Tmpl.HasContentAlignmentMarker Then Exit Sub

220 HAlign = Tmpl.ContentHorizontalAlignment
222 VAlign = Tmpl.ContentVerticalAlignment

230 If TypeOf Target Is IDependencyObject Then
240     Set Dep = Target
250     If Dep.DependencyProperties.Exists("HorizontalContentAlignment") Then
260         Dep.DependencyProperties.SetCurrentValue "HorizontalContentAlignment", HAlign
        End If
270     If Dep.DependencyProperties.Exists("VerticalContentAlignment") Then
280         Dep.DependencyProperties.SetCurrentValue "VerticalContentAlignment", VAlign
        End If
    End If

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "StyleManager", "PushTemplateContentAlignment"
End Sub

Private Sub SetObjectProperties(Obj As Object, Style As Style)
    Dim Dep As IDependencyObject

    On Error GoTo Handler

400 If TypeOf Obj Is IDependencyObject Then
410     Set Dep = Obj
420     SetDependencyProperties Dep, Style
    Else
430     SetProperties Obj, Style
    End If

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "StyleManager", "SetObjectProperties"
End Sub

Private Sub SetDependencyProperties(ByVal Dep As IDependencyObject, Style As Style)
    Dim Obj As Object
    Dim Setter As KeyValuePair
    Dim Value
    Dim e

    On Error GoTo Handler

500 Set Obj = API.CObj(Dep)

510 For Each e In Style
520     Setter = e
530     With New MarkupExtensions
540         API.CopyVariable .GetMarkupValue(Setter.Value, Nothing, Setter.Key), Value
        End With

550     If Dep.DependencyProperties.Exists(CStr(Setter.Key)) Then
560         SetDependencyProperty Dep, Setter.Key, Value
        Else
570         SetProperty Obj, Setter.Key, Value
        End If
580 Next

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "StyleManager", "SetDependencyProperties"
End Sub

Private Sub SetProperties(ByVal Obj As Object, Style As Style)
    Dim Setter As KeyValuePair
    Dim e
    Dim Value

    On Error GoTo Handler

600 For Each e In Style
610     Setter = e
620     With New MarkupExtensions
630         API.CopyVariable .GetMarkupValue(Setter.Value, Nothing, Setter.Key), Value
        End With
640     SetProperty Obj, Setter.Key, Value
650 Next

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "StyleManager", "SetProperties"
End Sub

Private Sub SetDependencyProperty(ByVal Obj As IDependencyObject, ByVal PropertyName As String, Value)
    Dim Prop As DependencyProperty

    On Error GoTo Handler

700 Set Prop = Obj.DependencyProperties.GetProperty(PropertyName)
710 If Prop Is Nothing Then Err.Raise 91, "StyleManager.SetDependencyProperty", "Property not found: " & PropertyName

720 With New XAMLDependencyPropertyManager
730     Obj.DependencyProperties.SetCurrentValue PropertyName, .GetPropertyValueFromString(Prop, Value)
    End With

    Exit Sub

Handler:
    modStyleApplyLog.LogErrorAndReraise "StyleManager", "SetDependencyProperty"
End Sub

' Intentional soft probe for non-DP property names — keep Resume Next (not part of fail-fast chain).
Private Function SetProperty(ByVal Obj As Object, ByVal Name As String, ByVal Value As String, Optional ByVal Recursive As Boolean = True) As Boolean
    On Error Resume Next

    Err.Clear

    If LCase$(Name) = "imagekey" Then
        If Not TypeOf Obj Is cWidgetBase Then
            With New XAMLImagePropertyManager
                Value = .LoadImage(Value)
            End With
        End If
    End If

    CallByName Obj, Name, VbLet, Value

    SetProperty = (Err.Number = 0)
    If SetProperty Then Exit Function
    If Not Recursive Then Exit Function

    Err.Clear
    If TypeOf Obj Is VCF.Window Then
        Dim ObjWin As VCF.Window
        Set ObjWin = Obj

        SetProperty ObjWin, Name, Value, False

        If Err.Number = 0 Then Exit Function
        Err.Clear
        SetProperty ObjWin.Form, Name, Value, False

        SetProperty ObjWin.Form.WidgetRoot, Name, Value, False
    Else
        SetProperty Obj.Widget, Name, Value
    End If

    Err.Clear
End Function
