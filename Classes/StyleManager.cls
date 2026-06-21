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
    If TypeName(Target) <> Style.TargetType Then
        Err.Raise vbObjectError, TypeName(Me), "Target Type mismatch"
        Exit Sub
    End If
        
    SetObjectProperties Target, Style
    modStyleTriggerEngine.ApplyActiveTriggers Style, Target
    
    On Error Resume Next
    Target.Widget.Refresh
    If Err Then Err.Clear
    
End Sub

'CSEH: ErrMsgBox
Private Sub SetObjectProperties(Obj As Object, Style As Style)
    Dim Dep As IDependencyObject
        
    ' Set DependencyProperties is implemented using the 'SetCurrentValue'
    ' method allowing both Styling and individual property setters
    ' through XAML or code. However if a Style Property is not offered
    ' as a dependency property, the method will write the property value
    ' of the VCF object, or the contained WidgetBase, if no match is found.
    ' This will allow the inheritance of styles, however any property
    ' setters through XAML or code, will be overwritten each time the
    ' style is applied. The registration of a DependencyProperty, for any
    ' Style related property, is highly recommended for all VCF objects.
    If TypeOf Obj Is IDependencyObject Then
        Set Dep = Obj
        SetDependencyProperties Dep, Style
    Else
        SetProperties Obj, Style
    End If
End Sub

Private Sub SetDependencyProperties(ByVal Dep As IDependencyObject, Style As Style)
    Dim Obj As Object
        
    Set Obj = API.CObj(Dep)
    
    Dim Setter As KeyValuePair
    Dim Value
    
    Dim e
    For Each e In Style
        
        Setter = e
        
        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(Setter.Value, Nothing, Setter.Key), Value
        End With
    
        If Dep.DependencyProperties.Exists(CStr(Setter.Key)) Then
            SetDependencyProperty Dep, Setter.Key, Value
        Else
            SetProperty Obj, Setter.Key, Value
        End If
    Next
End Sub

Private Sub SetProperties(ByVal Obj As Object, Style As Style)
    Dim Setter As KeyValuePair
    Dim e
    
    Dim Value
    For Each e In Style
        Setter = e
        
        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(Setter.Value, Nothing, Setter.Key), Value
        End With
        
        SetProperty Obj, Setter.Key, Value
    Next
End Sub

Private Sub SetDependencyProperty(ByVal Obj As IDependencyObject, ByVal PropertyName As String, Value)
    Dim Prop As DependencyProperty
    
    Set Prop = Obj.DependencyProperties.GetProperty(PropertyName)
    
    With New XAMLDependencyPropertyManager
        
        Obj.DependencyProperties.SetCurrentValue PropertyName, .GetPropertyValueFromString(Prop, Value)
    
    End With

End Sub

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
'    Err.Clear
    If Not Recursive Then Exit Function
    
'    SetProperty = SetProperty(Obj.Widget, Name, Value, False)
'    If SetProperty Then Exit Function
'
'    If TypeOf Obj Is Window Then
'        SetProperty = SetProperty(Obj.Form, Name, Value, False)
'        If SetProperty Then Exit Function
'    End If
    
    Err.Clear
    If TypeOf Obj Is VCF.Window Then
        Dim ObjWin As VCF.Window
        Set ObjWin = Obj
        
        SetProperty ObjWin, Name, Value, False

        If Err.Number = 0 Then Exit Function
        Err.Clear
        SetProperty ObjWin.Form, Name, Value, False

        'If Err.Number = 0 Then Exit Function
        'Err.Clear
        SetProperty ObjWin.Form.WidgetRoot, Name, Value, False
    Else
        SetProperty Obj.Widget, Name, Value
    End If
    
    Err.Clear
End Function

