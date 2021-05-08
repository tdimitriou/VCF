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

Private Function GetApplicationStyle() As Style
    On Error Resume Next
    
    Static Style As Style
    Static IsSet As Boolean
    
    If Not IsSet Then Set Style = Application.Current.TryFindResource("Style")

    Set GetApplicationStyle = Style
    
End Function

Public Sub ApplyStyle(ByVal Target As Object)
    SetObjectProperties Target
End Sub

'CSEH: ErrMsgBox
Private Sub SetObjectProperties(Obj As Object)
    Dim Dep As IDependencyObject
    
    If Not TypeOf Obj Is VCF.Button Then Exit Sub
    
    
    Dim Attributes As cSortedDictionary
    
'    If Not GetApplicationStyle.StyleDictionaryExists("Button") Then
'        Set Attributes = Style.AddStyleDictionary("Button")
'        Attributes.Add "CornerRadius", "0"
'        Attributes.Add "GradientBackground", 1
'    Else
'        Set Attributes = Style("Button")
'    End If
    
    If TypeOf Obj Is IDependencyObject Then
        Set Dep = Obj
        SetDependencyProperties Dep, Attributes
    Else
        SetProperties Obj, Attributes
    End If
End Sub

Private Sub SetDependencyProperties(ByVal Dep As IDependencyObject, Attributes As cSortedDictionary)
    Dim Key As String
    Dim Value As Variant
    Dim Obj As Object
    Dim Index As Long
    
    Set Obj = API.CObj(Dep)
    
    For Index = 0 To Attributes.Count - 1
        Key = Attributes.KeyByIndex(Index)
        
        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(Attributes(Key), Dep, Key), Value
        End With
    
        If Dep.DependencyProperties.Exists(Key) Then
            SetDependencyProperty Dep, Key, Value
        Else
            SetProperty Obj, Key, Value
        End If
    Next
End Sub

Private Sub SetProperties(ByVal Obj As Object, Attributes As cSortedDictionary)
    Dim Key As String
    Dim Index As Long
    
    For Index = 0 To Attributes.Count - 1
        SetProperty Obj, Key, Attributes(Key)
    Next
End Sub

Private Sub SetDependencyProperty(ByVal Obj As IDependencyObject, ByVal PropertyName As String, Value)
    Dim Prop As DependencyProperty
    
    Set Prop = Obj.DependencyProperties.GetProperty(PropertyName)
    
    With New XAMLDependencyPropertyManager
    
        Obj.DependencyProperties.SetValue PropertyName, .GetPropertyValueFromString(Prop, Value)
    
    End With

End Sub

Private Sub SetProperty(ByVal Obj As Object, ByVal Name As String, ByVal Value As String)
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
    
    If Err.Number = 0 Then Exit Sub
    
    SetProperty Obj.Widget, Name, Value
End Sub

