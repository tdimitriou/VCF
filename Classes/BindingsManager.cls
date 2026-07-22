VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BindingsManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function CreateBindingFromMarkup(ByVal Target As IDependencyObject, _
                                        ByVal TargetProperty As String, _
                                        ByVal MarkupProperties As String) As Object
    Dim ArgList
    Dim PropList
    
    Dim Arg As Variant
    Dim vProp As Variant
    Dim Prop As KeyValuePair
    Dim Properties As cSortedDictionary
    
    If ParseMarkupPropertiesString(MarkupProperties, ArgList, PropList, Target) <> 0 Then Exit Function
    If IsEmpty(ArgList) And IsEmpty(PropList) Then Exit Function
    
    Set Properties = New_c.SortedDictionary(TextCompare)
    
    If Not IsEmpty(PropList) Then
        For Each vProp In PropList
            Prop = vProp
            Properties.Add Prop.Key, Prop.Value
        Next
    End If

    Set CreateBindingFromMarkup = CreateBinding(ArgList, Properties, Target, TargetProperty)
End Function

'CSEH: ErrMsgBox
Private Function CreateBinding(ArgList As Variant, _
                                    ByVal Properties As cSortedDictionary, _
                                    ByVal Target As IDependencyObject, _
                                    ByVal TargetProperty As String) As Binding
    
    Dim Obj As Binding
    Dim Source As Object
    Dim RS As RelativeSource
    
    Set Obj = New Binding
        
    If Not Target.DependencyProperties.Exists(TargetProperty) Then Exit Function
    
    Set Obj.TargetProperty = Target.DependencyProperties.GetProperty(TargetProperty)
        
    If Properties.Exists("Source") Then
        Set Source = Properties("Source")
    ElseIf Properties.Exists("RelativeSource") Then
        Set RS = Properties("RelativeSource")
        Set Source = ResolveRelativeSource(Target, RS)
        If Source Is Nothing Then Exit Function
    ElseIf Properties.Exists("ElementName") Then
        Set Source = ResolveElementName(Target, CStr(Properties("ElementName")))
        If Source Is Nothing Then Exit Function
    Else
        'Not going to happen for the objects currently provided by the VCF Library
        If Not Target.DependencyProperties.Exists("DataContext") Then Exit Function
    
        Set Source = Target.DependencyProperties.GetProperty("DataContext")
    End If
    
    Set Obj.Source = Source

    'Exit because there is no way to get the BindingPath
    If IsEmpty(ArgList) And (Not Properties.Exists("Path")) Then Exit Function
        
    'The first Argument is the Path (SourceProperty) of the Binding
    If Not IsEmpty(ArgList) Then Obj.Path = ArgList(0)
    
    If Properties.Exists("Path") Then Obj.Path = Properties("Path")
    If Properties.Exists("Converter") Then Set Obj.Converter = modConstructors.CreateInstance("", Properties("Converter"))
    If Properties.Exists("ConverterParameter") Then Obj.ConverterParameter = Properties("ConverterParameter")
    If Properties.Exists("StringFormat") Then Obj.StringFormat = Properties("StringFormat")
    If Properties.Exists("Mode") Then Obj.Mode = Properties("Mode")
    If Properties.Exists("UpdateSourceTrigger") Then
        Obj.UpdateSourceTrigger = ParseUpdateSourceTrigger(Properties("UpdateSourceTrigger"))
    End If
    If Properties.Exists("UpdateSourceDelay") Then
        If IsNumeric(Properties("UpdateSourceDelay")) Then
            Obj.UpdateSourceDelay = CLng(Properties("UpdateSourceDelay"))
        End If
    End If
    
    Set Obj.Target = Target
    API.CObj(Target).Bindings.Add Obj
        
    Set CreateBinding = Obj
End Function

Private Function ParseUpdateSourceTrigger(ByVal Value As Variant) As UpdateSourceTriggerMode
    Dim S As String

    If IsNumeric(Value) Then
        ParseUpdateSourceTrigger = CLng(Value)
        Exit Function
    End If

    S = LCase$(Trim$(CStr(Value)))
    Select Case S
        Case "propertychanged", "0", "ustpropertychanged"
            ParseUpdateSourceTrigger = ustPropertyChanged
        Case "lostfocus", "1", "ustlostfocus"
            ParseUpdateSourceTrigger = ustLostFocus
        Case "explicit", "2", "ustexplicit"
            ParseUpdateSourceTrigger = ustExplicit
        Case "default", "3", "updatesourcetriggerdefault"
            ParseUpdateSourceTrigger = UpdateSourceTriggerDefault
        Case Else
            ParseUpdateSourceTrigger = UpdateSourceTriggerDefault
    End Select
End Function

Private Function ParseMarkupPropertiesString(ByVal MarkupProperties As String, _
                                            ByRef ArgList As Variant, _
                                            ByRef PropList As Variant, _
                                            Optional ByVal Target As IDependencyObject = Nothing) As Long

Try:
    On Error GoTo Catch
    
    Dim Props
    Dim Prop
    Dim Key As String
    Dim Value
    Dim Body
    Dim Pair As KeyValuePair
    Dim Literals
    Dim Ext As MarkupExtensions
    
    Set Ext = New MarkupExtensions
        
    Ext.ParseLiterals MarkupProperties, Literals
        
    Props = Split(MarkupProperties, ",")
    
    For Each Prop In Props
        Body = Split(Trim$(Prop), "=")
        Ext.RestoreLiterals Body(0), Literals
        
        If UBound(Body) = 0 Then
            'It is an Argument
            
            If IsEmpty(ArgList) Then
                ReDim ArgList(0 To 0) As Variant
            Else
                ReDim Preserve ArgList(0 To UBound(ArgList) + 1) As Variant
            End If
                        
            API.CopyVariable Ext.GetMarkupValue(Body(0), Target), ArgList(UBound(ArgList))
        ElseIf UBound(Body) = 1 Then
            'It is a Property
            
            Ext.RestoreLiterals Body(1), Literals
            
            If IsEmpty(PropList) Then
                ReDim PropList(0 To 0) As Variant
            Else
                ReDim Preserve PropList(0 To UBound(PropList) + 1) As Variant
            End If
            
            Key = Trim$(Body(0))
            Value = Trim$(Body(1))
            
            Pair.Key = Key
            API.CopyVariable Ext.GetMarkupValue(Value, Target), Pair.Value
            PropList(UBound(PropList)) = Pair
            
        End If
    Next

    ParseMarkupPropertiesString = 0
    
    Exit Function

Catch:

    On Error Resume Next
    
    ArgList = Empty
    PropList = Empty
    ParseMarkupPropertiesString = Err.Number
End Function

