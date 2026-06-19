VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLWriter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_IndentLevel As Long
Private m_IndentString As String
Private m_IndentSize As Long
Private m_Output As cStringBuilder
Private m_Namespaces As cSortedDictionary

'#Region Public Methods
Public Function Save(ByVal Object As Object, Optional ByVal CustomClassName As String = "") As String
    On Error Resume Next
    
    m_IndentLevel = 0
    m_IndentSize = 4
    m_IndentString = Space$(m_IndentSize)
    Set m_Output = New_c.StringBuilder
    Set m_Namespaces = New_c.SortedDictionary(TextCompare)
    
    ' Add default namespace
    m_Namespaces.Add "", "VCF"
    
    ' Start XML declaration
    m_Output.AppendNL "<?xml version=""1.0"" encoding=""utf-8""?>"
    
    ' Write root element
    WriteObject Object, CustomClassName
    
    Save = m_Output.ToString
    
    Set m_Output = Nothing
    Set m_Namespaces = Nothing
End Function

Public Function SaveApp(ByVal AppObject As VCF.IApplication, Optional ByVal CustomClassName As String = "") As String
    On Error Resume Next
    
    m_IndentLevel = 0
    m_IndentSize = 4
    m_IndentString = Space$(m_IndentSize)
    Set m_Output = New_c.StringBuilder
    Set m_Namespaces = New_c.SortedDictionary(TextCompare)
    
    ' Add default namespace
    m_Namespaces.Add "", "VCF"
    
    ' Start XML declaration
    m_Output.AppendNL "<?xml version=""1.0"" encoding=""utf-8""?>"
    
    ' Write Application element
    WriteApplication AppObject, CustomClassName
    
    SaveApp = m_Output.ToString
    
    Set m_Output = Nothing
    Set m_Namespaces = Nothing
End Function
'#End Region

'#Region Private Methods - Application
Private Sub WriteApplication(ByVal AppObject As VCF.IApplication, Optional ByVal CustomClassName As String = "")
    Dim ElementName As String
    Dim Attrs As cStringBuilder
    Dim BaseApp As VCF.Application
    
    Set BaseApp = AppObject.Base
    Set Attrs = New_c.StringBuilder
    
    ElementName = "Application"
    
    ' Add x:Class if custom class name provided
    If Len(CustomClassName) > 0 Then
        Attrs.Append " x:Class=""" & EscapeXml(CustomClassName) & """"
    End If
    
    ' Write opening tag
    WriteIndent
    m_Output.Append "<" & ElementName & Attrs.ToString
    
    ' Check if there are resources
    If BaseApp.Resources.Count > 0 Then
        m_Output.Append ">"
        m_Output.AppendNL
        m_IndentLevel = m_IndentLevel + 1
        
        ' Write Application.Resources
        WriteIndent
        m_Output.Append "<Application.Resources>"
        m_Output.AppendNL
        m_IndentLevel = m_IndentLevel + 1
        
        WriteResources BaseApp.Resources
        
        m_IndentLevel = m_IndentLevel - 1
        WriteIndent
        m_Output.Append "</Application.Resources>"
        m_Output.AppendNL
        
        m_IndentLevel = m_IndentLevel - 1
        WriteIndent
        m_Output.Append "</" & ElementName & ">"
    Else
        m_Output.Append " />"
    End If
    
    m_Output.AppendNL
End Sub

Private Sub WriteResources(ByVal Resources As VCF.ObservableDictionary)
    Dim Key
    Dim Value
    Dim ResourceType As String
    Dim ResourceKey As String
    
    For Each Key In Resources.Keys
        ResourceKey = Key
        VCF.API.CopyVariable Resources(Key), Value
        
        If IsEmpty(Value) Then GoTo SKIP_RESOURCE
        If IsNull(Value) Then GoTo SKIP_RESOURCE
        If IsNothing(Value) Then GoTo SKIP_RESOURCE
        
        ResourceType = TypeName(Value)
        
        ' Handle different resource types
        Select Case ResourceType
            Case "Style"
                WriteStyle Value, ResourceKey
            Case "String"
                WriteStringResource Value, ResourceKey
            Case Else
                WriteObjectResource Value, ResourceKey
        End Select
        
SKIP_RESOURCE:
    Next
End Sub

Private Sub WriteStyle(ByVal StyleObj As VCF.Style, ByVal Key As String)
    WriteIndent
    m_Output.Append "<Style"
    
    If Len(StyleObj.TargetType) > 0 Then
        m_Output.Append " TargetType=""" & EscapeXml(StyleObj.TargetType) & """"
    End If
    
    If Len(Key) > 0 Then
        m_Output.Append " x:Key=""" & EscapeXml(Key) & """"
    End If
    
    If StyleObj.Count > 0 Then
        m_Output.Append ">"
        m_Output.AppendNL
        m_IndentLevel = m_IndentLevel + 1
        
        Dim Index As Long
        Dim SetterPair As KeyValuePair
        For Index = 0 To StyleObj.Count - 1
            SetterPair = StyleObj.ItemAt(Index)
            WriteIndent
            m_Output.Append "<Setter"
            m_Output.Append " Property=""" & EscapeXml(SetterPair.Key) & """"
            m_Output.Append " Value=""" & EscapeXml(PropertyValueToString(SetterPair.Value)) & """"
            m_Output.Append " />"
            m_Output.AppendNL
        Next
        
        m_IndentLevel = m_IndentLevel - 1
        WriteIndent
        m_Output.Append "</Style>"
    Else
        m_Output.Append " />"
    End If
    
    m_Output.AppendNL
End Sub

Private Sub WriteStringResource(ByVal Value As String, ByVal Key As String)
    WriteIndent
    m_Output.Append "<string"
    If Len(Key) > 0 Then
        m_Output.Append " x:Key=""" & EscapeXml(Key) & """"
    End If
    m_Output.Append " Value=""" & EscapeXml(Value) & """"
    m_Output.Append " />"
    m_Output.AppendNL
End Sub

Private Sub WriteObjectResource(ByVal Value As Object, ByVal Key As String)
    ' Write as a regular object but with x:Key
    Dim SavedIndent As Long
    SavedIndent = m_IndentLevel
    m_IndentLevel = 0
    
    WriteObject Value, "", Key
    
    m_IndentLevel = SavedIndent
End Sub
'#End Region

'#Region Private Methods - Object Serialization
Private Sub WriteObject(ByVal Object As Object, Optional ByVal CustomClassName As String = "", Optional ByVal ResourceKey As String = "")
    On Error Resume Next
    
    If Object Is Nothing Then Exit Sub
    
    Dim ElementName As String
    Dim Attrs As cStringBuilder
    Dim HasChildren As Boolean
    Dim DepObj As VCF.IDependencyObject
    Dim UIElement As VCF.IUIElement
    Dim Control As VCF.IControl
    Dim Binding As VCF.Binding
    Dim PropName As String
    
    Set Attrs = New_c.StringBuilder
    
    ' Get element name
    ElementName = GetElementName(Object)
    
    ' Check for custom class name (x:Class)
    If Len(CustomClassName) > 0 Then
        Attrs.Append " x:Class=""" & EscapeXml(CustomClassName) & """"
    End If
    
    ' Check for resource key (x:Key)
    If Len(ResourceKey) > 0 Then
        Attrs.Append " x:Key=""" & EscapeXml(ResourceKey) & """"
    End If
    
    ' Check for Name property (x:Name)
    If TypeOf Object Is VCF.IUIElement Then
        Set UIElement = Object
        If Len(UIElement.Name) > 0 Then
            Attrs.Append " x:Name=""" & EscapeXml(UIElement.Name) & """"
        End If
    End If
    
    ' Write dependency properties and regular properties
    If TypeOf Object Is VCF.IDependencyObject Then
        Set DepObj = Object
        WriteDependencyProperties DepObj, Attrs
    Else
        WriteRegularProperties Object, Attrs
    End If
    
    ' Write attached properties
    If TypeOf Object Is VCF.IUIElement Then
        Set UIElement = Object
        WriteAttachedProperties UIElement, Attrs
    End If
    
    ' Check if object has children
    If TypeOf Object Is VCF.IControl Then
        Set Control = Object
        HasChildren = (Control.Children.Count > 0)
    ElseIf TypeOf Object Is VCF.IWindow Then
        Dim Win As VCF.IWindow
        Set Win = Object
        HasChildren = (Win.Base.Children.Count > 0)
    End If
    
    ' Write opening tag
    WriteIndent
    m_Output.Append "<" & ElementName & Attrs.ToString
    
    If HasChildren Then
        m_Output.Append ">"
        m_Output.AppendNL
        m_IndentLevel = m_IndentLevel + 1
        
        ' Write children
        If TypeOf Object Is VCF.IControl Then
            Set Control = Object
            WriteChildren Control.Children
        ElseIf TypeOf Object Is VCF.IWindow Then
            Dim Win As VCF.IWindow
            Set Win = Object
            WriteChildren Win.Base.Children
        End If
        
        m_IndentLevel = m_IndentLevel - 1
        WriteIndent
        m_Output.Append "</" & ElementName & ">"
    Else
        m_Output.Append " />"
    End If
    
    m_Output.AppendNL
    
    If Err Then Err.Clear
End Sub

Private Sub WriteDependencyProperties(ByVal DepObj As VCF.IDependencyObject, ByRef Attrs As cStringBuilder)
    On Error Resume Next
    
    Dim Prop As VCF.DependencyProperty
    Dim PropValue
    Dim PropName As String
    Dim Binding As VCF.Binding
    Dim HasBinding As Boolean
    Dim BindingsList As VCF.List
    Dim Obj As Object
    
    Set Obj = VCF.API.CObj(DepObj)
    
    ' Check if object has Bindings collection
    If HasProperty(Obj, "Bindings") Then
        VCF.API.CopyVariable CallByName(Obj, "Bindings", VbGet), BindingsList
    End If
    
    ' Iterate through registered dependency properties
    For Each Prop In DepObj.DependencyProperties.RegisteredProperties
        PropName = Prop.Name
        
        ' Skip certain properties that shouldn't be serialized
        If ShouldSkipProperty(PropName) Then GoTo SKIP_PROP
        
        ' Check if this property has a binding
        HasBinding = False
        If Not BindingsList Is Nothing Then
            For Each Binding In BindingsList
                If Not Binding.TargetProperty Is Nothing Then
                    If Binding.TargetProperty.Name = PropName Then
                        HasBinding = True
                        Exit For
                    End If
                End If
            Next
        End If
        
        If HasBinding Then
            ' Write binding markup extension
            Attrs.Append " " & PropName & "=""" & WriteBindingMarkup(Binding) & """"
        Else
            ' Get property value
            VCF.API.CopyVariable Prop.GetValue, PropValue
            
            ' Only write if value is not default/unset
            If Not IsDefaultValue(Prop, PropValue) Then
                Attrs.Append " " & PropName & "=""" & EscapeXml(PropertyValueToString(PropValue)) & """"
            End If
        End If
        
SKIP_PROP:
    Next
    
    If Err Then Err.Clear
End Sub

Private Sub WriteRegularProperties(ByVal Object As Object, ByRef Attrs As cStringBuilder)
    On Error Resume Next
    
    ' This would enumerate regular properties if needed
    ' For now, we focus on dependency properties
    
    If Err Then Err.Clear
End Sub

Private Sub WriteAttachedProperties(ByVal UIElement As VCF.IUIElement, ByRef Attrs As cStringBuilder)
    On Error Resume Next
    
    Dim DictKey
    Dim Dict As VCF.ObservableDictionary
    Dim PropKey
    Dim PropValue
    
    For Each DictKey In UIElement.AttachedProperties.Keys
        VCF.API.CopyVariable UIElement.AttachedProperties(DictKey), Dict
        
        If Not Dict Is Nothing Then
            For Each PropKey In Dict.Keys
                VCF.API.CopyVariable Dict(PropKey), PropValue
                Attrs.Append " " & DictKey & "." & PropKey & "=""" & EscapeXml(PropertyValueToString(PropValue)) & """"
            Next
        End If
    Next
    
    If Err Then Err.Clear
End Sub

Private Sub WriteChildren(ByVal Children As VCF.UIElementCollection)
    On Error Resume Next
    
    Dim Child As Object
    Dim ChildUIElement As VCF.IUIElement
    
    For Each Child In Children
        If TypeOf Child Is VCF.IUIElement Then
            Set ChildUIElement = Child
            WriteObject Child, , ""
        Else
            WriteObject Child, , ""
        End If
    Next
    
    If Err Then Err.Clear
End Sub
'#End Region

'#Region Private Methods - Property Value Conversion
Private Function PropertyValueToString(ByVal Value As Variant) As String
    On Error Resume Next
    
    Dim Result As String
    
    If IsEmpty(Value) Then
        PropertyValueToString = ""
        Exit Function
    End If
    
    If IsNull(Value) Then
        PropertyValueToString = ""
        Exit Function
    End If
    
    If IsNothing(Value) Then
        PropertyValueToString = ""
        Exit Function
    End If
    
    ' Handle different types
    Select Case VarType(Value)
        Case vbObject
            If TypeOf Value Is VCF.Binding Then
                PropertyValueToString = WriteBindingMarkup(Value)
            ElseIf TypeOf Value Is VCF.StaticResourceExtension Then
                PropertyValueToString = WriteStaticResourceMarkup(Value)
            ElseIf TypeOf Value Is VCF.Thickness Then
                PropertyValueToString = ThicknessToString(Value)
            ElseIf TypeOf Value Is VCF.SolidColorBrush Then
                PropertyValueToString = SolidColorBrushToString(Value)
            ElseIf TypeOf Value Is VCF.Color Then
                PropertyValueToString = ColorToString(Value)
            Else
                ' For other objects, try to get a string representation
                Result = TypeName(Value)
            End If
            
        Case vbString
            PropertyValueToString = CStr(Value)
            
        Case vbBoolean
            If Value Then
                PropertyValueToString = "True"
            Else
                PropertyValueToString = "False"
            End If
            
        Case vbLong, vbInteger, vbByte
            PropertyValueToString = CStr(Value)
            
        Case vbSingle, vbDouble
            PropertyValueToString = CStr(Value)
            
        Case Else
            PropertyValueToString = CStr(Value)
    End Select
    
    If Err Then Err.Clear
End Function

Private Function ThicknessToString(ByVal Thickness As VCF.Thickness) As String
    On Error Resume Next
    
    Dim Result As String
    
    ' Check if all sides are equal
    If Thickness.Left = Thickness.Top And _
       Thickness.Top = Thickness.Right And _
       Thickness.Right = Thickness.Bottom Then
        ThicknessToString = CStr(Thickness.Left)
    ElseIf Thickness.Left = Thickness.Right And _
           Thickness.Top = Thickness.Bottom Then
        ThicknessToString = CStr(Thickness.Left) & "," & CStr(Thickness.Top)
    Else
        ThicknessToString = CStr(Thickness.Left) & "," & _
                           CStr(Thickness.Top) & "," & _
                           CStr(Thickness.Right) & "," & _
                           CStr(Thickness.Bottom)
    End If
    
    If Err Then Err.Clear
End Function

Private Function SolidColorBrushToString(ByVal Brush As VCF.SolidColorBrush) As String
    On Error Resume Next
    
    SolidColorBrushToString = Brush.Color
    
    If Err Then Err.Clear
End Function

Private Function ColorToString(ByVal ColorValue As Long) As String
    On Error Resume Next
    
    ' Convert color to hex format if possible
    ColorToString = VCF.modStaticClasses.Color.ToHtml(ColorValue)
    
    If Err Then Err.Clear
End Function

Private Function WriteBindingMarkup(ByVal Binding As VCF.Binding) As String
    On Error Resume Next
    
    Dim Result As String
    Dim Parts As cStringBuilder
    Set Parts = New_c.StringBuilder
    
    Result = "{Binding"
    
    ' Add Path
    If Len(Binding.Path) > 0 Then
        Parts.Append "Path=" & Binding.Path
    End If
    
    ' Add Mode
    If Binding.Mode <> VCF.BindingMode.Default Then
        If Parts.Length > 0 Then Parts.Append ","
        Select Case Binding.Mode
            Case VCF.BindingMode.OneWay
                Parts.Append "Mode=OneWay"
            Case VCF.BindingMode.OneTime
                Parts.Append "Mode=OneTime"
            Case VCF.BindingMode.OneWayToSource
                Parts.Append "Mode=OneWayToSource"
            Case VCF.BindingMode.TwoWay
                Parts.Append "Mode=TwoWay"
        End Select
    End If
    
    ' Add Converter
    If Not Binding.Converter Is Nothing Then
        If Parts.Length > 0 Then Parts.Append ","
        Parts.Append "Converter=" & TypeName(Binding.Converter)
    End If
    
    ' Add ConverterParameter
    If Not IsEmpty(Binding.ConverterParameter) Then
        If Parts.Length > 0 Then Parts.Append ","
        Parts.Append "ConverterParameter=" & PropertyValueToString(Binding.ConverterParameter)
    End If
    
    ' Add StringFormat
    If Len(Binding.StringFormat) > 0 Then
        If Parts.Length > 0 Then Parts.Append ","
        Parts.Append "StringFormat=""" & EscapeXml(Binding.StringFormat) & """"
    End If
    
    If Parts.Length > 0 Then
        Result = Result & " " & Parts.ToString
    End If
    
    Result = Result & "}"
    
    WriteBindingMarkup = Result
    
    If Err Then Err.Clear
End Function

Private Function WriteStaticResourceMarkup(ByVal StaticResource As VCF.StaticResourceExtension) As String
    On Error Resume Next
    
    Dim Result As String
    
    Result = "{StaticResource"
    
    If Len(StaticResource.ResourceKey) > 0 Then
        Result = Result & " " & StaticResource.ResourceKey
    End If
    
    Result = Result & "}"
    
    WriteStaticResourceMarkup = Result
    
    If Err Then Err.Clear
End Function
'#End Region

'#Region Private Methods - Helper Functions
Private Function GetElementName(ByVal Object As Object) As String
    On Error Resume Next
    
    Dim TypeNameStr As String
    Dim Namespace As String
    
    TypeNameStr = TypeName(Object)
    
    ' Remove namespace prefix if present
    If InStr(1, TypeNameStr, ".") > 0 Then
        GetElementName = Split(TypeNameStr, ".")(UBound(Split(TypeNameStr, ".")))
    Else
        GetElementName = TypeNameStr
    End If
    
    If Err Then GetElementName = TypeName(Object)
End Function

Private Function EscapeXml(ByVal Text As String) As String
    On Error Resume Next
    
    Dim Result As String
    
    Result = Text
    Result = Replace$(Result, "&", "&amp;")
    Result = Replace$(Result, "<", "&lt;")
    Result = Replace$(Result, ">", "&gt;")
    Result = Replace$(Result, """", "&quot;")
    Result = Replace$(Result, "'", "&apos;")
    
    EscapeXml = Result
    
    If Err Then EscapeXml = Text
End Function

Private Sub WriteIndent()
    Dim i As Long
    For i = 1 To m_IndentLevel
        m_Output.Append m_IndentString
    Next
End Sub

Private Function ShouldSkipProperty(ByVal PropName As String) As Boolean
    ' Properties that shouldn't be serialized
    Select Case LCase$(PropName)
        Case "datacontext", "parent", "widget", "widgets"
            ShouldSkipProperty = True
        Case Else
            ShouldSkipProperty = False
    End Select
End Function

Private Function IsDefaultValue(ByVal Prop As VCF.DependencyProperty, ByVal Value As Variant) As Boolean
    On Error Resume Next
    
    ' Check if value is unset or default
    ' This is a simplified check - you may need to enhance this
    IsDefaultValue = False
    
    If IsEmpty(Value) Then
        IsDefaultValue = True
        Exit Function
    End If
    
    If IsNull(Value) Then
        IsDefaultValue = True
        Exit Function
    End If
    
    If IsNothing(Value) Then
        IsDefaultValue = True
        Exit Function
    End If
    
    ' Check against UnsetValue if available
    ' This would require access to DependencyPropertyMetadata
    
    If Err Then Err.Clear
End Function

Private Function HasProperty(ByVal Object As Object, ByVal PropName As String) As Boolean
    On Error Resume Next
    
    Dim Dummy
    
    Call VCF.API.CopyVariable(CallByName(Object, PropName, VbGet), Dummy)
    
    HasProperty = (Err.Number = 0)
    
    If Err Then Err.Clear
End Function
'#End Region

