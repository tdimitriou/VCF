Imports System
Imports System.Text
Imports System.Collections.Generic

Public Class XAMLWriter
    Private m_IndentLevel As Integer
    Private m_IndentString As String
    Private m_IndentSize As Integer
    Private m_Output As StringBuilder
    Private m_Namespaces As Dictionary(Of String, String)

    '#Region Public Methods
    Public Function Save(ByVal obj As Object, Optional ByVal customClassName As String = "") As String
        m_IndentLevel = 0
        m_IndentSize = 4
        m_IndentString = New String(" "c, m_IndentSize)
        m_Output = New StringBuilder()
        m_Namespaces = New Dictionary(Of String, String)()

        ' Add default namespace
        m_Namespaces.Add("", "VCF")

        ' Start XML declaration
        m_Output.AppendLine("<?xml version=""1.0"" encoding=""utf-8""?>")

        ' Write root element
        WriteObject(obj, customClassName)

        Return m_Output.ToString()
    End Function

    Public Function SaveApp(ByVal appObject As VCF.IApplication, Optional ByVal customClassName As String = "") As String
        m_IndentLevel = 0
        m_IndentSize = 4
        m_IndentString = New String(" "c, m_IndentSize)
        m_Output = New StringBuilder()
        m_Namespaces = New Dictionary(Of String, String)()

        ' Add default namespace
        m_Namespaces.Add("", "VCF")

        ' Start XML declaration
        m_Output.AppendLine("<?xml version=""1.0"" encoding=""utf-8""?>")

        ' Write Application element
        WriteApplication(appObject, customClassName)

        Return m_Output.ToString()
    End Function
    '#End Region

    '#Region Private Methods - Application
    Private Sub WriteApplication(ByVal appObject As VCF.IApplication, Optional ByVal customClassName As String = "")
        Dim elementName As String = "Application"
        Dim attrs As New StringBuilder()
        Dim baseApp As VCF.Application = appObject.Base

        ' Add x:Class if custom class name provided
        If Not String.IsNullOrEmpty(customClassName) Then
            attrs.Append(" x:Class=""" & EscapeXml(customClassName) & """")
        End If

        ' Write opening tag
        WriteIndent()
        m_Output.Append("<" & elementName & attrs.ToString())

        ' Check if there are resources
        If baseApp.Resources.Count > 0 Then
            m_Output.AppendLine(">")
            m_IndentLevel += 1

            ' Write Application.Resources
            WriteIndent()
            m_Output.AppendLine("<Application.Resources>")
            m_IndentLevel += 1

            WriteResources(baseApp.Resources)

            m_IndentLevel -= 1
            WriteIndent()
            m_Output.AppendLine("</Application.Resources>")

            m_IndentLevel -= 1
            WriteIndent()
            m_Output.AppendLine("</" & elementName & ">")
        Else
            m_Output.AppendLine(" />")
        End If
    End Sub

    Private Sub WriteResources(ByVal resources As VCF.ObservableDictionary)
        For Each key As Object In resources.Keys
            Dim resourceKey As String = key.ToString()
            Dim value As Object = resources(key)

            If value Is Nothing Then Continue For

            Dim resourceType As String = value.GetType().Name

            ' Handle different resource types
            Select Case resourceType
                Case "Style"
                    WriteStyle(DirectCast(value, VCF.Style), resourceKey)
                Case "String"
                    WriteStringResource(DirectCast(value, String), resourceKey)
                Case Else
                    WriteObjectResource(value, resourceKey)
            End Select
        Next
    End Sub

    Private Sub WriteStyle(ByVal styleObj As VCF.Style, ByVal key As String)
        WriteIndent()
        m_Output.Append("<Style")

        If Not String.IsNullOrEmpty(styleObj.TargetType) Then
            m_Output.Append(" TargetType=""" & EscapeXml(styleObj.TargetType) & """")
        End If

        If Not String.IsNullOrEmpty(key) Then
            m_Output.Append(" x:Key=""" & EscapeXml(key) & """")
        End If

        If styleObj.Count > 0 Then
            m_Output.AppendLine(">")
            m_IndentLevel += 1

            For index As Integer = 0 To styleObj.Count - 1
                Dim setterPair As Object = styleObj.ItemAt(index)
                ' Note: KeyValuePair handling may need adjustment based on VCF implementation
                WriteIndent()
                m_Output.Append("<Setter")
                ' m_Output.Append(" Property=""" & EscapeXml(setterPair.Key) & """")
                ' m_Output.Append(" Value=""" & EscapeXml(PropertyValueToString(setterPair.Value)) & """")
                m_Output.AppendLine(" />")
            Next

            m_IndentLevel -= 1
            WriteIndent()
            m_Output.AppendLine("</Style>")
        Else
            m_Output.AppendLine(" />")
        End If
    End Sub

    Private Sub WriteStringResource(ByVal value As String, ByVal key As String)
        WriteIndent()
        m_Output.Append("<string")
        If Not String.IsNullOrEmpty(key) Then
            m_Output.Append(" x:Key=""" & EscapeXml(key) & """")
        End If
        m_Output.Append(" Value=""" & EscapeXml(value) & """")
        m_Output.AppendLine(" />")
    End Sub

    Private Sub WriteObjectResource(ByVal value As Object, ByVal key As String)
        ' Write as a regular object but with x:Key
        Dim savedIndent As Integer = m_IndentLevel
        m_IndentLevel = 0

        WriteObject(value, "", key)

        m_IndentLevel = savedIndent
    End Sub
    '#End Region

    '#Region Private Methods - Object Serialization
    Private Sub WriteObject(ByVal obj As Object, Optional ByVal customClassName As String = "", Optional ByVal resourceKey As String = "")
        If obj Is Nothing Then Return

        Dim elementName As String
        Dim attrs As New StringBuilder()
        Dim hasChildren As Boolean = False
        Dim depObj As VCF.IDependencyObject = Nothing
        Dim uiElement As VCF.IUIElement = Nothing
        Dim control As VCF.IControl = Nothing

        ' Get element name
        elementName = GetElementName(obj)

        ' Check for custom class name (x:Class)
        If Not String.IsNullOrEmpty(customClassName) Then
            attrs.Append(" x:Class=""" & EscapeXml(customClassName) & """")
        End If

        ' Check for resource key (x:Key)
        If Not String.IsNullOrEmpty(resourceKey) Then
            attrs.Append(" x:Key=""" & EscapeXml(resourceKey) & """")
        End If

        ' Check for Name property (x:Name)
        If TypeOf obj Is VCF.IUIElement Then
            uiElement = DirectCast(obj, VCF.IUIElement)
            If Not String.IsNullOrEmpty(uiElement.Name) Then
                attrs.Append(" x:Name=""" & EscapeXml(uiElement.Name) & """")
            End If
        End If

        ' Write dependency properties and regular properties
        If TypeOf obj Is VCF.IDependencyObject Then
            depObj = DirectCast(obj, VCF.IDependencyObject)
            WriteDependencyProperties(depObj, attrs)
        Else
            WriteRegularProperties(obj, attrs)
        End If

        ' Write attached properties
        If TypeOf obj Is VCF.IUIElement Then
            uiElement = DirectCast(obj, VCF.IUIElement)
            WriteAttachedProperties(uiElement, attrs)
        End If

        ' Check if object has children
        If TypeOf obj Is VCF.IControl Then
            control = DirectCast(obj, VCF.IControl)
            hasChildren = (control.Children.Count > 0)
        ElseIf TypeOf obj Is VCF.IWindow Then
            Dim win As VCF.IWindow = DirectCast(obj, VCF.IWindow)
            hasChildren = (win.Base.Children.Count > 0)
        End If

        ' Write opening tag
        WriteIndent()
        m_Output.Append("<" & elementName & attrs.ToString())

        If hasChildren Then
            m_Output.AppendLine(">")
            m_IndentLevel += 1

            ' Write children
            If TypeOf obj Is VCF.IControl Then
                control = DirectCast(obj, VCF.IControl)
                WriteChildren(control.Children)
            ElseIf TypeOf obj Is VCF.IWindow Then
                Dim win As VCF.IWindow = DirectCast(obj, VCF.IWindow)
                WriteChildren(win.Base.Children)
            End If

            m_IndentLevel -= 1
            WriteIndent()
            m_Output.AppendLine("</" & elementName & ">")
        Else
            m_Output.AppendLine(" />")
        End If
    End Sub

    Private Sub WriteDependencyProperties(ByVal depObj As VCF.IDependencyObject, ByRef attrs As StringBuilder)
        Dim bindingsList As VCF.List = Nothing

        ' Check if object has Bindings collection (via reflection or COM interop)
        Try
            Dim objType As Type = depObj.GetType()
            Dim bindingsProp As Reflection.PropertyInfo = objType.GetProperty("Bindings")
            If bindingsProp IsNot Nothing Then
                bindingsList = DirectCast(bindingsProp.GetValue(depObj, Nothing), VCF.List)
            End If
        Catch
            ' Property doesn't exist
        End Try

        ' Iterate through registered dependency properties
        For Each prop As VCF.DependencyProperty In depObj.DependencyProperties.RegisteredProperties
            Dim propName As String = prop.Name

            ' Skip certain properties that shouldn't be serialized
            If ShouldSkipProperty(propName) Then Continue For

            ' Check if this property has a binding
            Dim hasBinding As Boolean = False
            Dim binding As VCF.Binding = Nothing

            If bindingsList IsNot Nothing Then
                For Each b As Object In bindingsList
                    binding = TryCast(b, VCF.Binding)
                    If binding IsNot Nothing AndAlso binding.TargetProperty IsNot Nothing Then
                        If binding.TargetProperty.Name = propName Then
                            hasBinding = True
                            Exit For
                        End If
                    End If
                Next
            End If

            If hasBinding Then
                ' Write binding markup extension
                attrs.Append(" " & propName & "=""" & WriteBindingMarkup(binding) & """")
            Else
                ' Get property value
                Dim propValue As Object = prop.GetValue()

                ' Only write if value is not default/unset
                If Not IsDefaultValue(prop, propValue) Then
                    attrs.Append(" " & propName & "=""" & EscapeXml(PropertyValueToString(propValue)) & """")
                End If
            End If
        Next
    End Sub

    Private Sub WriteRegularProperties(ByVal obj As Object, ByRef attrs As StringBuilder)
        ' This would enumerate regular properties if needed
        ' For now, we focus on dependency properties
    End Sub

    Private Sub WriteAttachedProperties(ByVal uiElement As VCF.IUIElement, ByRef attrs As StringBuilder)
        For Each dictKey As Object In uiElement.AttachedProperties.Keys
            Dim dict As VCF.ObservableDictionary = DirectCast(uiElement.AttachedProperties(dictKey), VCF.ObservableDictionary)

            If dict IsNot Nothing Then
                For Each propKey As Object In dict.Keys
                    Dim propValue As Object = dict(propKey)
                    attrs.Append(" " & dictKey.ToString() & "." & propKey.ToString() & "=""" & EscapeXml(PropertyValueToString(propValue)) & """")
                Next
            End If
        Next
    End Sub

    Private Sub WriteChildren(ByVal children As VCF.UIElementCollection)
        For Each child As Object In children
            If TypeOf child Is VCF.IUIElement Then
                WriteObject(child, "", "")
            Else
                WriteObject(child, "", "")
            End If
        Next
    End Sub
    '#End Region

    '#Region Private Methods - Property Value Conversion
    Private Function PropertyValueToString(ByVal value As Object) As String
        If value Is Nothing Then Return ""

        ' Handle different types
        If TypeOf value Is VCF.Binding Then
            Return WriteBindingMarkup(DirectCast(value, VCF.Binding))
        ElseIf TypeOf value Is VCF.StaticResourceExtension Then
            Return WriteStaticResourceMarkup(DirectCast(value, VCF.StaticResourceExtension))
        ElseIf TypeOf value Is VCF.Thickness Then
            Return ThicknessToString(DirectCast(value, VCF.Thickness))
        ElseIf TypeOf value Is VCF.SolidColorBrush Then
            Return SolidColorBrushToString(DirectCast(value, VCF.SolidColorBrush))
        ElseIf TypeOf value Is VCF.Color Then
            Return ColorToString(DirectCast(value, VCF.Color))
        ElseIf TypeOf value Is Boolean Then
            Return If(DirectCast(value, Boolean), "True", "False")
        Else
            Return value.ToString()
        End If
    End Function

    Private Function ThicknessToString(ByVal thickness As VCF.Thickness) As String
        ' Check if all sides are equal
        If thickness.Left = thickness.Top AndAlso
           thickness.Top = thickness.Right AndAlso
           thickness.Right = thickness.Bottom Then
            Return thickness.Left.ToString()
        ElseIf thickness.Left = thickness.Right AndAlso
               thickness.Top = thickness.Bottom Then
            Return thickness.Left.ToString() & "," & thickness.Top.ToString()
        Else
            Return thickness.Left.ToString() & "," &
                   thickness.Top.ToString() & "," &
                   thickness.Right.ToString() & "," &
                   thickness.Bottom.ToString()
        End If
    End Function

    Private Function SolidColorBrushToString(ByVal brush As VCF.SolidColorBrush) As String
        Return brush.Color.ToString()
    End Function

    Private Function ColorToString(ByVal colorValue As Object) As String
        ' Convert color to hex format if possible
        Try
            Return VCF.modStaticClasses.Color.ToHtml(colorValue)
        Catch
            Return colorValue.ToString()
        End Try
    End Function

    Private Function WriteBindingMarkup(ByVal binding As VCF.Binding) As String
        Dim result As New StringBuilder()
        Dim parts As New List(Of String)()

        result.Append("{Binding")

        ' Add Path
        If Not String.IsNullOrEmpty(binding.Path) Then
            parts.Add("Path=" & binding.Path)
        End If

        ' Add Mode
        If binding.Mode <> VCF.BindingMode.Default Then
            Select Case binding.Mode
                Case VCF.BindingMode.OneWay
                    parts.Add("Mode=OneWay")
                Case VCF.BindingMode.OneTime
                    parts.Add("Mode=OneTime")
                Case VCF.BindingMode.OneWayToSource
                    parts.Add("Mode=OneWayToSource")
                Case VCF.BindingMode.TwoWay
                    parts.Add("Mode=TwoWay")
            End Select
        End If

        ' Add Converter
        If binding.Converter IsNot Nothing Then
            parts.Add("Converter=" & binding.Converter.GetType().Name)
        End If

        ' Add ConverterParameter
        If binding.ConverterParameter IsNot Nothing Then
            parts.Add("ConverterParameter=" & PropertyValueToString(binding.ConverterParameter))
        End If

        ' Add StringFormat
        If Not String.IsNullOrEmpty(binding.StringFormat) Then
            parts.Add("StringFormat=""" & EscapeXml(binding.StringFormat) & """")
        End If

        If parts.Count > 0 Then
            result.Append(" " & String.Join(", ", parts))
        End If

        result.Append("}")

        Return result.ToString()
    End Function

    Private Function WriteStaticResourceMarkup(ByVal staticResource As VCF.StaticResourceExtension) As String
        Dim result As New StringBuilder()

        result.Append("{StaticResource")

        If Not String.IsNullOrEmpty(staticResource.ResourceKey) Then
            result.Append(" " & staticResource.ResourceKey)
        End If

        result.Append("}")

        Return result.ToString()
    End Function
    '#End Region

    '#Region Private Methods - Helper Functions
    Private Function GetElementName(ByVal obj As Object) As String
        Dim typeNameStr As String = obj.GetType().Name

        ' Remove namespace prefix if present
        Dim lastDot As Integer = typeNameStr.LastIndexOf("."c)
        If lastDot >= 0 Then
            Return typeNameStr.Substring(lastDot + 1)
        Else
            Return typeNameStr
        End If
    End Function

    Private Function EscapeXml(ByVal text As String) As String
        If String.IsNullOrEmpty(text) Then Return ""

        Dim result As New StringBuilder(text)
        result.Replace("&", "&amp;")
        result.Replace("<", "&lt;")
        result.Replace(">", "&gt;")
        result.Replace("""", "&quot;")
        result.Replace("'", "&apos;")

        Return result.ToString()
    End Function

    Private Sub WriteIndent()
        For i As Integer = 1 To m_IndentLevel
            m_Output.Append(m_IndentString)
        Next
    End Sub

    Private Function ShouldSkipProperty(ByVal propName As String) As Boolean
        ' Properties that shouldn't be serialized
        Select Case propName.ToLower()
            Case "datacontext", "parent", "widget", "widgets"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Function IsDefaultValue(ByVal prop As VCF.DependencyProperty, ByVal value As Object) As Boolean
        ' Check if value is unset or default
        ' This is a simplified check - you may need to enhance this
        If value Is Nothing Then Return True

        ' Check against UnsetValue if available
        ' This would require access to DependencyPropertyMetadata

        Return False
    End Function
    '#End Region
End Class

