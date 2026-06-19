Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic
Imports System.Reflection

Public Class PropertyEditor
    Private m_Container As PictureBox
    Private m_SelectedObject As Object
    Private m_Properties As List(Of PropertyItem)
    Private m_ScrollPosition As Integer

    Public Property SelectedObject() As Object
        Get
            Return m_SelectedObject
        End Get
        Set(ByVal value As Object)
            m_SelectedObject = value
            Refresh()
        End Set
    End Property

    Public Sub Initialize(ByVal container As PictureBox)
        m_Container = container
        m_Properties = New List(Of PropertyItem)()

        ' Setup container
        container.BackColor = Color.White
        container.AutoScroll = True
    End Sub

    Public Sub Refresh()
        If m_Container Is Nothing Then Return

        m_Properties.Clear()

        If m_SelectedObject Is Nothing Then
            DrawEmpty()
            Return
        End If

        ' Collect properties
        CollectProperties()

        ' Draw properties
        DrawProperties()
    End Sub

    Private Sub DrawEmpty()
        Using g As Graphics = m_Container.CreateGraphics()
            Using brush As New SolidBrush(Color.Black)
                g.DrawString("No selection", SystemFonts.DefaultFont, brush, 10, 10)
            End Using
        End Using
        m_Container.Invalidate()
    End Sub

    Private Sub CollectProperties()
        m_Properties.Clear()

        If m_SelectedObject Is Nothing Then Return

        ' Get object type
        Dim objType As String = m_SelectedObject.GetType().Name

        ' Add type name
        Dim typeItem As New PropertyItem()
        typeItem.Name = "Type"
        typeItem.Value = objType
        m_Properties.Add(typeItem)

        ' If it's a UIElement, get common properties
        If TypeOf m_SelectedObject Is VCF.IUIElement Then
            Dim uiElement As VCF.IUIElement = DirectCast(m_SelectedObject, VCF.IUIElement)

            AddProperty("Name", uiElement.Name)
            AddProperty("DesignLeft", uiElement.DesignLeft)
            AddProperty("DesignTop", uiElement.DesignTop)
            AddProperty("DesignWidth", uiElement.DesignWidth)
            AddProperty("DesignHeight", uiElement.DesignHeight)
            AddProperty("Visibility", GetVisibilityString(uiElement.Visibility))
        End If

        ' If it's a DependencyObject, get dependency properties
        If TypeOf m_SelectedObject Is VCF.IDependencyObject Then
            Dim depObj As VCF.IDependencyObject = DirectCast(m_SelectedObject, VCF.IDependencyObject)
            CollectDependencyProperties(depObj)
        End If

        ' Try to get other common properties
        CollectCommonProperties()
    End Sub

    Private Sub CollectDependencyProperties(ByVal depObj As VCF.IDependencyObject)
        For Each prop As VCF.DependencyProperty In depObj.DependencyProperties.RegisteredProperties
            Dim propName As String = prop.Name

            ' Skip certain properties
            If ShouldSkipProperty(propName) Then Continue For

            ' Get value
            Dim propValue As Object = prop.GetValue()

            ' Format value
            AddProperty(propName, FormatPropertyValue(propValue))
        Next
    End Sub

    Private Sub CollectCommonProperties()
        Dim propNames() As String = {"Text", "Content", "Title", "Caption", "BackColor", "ForeColor", "FontSize", "FontBold"}

        For Each propName As String In propNames
            Dim propValue As Object = Nothing
            If GetPropertyValue(m_SelectedObject, propName, propValue) Then
                AddProperty(propName, FormatPropertyValue(propValue))
            End If
        Next
    End Sub

    Private Function GetPropertyValue(ByVal obj As Object, ByVal propName As String, ByRef value As Object) As Boolean
        Try
            Dim propInfo As PropertyInfo = obj.GetType().GetProperty(propName)
            If propInfo IsNot Nothing AndAlso propInfo.CanRead Then
                value = propInfo.GetValue(obj, Nothing)
                Return True
            End If
        Catch
            ' Property doesn't exist or can't be read
        End Try
        Return False
    End Function

    Private Sub AddProperty(ByVal name As String, ByVal value As Object)
        Dim prop As New PropertyItem()
        prop.Name = name
        prop.Value = value
        m_Properties.Add(prop)
    End Sub

    Private Function FormatPropertyValue(ByVal value As Object) As String
        If value Is Nothing Then
            Return "(nothing)"
        End If

        If TypeOf value Is String AndAlso String.IsNullOrEmpty(DirectCast(value, String)) Then
            Return "(empty)"
        End If

        Return value.ToString()
    End Function

    Private Function GetVisibilityString(ByVal visibility As Integer) As String
        Select Case visibility
            Case 0 ' Visible
                Return "Visible"
            Case 1 ' Hidden
                Return "Hidden"
            Case 2 ' Collapsed
                Return "Collapsed"
            Case Else
                Return visibility.ToString()
        End Select
    End Function

    Private Function ShouldSkipProperty(ByVal propName As String) As Boolean
        Select Case propName.ToLower()
            Case "datacontext", "parent", "widget", "widgets", "children"
                Return True
            Case Else
                Return False
        End Select
    End Function

    Private Sub DrawProperties()
        If m_Properties Is Nothing Then Return

        Using g As Graphics = m_Container.CreateGraphics()
            g.Clear(Color.White)

            Dim y As Integer = 10 - m_ScrollPosition
            Dim lineHeight As Integer = 20

            For Each prop As PropertyItem In m_Properties
                ' Draw property name
                Using brush As New SolidBrush(Color.Blue)
                    g.DrawString(prop.Name & ":", SystemFonts.DefaultFont, brush, 5, y)
                End Using

                ' Draw property value
                Using brush As New SolidBrush(Color.Black)
                    g.DrawString(prop.Value.ToString(), SystemFonts.DefaultFont, brush, 100, y)
                End Using

                y += lineHeight

                ' Check if we need to scroll
                If y > m_Container.Height Then Exit For
            Next
        End Using

        m_Container.Invalidate()
    End Sub

    Public Sub SetPropertyValue(ByVal propName As String, ByVal value As Object)
        If m_SelectedObject Is Nothing Then Return

        Try
            ' Try to set the property
            If TypeOf m_SelectedObject Is VCF.IDependencyObject Then
                Dim depObj As VCF.IDependencyObject = DirectCast(m_SelectedObject, VCF.IDependencyObject)

                If depObj.DependencyProperties.Exists(propName) Then
                    depObj.DependencyProperties.SetValue(propName, value)
                Else
                    ' Try regular property
                    Dim propInfo As PropertyInfo = m_SelectedObject.GetType().GetProperty(propName)
                    If propInfo IsNot Nothing AndAlso propInfo.CanWrite Then
                        propInfo.SetValue(m_SelectedObject, value, Nothing)
                    End If
                End If
            Else
                ' Try regular property
                Dim propInfo As PropertyInfo = m_SelectedObject.GetType().GetProperty(propName)
                If propInfo IsNot Nothing AndAlso propInfo.CanWrite Then
                    propInfo.SetValue(m_SelectedObject, value, Nothing)
                End If
            End If

            ' Refresh display
            Refresh()
        Catch ex As Exception
            MessageBox.Show("Error setting property: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Public Sub Scroll(ByVal delta As Integer)
        m_ScrollPosition += delta
        If m_ScrollPosition < 0 Then m_ScrollPosition = 0

        Refresh()
    End Sub
End Class

