Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class DesignSurface
    Private m_Container As PictureBox
    Private m_DesignObject As Object
    Private m_SelectionManager As DesignerSelectionManager
    Private m_IsDesignMode As Boolean

    Public Property Container() As PictureBox
        Get
            Return m_Container
        End Get
        Set(ByVal value As PictureBox)
            m_Container = value
        End Set
    End Property

    Public Property DesignObject() As Object
        Get
            Return m_DesignObject
        End Get
        Set(ByVal value As Object)
            m_DesignObject = value

            ' If it's a Window, we need to host it in the container
            If TypeOf value Is VCF.IWindow Then
                Dim win As VCF.IWindow = DirectCast(value, VCF.IWindow)
                ' TODO: Host the window's form in the container
            End If

            Refresh()
        End Set
    End Property

    Public Property SelectionManager() As DesignerSelectionManager
        Get
            Return m_SelectionManager
        End Get
        Set(ByVal value As DesignerSelectionManager)
            m_SelectionManager = value
        End Set
    End Property

    Public Property IsDesignMode() As Boolean
        Get
            Return m_IsDesignMode
        End Get
        Set(ByVal value As Boolean)
            m_IsDesignMode = value
        End Set
    End Property

    Public Sub Initialize(ByVal container As PictureBox)
        m_Container = container
        m_IsDesignMode = True

        ' Setup container
        container.BackColor = Color.White

        ' Draw grid
        DrawGrid()
    End Sub

    Public Sub Refresh()
        If m_Container Is Nothing Then Return

        ' Redraw the design surface
        Using g As Graphics = m_Container.CreateGraphics()
            g.Clear(Color.White)
            DrawGrid(g)

            ' Render the design object and its children
            RenderDesignObject(g)

            ' Draw selection if any
            If m_SelectionManager IsNot Nothing Then
                If m_SelectionManager.SelectedElement IsNot Nothing Then
                    m_SelectionManager.DrawSelection(g)
                End If
            End If
        End Using

        m_Container.Invalidate()
    End Sub

    Private Sub RenderDesignObject(Optional ByVal g As Graphics = Nothing)
        If m_DesignObject Is Nothing Then Return
        If m_Container Is Nothing Then Return

        Dim disposeGraphics As Boolean = (g Is Nothing)
        If disposeGraphics Then
            g = m_Container.CreateGraphics()
        End If

        Try
            Dim children As VCF.UIElementCollection = Nothing

            ' Get children collection
            If TypeOf m_DesignObject Is VCF.IControl Then
                Dim control As VCF.IControl = DirectCast(m_DesignObject, VCF.IControl)
                children = control.Children
            ElseIf TypeOf m_DesignObject Is VCF.IWindow Then
                Dim win As VCF.IWindow = DirectCast(m_DesignObject, VCF.IWindow)
                children = win.Base.Children
            Else
                Return
            End If

            If children Is Nothing Then Return

            ' Render each child
            For Each child As Object In children
                If TypeOf child Is VCF.IUIElement Then
                    Dim uiElement As VCF.IUIElement = DirectCast(child, VCF.IUIElement)
                    RenderElement(uiElement, g)
                End If
            Next
        Finally
            If disposeGraphics AndAlso g IsNot Nothing Then
                g.Dispose()
            End If
        End Try
    End Sub

    Private Sub RenderElement(ByVal uiElement As VCF.IUIElement, ByVal g As Graphics)
        Dim left As Single = uiElement.DesignLeft
        Dim top As Single = uiElement.DesignTop
        Dim width As Single = uiElement.DesignWidth
        Dim height As Single = uiElement.DesignHeight

        ' Draw element rectangle
        Using pen As New Pen(Color.FromArgb(&HC0, &HC0, &HC0))
            g.DrawRectangle(pen, left, top, width, height)
        End Using

        ' Draw element type name
        Dim objType As String = uiElement.GetType().Name
        Using brush As New SolidBrush(Color.Black)
            g.DrawString(objType, SystemFonts.DefaultFont, brush, left + 2, top + 2)
        End Using

        ' If it's a control with children, render them
        If TypeOf uiElement Is VCF.IControl Then
            Dim control As VCF.IControl = DirectCast(uiElement, VCF.IControl)
            If control.Children.Count > 0 Then
                For Each child As Object In control.Children
                    If TypeOf child Is VCF.IUIElement Then
                        Dim childElement As VCF.IUIElement = DirectCast(child, VCF.IUIElement)
                        RenderElement(childElement, g)
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub DrawGrid(Optional ByVal g As Graphics = Nothing)
        If m_Container Is Nothing Then Return

        Dim disposeGraphics As Boolean = (g Is Nothing)
        If disposeGraphics Then
            g = m_Container.CreateGraphics()
        End If

        Try
            Dim gridSize As Integer = 10
            Using pen As New Pen(Color.FromArgb(&HE0, &HE0, &HE0))

                ' Draw vertical lines
                For i As Integer = 0 To m_Container.Width Step gridSize
                    g.DrawLine(pen, i, 0, i, m_Container.Height)
                Next

                ' Draw horizontal lines
                For i As Integer = 0 To m_Container.Height Step gridSize
                    g.DrawLine(pen, 0, i, m_Container.Width, i)
                Next
            End Using
        Finally
            If disposeGraphics AndAlso g IsNot Nothing Then
                g.Dispose()
            End If
        End Try
    End Sub

    Public Sub AddControl(ByVal controlType As String, ByVal x As Single, ByVal y As Single)
        If m_DesignObject Is Nothing Then Return

        Dim newControl As Object = CreateControl(controlType)

        If newControl Is Nothing Then Return

        ' Set position
        If TypeOf newControl Is VCF.IUIElement Then
            Dim uiElement As VCF.IUIElement = DirectCast(newControl, VCF.IUIElement)
            uiElement.DesignLeft = x
            uiElement.DesignTop = y
        End If

        ' Add to design object
        If TypeOf m_DesignObject Is VCF.IControl Then
            Dim control As VCF.IControl = DirectCast(m_DesignObject, VCF.IControl)
            control.Children.Add(newControl)
        ElseIf TypeOf m_DesignObject Is VCF.IWindow Then
            Dim win As VCF.IWindow = DirectCast(m_DesignObject, VCF.IWindow)
            win.Base.Children.Add(newControl)
        End If

        Refresh()
    End Sub

    Private Function CreateControl(ByVal controlType As String) As Object
        Dim obj As Object = Nothing

        Select Case controlType.ToLower()
            Case "button"
                obj = New VCF.Button()
            Case "panel"
                obj = New VCF.Panel()
            Case "border"
                obj = New VCF.Border()
            Case "textblock"
                obj = New VCF.TextBlock()
            Case "textbox"
                obj = New VCF.TextBox()
            Case "image"
                obj = New VCF.Image()
            Case "uniformgrid"
                obj = New VCF.UniformGrid()
            Case Else
                ' Try to create via COM
                Try
                    obj = Activator.CreateInstance(Type.GetTypeFromProgID("VCF." & controlType))
                Catch
                    ' Failed to create
                End Try
        End Select

        Return obj
    End Function
End Class

