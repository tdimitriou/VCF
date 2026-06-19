Imports System
Imports System.Drawing
Imports System.Windows.Forms

Public Class DesignerSelectionManager
    Private m_DesignSurface As DesignSurface
    Private m_SelectedElement As Object
    Private m_PropertyEditor As PropertyEditor

    Public Property SelectedElement() As Object
        Get
            Return m_SelectedElement
        End Get
        Set(ByVal value As Object)
            m_SelectedElement = value
            DrawSelection()

            ' Update property editor
            If m_PropertyEditor IsNot Nothing Then
                m_PropertyEditor.SelectedObject = value
            End If
        End Set
    End Property

    Public Property PropertyEditor() As PropertyEditor
        Get
            Return m_PropertyEditor
        End Get
        Set(ByVal value As PropertyEditor)
            m_PropertyEditor = value
        End Set
    End Property

    Public Property DesignSurface() As DesignSurface
        Get
            Return m_DesignSurface
        End Get
        Set(ByVal value As DesignSurface)
            m_DesignSurface = value
        End Set
    End Property

    Public Sub Initialize(ByVal designSurface As DesignSurface)
        m_DesignSurface = designSurface
    End Sub

    Public Sub DrawSelection(Optional ByVal g As Graphics = Nothing)
        If m_DesignSurface Is Nothing Then Return
        If m_DesignSurface.Container Is Nothing Then Return
        If m_SelectedElement Is Nothing Then Return

        Dim container As PictureBox = m_DesignSurface.Container
        Dim disposeGraphics As Boolean = (g Is Nothing)
        If disposeGraphics Then
            g = container.CreateGraphics()
        End If

        Try
            ' Get element bounds
            Dim left As Single, top As Single, width As Single, height As Single

            If TypeOf m_SelectedElement Is VCF.IUIElement Then
                Dim uiElement As VCF.IUIElement = DirectCast(m_SelectedElement, VCF.IUIElement)
                left = uiElement.DesignLeft
                top = uiElement.DesignTop
                width = uiElement.DesignWidth
                height = uiElement.DesignHeight
            Else
                Return
            End If

            ' Draw selection rectangle
            Using pen As New Pen(Color.Blue, 2)
                g.DrawRectangle(pen, left, top, width, height)
            End Using

            ' Draw resize handles (8 handles around the rectangle)
            DrawHandle(g, left, top)
            DrawHandle(g, left + width / 2, top)
            DrawHandle(g, left + width, top)
            DrawHandle(g, left + width, top + height / 2)
            DrawHandle(g, left + width, top + height)
            DrawHandle(g, left + width / 2, top + height)
            DrawHandle(g, left, top + height)
            DrawHandle(g, left, top + height / 2)
        Finally
            If disposeGraphics AndAlso g IsNot Nothing Then
                g.Dispose()
            End If
        End Try

        container.Invalidate()
    End Sub

    Private Sub DrawHandle(ByVal g As Graphics, ByVal x As Single, ByVal y As Single)
        Dim handleSize As Integer = 6

        g.FillRectangle(Brushes.Blue, x - handleSize / 2, y - handleSize / 2, handleSize, handleSize)
    End Sub

    Public Sub ClearSelection()
        m_SelectedElement = Nothing

        ' Clear property editor
        If m_PropertyEditor IsNot Nothing Then
            m_PropertyEditor.SelectedObject = Nothing
        End If

        If m_DesignSurface IsNot Nothing Then
            If m_DesignSurface.Container IsNot Nothing Then
                m_DesignSurface.Refresh()
            End If
        End If
    End Sub

    Public Function HitTest(ByVal x As Single, ByVal y As Single) As Object
        ' Search through design object's children to find element at X, Y
        If m_DesignSurface Is Nothing Then Return Nothing
        If m_DesignSurface.DesignObject Is Nothing Then Return Nothing

        Dim children As VCF.UIElementCollection = Nothing

        ' Get children collection
        If TypeOf m_DesignSurface.DesignObject Is VCF.IControl Then
            Dim control As VCF.IControl = DirectCast(m_DesignSurface.DesignObject, VCF.IControl)
            children = control.Children
        ElseIf TypeOf m_DesignSurface.DesignObject Is VCF.IWindow Then
            Dim win As VCF.IWindow = DirectCast(m_DesignSurface.DesignObject, VCF.IWindow)
            children = win.Base.Children
        Else
            Return Nothing
        End If

        If children Is Nothing Then Return Nothing

        ' Search children (reverse order to get top-most first)
        For i As Integer = children.Count - 1 To 0 Step -1
            Dim child As Object = children(i)

            If TypeOf child Is VCF.IUIElement Then
                Dim uiElement As VCF.IUIElement = DirectCast(child, VCF.IUIElement)

                ' Check if point is within bounds
                If x >= uiElement.DesignLeft AndAlso x <= uiElement.DesignLeft + uiElement.DesignWidth AndAlso
                   y >= uiElement.DesignTop AndAlso y <= uiElement.DesignTop + uiElement.DesignHeight Then
                    Return child
                End If
            End If
        Next

        Return Nothing
    End Function
End Class

