Imports System
Imports System.Drawing
Imports System.Windows.Forms
Imports System.Collections.Generic

Public Class DesignerToolbox
    Private m_Container As PictureBox
    Private m_DesignSurface As DesignSurface
    Private m_ControlTypes As List(Of String)
    Private m_DragControlType As String
    Private m_IsDragging As Boolean

    Public Property DesignSurface() As DesignSurface
        Get
            Return m_DesignSurface
        End Get
        Set(ByVal value As DesignSurface)
            m_DesignSurface = value
        End Set
    End Property

    Public Sub Initialize(ByVal container As PictureBox)
        m_Container = container
        m_ControlTypes = New List(Of String)()

        ' Setup container
        container.BackColor = Color.FromArgb(&HF0, &HF0, &HF0)

        ' Initialize control types
        InitializeControlTypes()

        ' Draw toolbox
        DrawToolbox()
    End Sub

    Private Sub InitializeControlTypes()
        ' Add available control types
        m_ControlTypes.Add("Button")
        m_ControlTypes.Add("Panel")
        m_ControlTypes.Add("Border")
        m_ControlTypes.Add("TextBlock")
        m_ControlTypes.Add("TextBox")
        m_ControlTypes.Add("Image")
        m_ControlTypes.Add("UniformGrid")
        m_ControlTypes.Add("ListView")
    End Sub

    Private Sub DrawToolbox()
        If m_Container Is Nothing Then Return
        If m_ControlTypes Is Nothing Then Return

        Using g As Graphics = m_Container.CreateGraphics()
            g.Clear(Color.FromArgb(&HF0, &HF0, &HF0))

            Dim y As Integer = 10
            Dim itemHeight As Integer = 30

            ' Draw title
            Using brush As New SolidBrush(Color.Black)
                g.DrawString("Toolbox", SystemFonts.DefaultFont, brush, 10, y)
            End Using
            y += 20

            ' Draw control type buttons
            For Each controlType As String In m_ControlTypes
                ' Draw button background
                g.FillRectangle(Brushes.White, 5, y, m_Container.Width - 10, itemHeight)
                Using pen As New Pen(Color.FromArgb(&HC0, &HC0, &HC0))
                    g.DrawRectangle(pen, 5, y, m_Container.Width - 10, itemHeight)
                End Using

                ' Draw text
                Using brush As New SolidBrush(Color.Black)
                    g.DrawString(controlType, SystemFonts.DefaultFont, brush, 10, y + 8)
                End Using

                y += itemHeight + 5
            Next
        End Using

        m_Container.Invalidate()
    End Sub

    Public Function GetControlTypeAt(ByVal x As Single, ByVal y As Single) As String
        ' Calculate which control type is at the given coordinates
        Dim itemHeight As Integer = 30
        Dim startY As Integer = 30 ' After title

        Dim index As Integer = CInt((y - startY) / (itemHeight + 5))

        If index >= 0 AndAlso index < m_ControlTypes.Count Then
            Return m_ControlTypes(index)
        End If

        Return String.Empty
    End Function

    Public Sub StartDrag(ByVal controlType As String)
        m_DragControlType = controlType
        m_IsDragging = True
    End Sub

    Public Sub EndDrag()
        m_IsDragging = False
        m_DragControlType = String.Empty
    End Sub

    Public ReadOnly Property IsDragging() As Boolean
        Get
            Return m_IsDragging
        End Get
    End Property

    Public ReadOnly Property DragControlType() As String
        Get
            Return m_DragControlType
        End Get
    End Property
End Class

