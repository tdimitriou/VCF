Imports System
Imports System.Windows.Forms
Imports System.Drawing

Public Class frmDesigner
    Inherits Form

    Private m_DesignSurface As DesignSurface
    Private m_Toolbox As DesignerToolbox
    Private m_SelectionManager As DesignerSelectionManager
    Private m_PropertyEditor As PropertyEditor
    Private m_DragStartX As Single
    Private m_DragStartY As Single
    Private m_IsDragging As Boolean

    Private picToolbox As PictureBox
    Private picProperties As PictureBox
    Private picDesignSurface As PictureBox

    Public Sub New()
        InitializeComponent()
    End Sub

    Private Sub InitializeComponent()
        Me.Text = "VCF Designer"
        Me.Size = New Size(1200, 900)
        Me.StartPosition = FormStartPosition.WindowsDefaultLocation

        ' Create PictureBox controls
        picToolbox = New PictureBox()
        picToolbox.Dock = DockStyle.Left
        picToolbox.Width = 240
        picToolbox.BackColor = Color.FromArgb(&HF0, &HF0, &HF0)

        picProperties = New PictureBox()
        picProperties.Dock = DockStyle.Right
        picProperties.Width = 240
        picProperties.BackColor = Color.White

        picDesignSurface = New PictureBox()
        picDesignSurface.Dock = DockStyle.Fill
        picDesignSurface.BackColor = Color.White

        ' Add controls to form
        Me.Controls.Add(picDesignSurface)
        Me.Controls.Add(picProperties)
        Me.Controls.Add(picToolbox)

        ' Setup mouse events
        AddHandler picToolbox.MouseDown, AddressOf picToolbox_MouseDown
        AddHandler picToolbox.MouseMove, AddressOf picToolbox_MouseMove
        AddHandler picToolbox.MouseUp, AddressOf picToolbox_MouseUp

        AddHandler picDesignSurface.MouseDown, AddressOf picDesignSurface_MouseDown
        AddHandler picDesignSurface.MouseMove, AddressOf picDesignSurface_MouseMove
        AddHandler picDesignSurface.MouseUp, AddressOf picDesignSurface_MouseUp

        AddHandler Me.Load, AddressOf Form_Load
        AddHandler Me.Resize, AddressOf Form_Resize
        AddHandler Me.FormClosing, AddressOf Form_Unload
    End Sub

    Private Sub Form_Load(sender As Object, e As EventArgs)
        Try
            ' Initialize VCF Application for design-time
            Dim designApp As New DesignApplication()

            ' Initialize designer components
            m_DesignSurface = New DesignSurface()
            m_Toolbox = New DesignerToolbox()
            m_SelectionManager = New DesignerSelectionManager()

            ' Setup design surface
            m_DesignSurface.Initialize(picDesignSurface)
            m_DesignSurface.SelectionManager = m_SelectionManager

            ' Setup toolbox
            m_Toolbox.Initialize(picToolbox)
            m_Toolbox.DesignSurface = m_DesignSurface

            ' Setup selection manager
            m_SelectionManager.Initialize(m_DesignSurface)

            ' Setup property editor
            m_PropertyEditor = New PropertyEditor()
            m_PropertyEditor.Initialize(picProperties)

            ' Wire selection changed event
            m_SelectionManager.PropertyEditor = m_PropertyEditor

            ' Create a default Window for design
            Dim designWindow As VCF.IWindow = CreateDesignWindow()
            m_DesignSurface.DesignObject = designWindow
        Catch ex As Exception
            MessageBox.Show("Error initializing designer: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Private Function CreateDesignWindow() As VCF.IWindow
        ' Create a simple design window wrapper
        Dim win As New DesignWindow()
        Return win
    End Function

    Private Sub Form_Resize(sender As Object, e As EventArgs)
        ' Adjust design surface if needed
        If m_DesignSurface IsNot Nothing Then
            m_DesignSurface.Refresh()
        End If
    End Sub

    Private Sub Form_Unload(sender As Object, e As FormClosingEventArgs)
        ' Cleanup
        m_PropertyEditor = Nothing
        m_SelectionManager = Nothing
        m_Toolbox = Nothing
        m_DesignSurface = Nothing
    End Sub

    Public Sub SaveXAML()
        If m_DesignSurface Is Nothing Then Return
        If m_DesignSurface.DesignObject Is Nothing Then Return

        Try
            Dim writer As New XAMLWriter()

            Dim xaml As String = ""
            If TypeOf m_DesignSurface.DesignObject Is VCF.IWindow Then
                xaml = writer.Save(m_DesignSurface.DesignObject, "DesignWindow")
            ElseIf TypeOf m_DesignSurface.DesignObject Is VCF.IUserControl Then
                xaml = writer.Save(m_DesignSurface.DesignObject, "DesignUserControl")
            Else
                xaml = writer.Save(m_DesignSurface.DesignObject)
            End If

            ' TODO: Show save dialog and save to file
            System.Diagnostics.Debug.WriteLine(xaml)
        Catch ex As Exception
            MessageBox.Show("Error saving XAML: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    Public Sub LoadXAML(ByVal filePath As String)
        If m_DesignSurface Is Nothing Then Return

        Try
            Dim xml As String = System.IO.File.ReadAllText(filePath)

            Dim designObj As Object = Nothing
            Dim reader As VCF.XAMLReader = New VCF.XAMLReader()
            designObj = reader.Load(xml)

            If designObj IsNot Nothing Then
                m_DesignSurface.DesignObject = designObj
            End If
        Catch ex As Exception
            MessageBox.Show("Error loading XAML: " & ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
        End Try
    End Sub

    ' Mouse event handlers for drag-drop
    Private Sub picToolbox_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            Dim controlType As String = m_Toolbox.GetControlTypeAt(e.X, e.Y)

            If Not String.IsNullOrEmpty(controlType) Then
                m_Toolbox.StartDrag(controlType)
                m_IsDragging = True
                m_DragStartX = e.X
                m_DragStartY = e.Y
                picToolbox.Cursor = Cursors.Hand
            End If
        End If
    End Sub

    Private Sub picToolbox_MouseMove(sender As Object, e As MouseEventArgs)
        If m_IsDragging AndAlso e.Button = MouseButtons.Left Then
            ' Visual feedback during drag
            picToolbox.Cursor = Cursors.Hand
        End If
    End Sub

    Private Sub picToolbox_MouseUp(sender As Object, e As MouseEventArgs)
        If m_IsDragging Then
            m_IsDragging = False
            m_Toolbox.EndDrag()
            picToolbox.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub picDesignSurface_MouseDown(sender As Object, e As MouseEventArgs)
        If e.Button = MouseButtons.Left Then
            ' Check if we're dropping from toolbox
            If m_Toolbox.IsDragging Then
                ' Add control at drop location
                m_DesignSurface.AddControl(m_Toolbox.DragControlType, e.X, e.Y)
                m_Toolbox.EndDrag()
                m_IsDragging = False
            Else
                ' Try to select an element
                Dim hitElement As Object = m_SelectionManager.HitTest(e.X, e.Y)

                If hitElement IsNot Nothing Then
                    m_SelectionManager.SelectedElement = hitElement
                Else
                    m_SelectionManager.ClearSelection()
                End If
            End If
        End If
    End Sub

    Private Sub picDesignSurface_MouseMove(sender As Object, e As MouseEventArgs)
        If m_Toolbox.IsDragging Then
            picDesignSurface.Cursor = Cursors.Hand
        Else
            picDesignSurface.Cursor = Cursors.Default
        End If
    End Sub

    Private Sub picDesignSurface_MouseUp(sender As Object, e As MouseEventArgs)
        picDesignSurface.Cursor = Cursors.Default
    End Sub
End Class

