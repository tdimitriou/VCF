VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmDesigner 
   Caption         =   "VCF Designer"
   ClientHeight    =   9000
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   12000
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox picToolbox 
      Align           =   3  'Align Left
      Height          =   9000
      Left            =   0
      ScaleHeight     =   8940
      ScaleWidth      =   2400
      TabIndex        =   0
      Top             =   0
      Width           =   2400
   End
   Begin VB.PictureBox picProperties 
      Align           =   4  'Align Right
      Height          =   9000
      Left            =   9600
      ScaleHeight     =   8940
      ScaleWidth      =   2400
      TabIndex        =   1
      Top             =   0
      Width           =   2400
   End
   Begin VB.PictureBox picDesignSurface 
      Align           =   1  'Align Top
      Height          =   9000
      Left            =   2400
      ScaleHeight     =   8940
      ScaleWidth      =   7200
      TabIndex        =   2
      Top             =   0
      Width           =   7200
   End
End
Attribute VB_Name = "frmDesigner"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_DesignSurface As DesignSurface
Private m_Toolbox As DesignerToolbox
Private m_SelectionManager As DesignerSelectionManager
Private m_PropertyEditor As PropertyEditor
Private m_DragStartX As Single
Private m_DragStartY As Single
Private m_IsDragging As Boolean

Private Sub Form_Load()
    On Error Resume Next
    
    ' Initialize VCF Application for design-time
    Dim DesignApp As New DesignApplication
    
    ' Initialize designer components
    Set m_DesignSurface = New DesignSurface
    Set m_Toolbox = New DesignerToolbox
    Set m_SelectionManager = New DesignerSelectionManager
    
    ' Setup design surface
    m_DesignSurface.Initialize picDesignSurface
    m_DesignSurface.SelectionManager = m_SelectionManager
    
    ' Setup toolbox
    m_Toolbox.Initialize picToolbox
    m_Toolbox.DesignSurface = m_DesignSurface
    
    ' Setup selection manager
    m_SelectionManager.Initialize m_DesignSurface
    
    ' Setup property editor
    Set m_PropertyEditor = New PropertyEditor
    m_PropertyEditor.Initialize picProperties
    
    ' Wire selection changed event
    Set m_SelectionManager.PropertyEditor = m_PropertyEditor
    
    ' Create a default Window for design
    Dim DesignWindow As VCF.IWindow
    Set DesignWindow = CreateDesignWindow()
    m_DesignSurface.DesignObject = DesignWindow
    
    If Err Then
        MsgBox "Error initializing designer: " & Err.Description, vbExclamation
    End If
End Sub

Private Function CreateDesignWindow() As VCF.IWindow
    On Error Resume Next
    
    ' Create a simple design window wrapper
    Dim Win As New DesignWindow
    Set CreateDesignWindow = Win
    
    If Err Then Err.Clear
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    ' Adjust design surface if needed
    If Not m_DesignSurface Is Nothing Then
        m_DesignSurface.Refresh
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    
    ' Cleanup
    Set m_PropertyEditor = Nothing
    Set m_SelectionManager = Nothing
    Set m_Toolbox = Nothing
    Set m_DesignSurface = Nothing
    
    If Err Then Err.Clear
End Sub

Public Sub SaveXAML()
    On Error Resume Next
    
    If m_DesignSurface Is Nothing Then Exit Sub
    If m_DesignSurface.DesignObject Is Nothing Then Exit Sub
    
    Dim Writer As XAMLWriter
    Set Writer = New XAMLWriter
    
    Dim XAML As String
    If TypeOf m_DesignSurface.DesignObject Is VCF.IWindow Then
        XAML = Writer.Save(m_DesignSurface.DesignObject, "DesignWindow")
    ElseIf TypeOf m_DesignSurface.DesignObject Is VCF.IUserControl Then
        XAML = Writer.Save(m_DesignSurface.DesignObject, "DesignUserControl")
    Else
        XAML = Writer.Save(m_DesignSurface.DesignObject)
    End If
    
    ' TODO: Show save dialog and save to file
    Debug.Print XAML
    
    If Err Then
        MsgBox "Error saving XAML: " & Err.Description, vbExclamation
    End If
End Sub

Public Sub LoadXAML(ByVal FilePath As String)
    On Error Resume Next
    
    If m_DesignSurface Is Nothing Then Exit Sub
    
    Dim XML As String
    XML = New_c.FSO.ReadTextContent(FilePath)
    
    Dim DesignObj As Object
    Dim Reader As VCF.XAMLReader
    Set Reader = New VCF.XAMLReader
    Set DesignObj = Reader.Load(XML)
    
    If Not DesignObj Is Nothing Then
        m_DesignSurface.DesignObject = DesignObj
    End If
    
    If Err Then
        MsgBox "Error loading XAML: " & Err.Description, vbExclamation
    End If
End Sub

' Mouse event handlers for drag-drop
Private Sub picToolbox_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then ' Left button
        Dim ControlType As String
        ControlType = m_Toolbox.GetControlTypeAt(X, Y)
        
        If Len(ControlType) > 0 Then
            m_Toolbox.StartDrag ControlType
            m_IsDragging = True
            m_DragStartX = X
            m_DragStartY = Y
            picToolbox.MousePointer = vbCustom
        End If
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub picToolbox_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If m_IsDragging And Button = 1 Then
        ' Visual feedback during drag
        picToolbox.MousePointer = vbCustom
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub picToolbox_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If m_IsDragging Then
        m_IsDragging = False
        m_Toolbox.EndDrag
        picToolbox.MousePointer = vbDefault
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub picDesignSurface_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If Button = 1 Then ' Left button
        ' Check if we're dropping from toolbox
        If m_Toolbox.IsDragging Then
            ' Add control at drop location
            m_DesignSurface.AddControl m_Toolbox.DragControlType, X, Y
            m_Toolbox.EndDrag
            m_IsDragging = False
        Else
            ' Try to select an element
            Dim HitElement As Object
            Set HitElement = m_SelectionManager.HitTest(X, Y)
            
            If Not HitElement Is Nothing Then
                m_SelectionManager.SelectedElement = HitElement
            Else
                m_SelectionManager.ClearSelection
            End If
        End If
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub picDesignSurface_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    If m_Toolbox.IsDragging Then
        picDesignSurface.MousePointer = vbCustom
    Else
        picDesignSurface.MousePointer = vbDefault
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub picDesignSurface_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    
    picDesignSurface.MousePointer = vbDefault
    
    If Err Then Err.Clear
End Sub
