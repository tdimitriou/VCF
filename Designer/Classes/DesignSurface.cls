VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignSurface"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Container As PictureBox
Private m_DesignObject As Object
Private m_SelectionManager As DesignerSelectionManager
Private m_IsDesignMode As Boolean

Public Property Get Container() As PictureBox
    Set Container = m_Container
End Property

Public Property Get DesignObject() As Object
    Set DesignObject = m_DesignObject
End Property

Public Property Set DesignObject(ByVal Value As Object)
    Set m_DesignObject = Value
    
    ' If it's a Window, we need to host it in the container
    If TypeOf Value Is VCF.IWindow Then
        Dim Win As VCF.IWindow
        Set Win = Value
        ' TODO: Host the window's form in the container
    End If
    
    Refresh
End Property

Public Property Get SelectionManager() As DesignerSelectionManager
    Set SelectionManager = m_SelectionManager
End Property

Public Property Set SelectionManager(ByVal Value As DesignerSelectionManager)
    Set m_SelectionManager = Value
End Property

Public Property Get IsDesignMode() As Boolean
    IsDesignMode = m_IsDesignMode
End Property

Public Property Let IsDesignMode(ByVal Value As Boolean)
    m_IsDesignMode = Value
End Property

Public Sub Initialize(ByVal Container As PictureBox)
    On Error Resume Next
    
    Set m_Container = Container
    m_IsDesignMode = True
    
    ' Setup container
    Container.AutoRedraw = True
    Container.BackColor = vbWhite
    
    ' Draw grid
    DrawGrid
    
    If Err Then Err.Clear
End Sub

Public Sub Refresh()
    On Error Resume Next
    
    If m_Container Is Nothing Then Exit Sub
    
    ' Redraw the design surface
    m_Container.Cls
    DrawGrid
    
    ' Render the design object and its children
    RenderDesignObject
    
    ' Draw selection if any
    If Not m_SelectionManager Is Nothing Then
        If Not m_SelectionManager.SelectedElement Is Nothing Then
            m_SelectionManager.DrawSelection
        End If
    End If
    
    m_Container.Refresh
    
    If Err Then Err.Clear
End Sub

Private Sub RenderDesignObject()
    On Error Resume Next
    
    If m_DesignObject Is Nothing Then Exit Sub
    If m_Container Is Nothing Then Exit Sub
    
    Dim Children As VCF.UIElementCollection
    Dim Child As Object
    Dim UIElement As IUIElement
    
    ' Get children collection
    If TypeOf m_DesignObject Is VCF.IControl Then
        Dim Control As VCF.IControl
        Set Control = m_DesignObject
        Set Children = Control.Children
    ElseIf TypeOf m_DesignObject Is VCF.IWindow Then
        Dim Win As VCF.IWindow
        Set Win = m_DesignObject
        Set Children = Win.Base.Children
    Else
        Exit Sub
    End If
    
    If Children Is Nothing Then Exit Sub
    
    ' Render each child
    For Each Child In Children
        If TypeOf Child Is VCF.IUIElement Then
            Set UIElement = Child
            RenderElement UIElement
        End If
    Next
    
    If Err Then Err.Clear
End Sub

Private Sub RenderElement(ByVal UIElement As VCF.IUIElement)
    On Error Resume Next
    
    Dim Left As Single, Top As Single, Width As Single, Height As Single
    Left = UIElement.DesignLeft
    Top = UIElement.DesignTop
    Width = UIElement.DesignWidth
    Height = UIElement.DesignHeight
    
    ' Draw element rectangle
    m_Container.Line (Left, Top)-(Left + Width, Top + Height), &HC0C0C0, B
    
    ' Draw element type name
    Dim ObjType As String
    ObjType = TypeName(UIElement)
    m_Container.CurrentX = Left + 2
    m_Container.CurrentY = Top + 2
    m_Container.Print ObjType
    
    ' If it's a control with children, render them
    If TypeOf UIElement Is VCF.IControl Then
        Dim Control As VCF.IControl
        Set Control = UIElement
        If Control.Children.Count > 0 Then
            Dim Child As Object
            For Each Child In Control.Children
                If TypeOf Child Is VCF.IUIElement Then
                    RenderElement Child
                End If
            Next
        End If
    End If
    
    If Err Then Err.Clear
End Sub

Private Sub DrawGrid()
    On Error Resume Next
    
    If m_Container Is Nothing Then Exit Sub
    
    Dim i As Long
    Dim GridSize As Long
    GridSize = 10
    
    ' Draw vertical lines
    For i = 0 To m_Container.ScaleWidth Step GridSize
        m_Container.Line (i, 0)-(i, m_Container.ScaleHeight), &HE0E0E0
    Next
    
    ' Draw horizontal lines
    For i = 0 To m_Container.ScaleHeight Step GridSize
        m_Container.Line (0, i)-(m_Container.ScaleWidth, i), &HE0E0E0
    Next
    
    If Err Then Err.Clear
End Sub

Public Sub AddControl(ByVal ControlType As String, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    
    If m_DesignObject Is Nothing Then Exit Sub
    
    Dim NewControl As Object
    Set NewControl = CreateControl(ControlType)
    
    If NewControl Is Nothing Then Exit Sub
    
    ' Set position
    If TypeOf NewControl Is VCF.IUIElement Then
        Dim UIElement As VCF.IUIElement
        Set UIElement = NewControl
        UIElement.DesignLeft = X
        UIElement.DesignTop = Y
    End If
    
    ' Add to design object
    If TypeOf m_DesignObject Is VCF.IControl Then
        Dim Control As VCF.IControl
        Set Control = m_DesignObject
        Control.Children.Add NewControl
    ElseIf TypeOf m_DesignObject Is VCF.IWindow Then
        Dim Win As VCF.IWindow
        Set Win = m_DesignObject
        Win.Base.Children.Add NewControl
    End If
    
    Refresh
    
    If Err Then
        MsgBox "Error adding control: " & Err.Description, vbExclamation
    End If
End Sub

Private Function CreateControl(ByVal ControlType As String) As Object
    On Error Resume Next
    
    Dim Obj As Object
    
    Select Case LCase$(ControlType)
        Case "button"
            Set Obj = New VCF.Button
        Case "panel"
            Set Obj = New VCF.Panel
        Case "border"
            Set Obj = New VCF.Border
        Case "textblock"
            Set Obj = New VCF.TextBlock
        Case "textbox"
            Set Obj = New VCF.TextBox
        Case "image"
            Set Obj = New VCF.Image
        Case "uniformgrid"
            Set Obj = New VCF.UniformGrid
        Case Else
            ' Try to create via CreateObject
            On Error Resume Next
            Set Obj = CreateObject("VCF." & ControlType)
            If Err Then Err.Clear
    End Select
    
    Set CreateControl = Obj
    
    If Err Then Err.Clear
End Function

