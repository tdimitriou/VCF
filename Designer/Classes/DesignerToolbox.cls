VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignerToolbox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Container As PictureBox
Private m_DesignSurface As DesignSurface
Private m_ControlTypes As Collection
Private m_DragControlType As String
Private m_IsDragging As Boolean

Public Property Get DesignSurface() As DesignSurface
    Set DesignSurface = m_DesignSurface
End Property

Public Property Set DesignSurface(ByVal Value As DesignSurface)
    Set m_DesignSurface = Value
End Property

Public Sub Initialize(ByVal Container As PictureBox)
    On Error Resume Next
    
    Set m_Container = Container
    Set m_ControlTypes = New Collection
    
    ' Setup container
    Container.AutoRedraw = True
    Container.BackColor = &HF0F0F0
    
    ' Initialize control types
    InitializeControlTypes
    
    ' Draw toolbox
    DrawToolbox
    
    ' Setup mouse events
    SetupMouseEvents
    
    If Err Then Err.Clear
End Sub

Private Sub InitializeControlTypes()
    On Error Resume Next
    
    ' Add available control types
    m_ControlTypes.Add "Button", "Button"
    m_ControlTypes.Add "Panel", "Panel"
    m_ControlTypes.Add "Border", "Border"
    m_ControlTypes.Add "TextBlock", "TextBlock"
    m_ControlTypes.Add "TextBox", "TextBox"
    m_ControlTypes.Add "Image", "Image"
    m_ControlTypes.Add "UniformGrid", "UniformGrid"
    m_ControlTypes.Add "ListView", "ListView"
    
    If Err Then Err.Clear
End Sub

Private Sub DrawToolbox()
    On Error Resume Next
    
    If m_Container Is Nothing Then Exit Sub
    If m_ControlTypes Is Nothing Then Exit Sub
    
    m_Container.Cls
    
    Dim i As Long
    Dim Y As Long
    Dim ItemHeight As Long
    ItemHeight = 30
    Y = 10
    
    ' Draw title
    m_Container.CurrentX = 10
    m_Container.CurrentY = Y
    m_Container.Print "Toolbox"
    Y = Y + 20
    
    ' Draw control type buttons
    Dim ControlType As Variant
    For Each ControlType In m_ControlTypes
        ' Draw button background
        m_Container.Line (5, Y)-(m_Container.ScaleWidth - 5, Y + ItemHeight), vbWhite, BF
        m_Container.Line (5, Y)-(m_Container.ScaleWidth - 5, Y + ItemHeight), &HC0C0C0, B
        
        ' Draw text
        m_Container.CurrentX = 10
        m_Container.CurrentY = Y + 8
        m_Container.Print ControlType
        
        Y = Y + ItemHeight + 5
    Next
    
    m_Container.Refresh
    
    If Err Then Err.Clear
End Sub

Private Sub SetupMouseEvents()
    On Error Resume Next
    
    ' Note: In VB6, we'd need to subclass or use a different approach
    ' For now, this is a placeholder
    
    If Err Then Err.Clear
End Sub

Public Function GetControlTypeAt(ByVal X As Single, ByVal Y As Single) As String
    On Error Resume Next
    
    ' Calculate which control type is at the given coordinates
    Dim ItemHeight As Long
    ItemHeight = 30
    Dim StartY As Long
    StartY = 30 ' After title
    
    Dim Index As Long
    Index = Int((Y - StartY) / (ItemHeight + 5))
    
    If Index >= 1 And Index <= m_ControlTypes.Count Then
        GetControlTypeAt = m_ControlTypes(Index)
    End If
    
    If Err Then Err.Clear
End Function

Public Sub StartDrag(ByVal ControlType As String)
    On Error Resume Next
    
    m_DragControlType = ControlType
    m_IsDragging = True
    
    If Err Then Err.Clear
End Sub

Public Sub EndDrag()
    On Error Resume Next
    
    m_IsDragging = False
    m_DragControlType = ""
    
    If Err Then Err.Clear
End Sub

Public Property Get IsDragging() As Boolean
    IsDragging = m_IsDragging
End Property

Public Property Get DragControlType() As String
    DragControlType = m_DragControlType
End Property



