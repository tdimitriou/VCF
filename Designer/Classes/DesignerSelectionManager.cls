VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignerSelectionManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_DesignSurface As DesignSurface
Private m_SelectedElement As Object
Private m_SelectionHandles As Collection
Private m_PropertyEditor As PropertyEditor

Public Property Get SelectedElement() As Object
    Set SelectedElement = m_SelectedElement
End Property

Public Property Set SelectedElement(ByVal Value As Object)
    Set m_SelectedElement = Value
    DrawSelection
    
    ' Update property editor
    If Not m_PropertyEditor Is Nothing Then
        Set m_PropertyEditor.SelectedObject = Value
    End If
End Property

Public Property Get PropertyEditor() As PropertyEditor
    Set PropertyEditor = m_PropertyEditor
End Property

Public Property Set PropertyEditor(ByVal Value As PropertyEditor)
    Set m_PropertyEditor = Value
End Property

Public Property Get DesignSurface() As DesignSurface
    Set DesignSurface = m_DesignSurface
End Property

Public Sub Initialize(ByVal DesignSurface As DesignSurface)
    On Error Resume Next
    
    Set m_DesignSurface = DesignSurface
    Set m_SelectionHandles = New Collection
    
    If Err Then Err.Clear
End Sub

Public Sub DrawSelection()
    On Error Resume Next
    
    If m_DesignSurface Is Nothing Then Exit Sub
    If m_DesignSurface.Container Is Nothing Then Exit Sub
    If m_SelectedElement Is Nothing Then Exit Sub
    
    Dim Container As PictureBox
    Set Container = m_DesignSurface.Container
    
    ' Get element bounds
    Dim Left As Single, Top As Single, Width As Single, Height As Single
    
    If TypeOf m_SelectedElement Is VCF.IUIElement Then
        Dim UIElement As VCF.IUIElement
        Set UIElement = m_SelectedElement
        Left = UIElement.DesignLeft
        Top = UIElement.DesignTop
        Width = UIElement.DesignWidth
        Height = UIElement.DesignHeight
    Else
        Exit Sub
    End If
    
    ' Draw selection rectangle
    Container.Line (Left, Top)-(Left + Width, Top + Height), vbBlue, B
    
    ' Draw resize handles (8 handles around the rectangle)
    DrawHandle Container, Left, Top
    DrawHandle Container, Left + Width / 2, Top
    DrawHandle Container, Left + Width, Top
    DrawHandle Container, Left + Width, Top + Height / 2
    DrawHandle Container, Left + Width, Top + Height
    DrawHandle Container, Left + Width / 2, Top + Height
    DrawHandle Container, Left, Top + Height
    DrawHandle Container, Left, Top + Height / 2
    
    Container.Refresh
    
    If Err Then Err.Clear
End Sub

Private Sub DrawHandle(ByVal Container As PictureBox, ByVal X As Single, ByVal Y As Single)
    On Error Resume Next
    
    Dim HandleSize As Long
    HandleSize = 6
    
    Container.Line (X - HandleSize / 2, Y - HandleSize / 2)-(X + HandleSize / 2, Y + HandleSize / 2), vbBlue, BF
    
    If Err Then Err.Clear
End Sub

Public Sub ClearSelection()
    On Error Resume Next
    
    Set m_SelectedElement = Nothing
    
    ' Clear property editor
    If Not m_PropertyEditor Is Nothing Then
        Set m_PropertyEditor.SelectedObject = Nothing
    End If
    
    If Not m_DesignSurface Is Nothing Then
        If Not m_DesignSurface.Container Is Nothing Then
            m_DesignSurface.Refresh
        End If
    End If
    
    If Err Then Err.Clear
End Sub

Public Function HitTest(ByVal X As Single, ByVal Y As Single) As Object
    On Error Resume Next
    
    ' Search through design object's children to find element at X, Y
    If m_DesignSurface Is Nothing Then Exit Function
    If m_DesignSurface.DesignObject Is Nothing Then Exit Function
    
    Dim Children As UIElementCollection
    Dim Child As Object
    Dim UIElement As IUIElement
    
    ' Get children collection
    If TypeOf m_DesignSurface.DesignObject Is VCF.IControl Then
        Dim Control As VCF.IControl
        Set Control = m_DesignSurface.DesignObject
        Set Children = Control.Children
    ElseIf TypeOf m_DesignSurface.DesignObject Is VCF.IWindow Then
        Dim Win As VCF.IWindow
        Set Win = m_DesignSurface.DesignObject
        Set Children = Win.Base.Children
    Else
        Exit Function
    End If
    
    If Children Is Nothing Then Exit Function
    
    ' Search children (reverse order to get top-most first)
    Dim i As Long
    For i = Children.Count - 1 To 0 Step -1
        Set Child = Children(i)
        
        If TypeOf Child Is VCF.IUIElement Then
            Set UIElement = Child
            
            ' Check if point is within bounds
            If X >= UIElement.DesignLeft And X <= UIElement.DesignLeft + UIElement.DesignWidth And _
               Y >= UIElement.DesignTop And Y <= UIElement.DesignTop + UIElement.DesignHeight Then
                Set HitTest = Child
                Exit Function
            End If
        End If
    Next
    
    Set HitTest = Nothing
    
    If Err Then Err.Clear
End Function

