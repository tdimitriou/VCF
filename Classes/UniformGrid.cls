VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UniformGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

Private Type Coordinates
    L As Single
    T As Single
    W As Single
    H As Single
End Type

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents m_Children As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_DesignWidth As Single
Private m_DesignHeight As Single
Private m_ParentPtr As Long
Private m_Rows As Long
Private m_Columns As Long
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase

Public Name As String

'#Region Public Properties
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Get Style() As Style
        Call API.CopyVariable(m_DependencyProperties.GetValue("Style"), Style)
    End Property

    Public Property Set Style(ByVal Value As Style)
        Call m_DependencyProperties.SetValue("Style", Value)
    End Property
    
    Public Property Get Padding() As Thickness
        Call API.CopyVariable(m_DependencyProperties.GetValue("Padding"), Padding)
    End Property

    Public Property Set Padding(ByVal Value As Thickness)
        Call m_DependencyProperties.SetValue("Padding", Value)
        
        Call MoveChildren
    End Property
    
    Public Property Get Rows() As Long
        Rows = m_Rows
    End Property

    Public Property Let Rows(ByVal Value As Long)
        If Value < 1 Then Value = 1
        If m_Rows = Value Then Exit Property
        
        m_Rows = Value
        
        Call MoveChildren
    End Property

    Public Property Get Columns() As Long
        Columns = m_Columns
    End Property

    Public Property Let Columns(ByVal Value As Long)
        If Value < 1 Then Value = 1
        If m_Columns = Value Then Exit Property
        
        m_Columns = Value
        
        Call MoveChildren
    End Property
'#End Region

'#Region IUIElement Inheritance
    Public Property Get AttachedProperties() As ObservableDictionary
        Set AttachedProperties = m_BaseUIElement.AttachedProperties
    End Property
    
    Public Property Let DesignHeight(ByVal Value As Double)
        m_DesignHeight = Value
    End Property
    
    Public Property Get DesignHeight() As Double
        DesignHeight = m_DesignHeight
    End Property
    
    Public Property Let DesignLeft(ByVal Value As Double)
        m_DesignLeft = Value
    End Property
    
    Public Property Get DesignLeft() As Double
        DesignLeft = m_DesignLeft
    End Property
    
    Public Property Let DesignTop(ByVal Value As Double)
        m_DesignTop = Value
    End Property
    
    Public Property Get DesignTop() As Double
        DesignTop = m_DesignTop
    End Property
    
    Public Property Let DesignWidth(ByVal Value As Double)
        m_DesignWidth = Value
    End Property
    
    Public Property Get DesignWidth() As Double
        DesignWidth = m_DesignWidth
    End Property

    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        If Parent Is Nothing Then Exit Sub
        
        'If Parent.Widgets.Exists(Widget.Key) Then Parent.Widgets.Remove Widget.Key
        'Parent.Widgets.Add Me, "_" & ObjPtr(Me), Left, Top, Width, Height, DependencyProperties.GetValue("Visible")
            
        W.Move Left, Top, Width, Height
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        m_ParentPtr = ObjPtr(Value)
        DependencyPropertiesStatic.InheritPropertyValues Me
    End Property
    
    Public Property Get Parent() As IControl
        Call API.CopyVariable(API.ObjFromPtr(m_ParentPtr), Parent)
    End Property
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Bindings = New VCF.List
        
        Set W = Cairo.WidgetBase
        W.LockRefresh = True
        W.CanGetFocus = False
        
        Set m_Children = NewUIElementCollection(Me)
        Set m_DependencyProperties = NewDependencyProperties(Me)
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ShowGridLines", vbBoolean)
            Call .Register("Visible", vbBoolean)
            .SetCurrentValue "Visible", True
            Call .Register("Padding", vbObject, "Thickness")
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            .SetValue "Style", Nothing
        End With
        
        DesignLeft = 0
        DesignTop = 0
        DesignWidth = 300
        DesignHeight = 300
            
        m_Columns = 3
        m_Rows = 3
        
        Set Me.Padding = NewThickness(2)
        
        NewUIElementBase Me, m_BaseUIElement

        'EnumerateProperties Me

        Set Style = GetBaseStyle
        W.LockRefresh = False
    End Sub
'#End Region

'#Region IControl Interface Implementation
    Private Property Get IControl_Children() As UIElementCollection
        Set IControl_Children = Children
    End Property
    
    Private Property Get IControl_Widget() As vbRichClient5.cWidgetBase
        Set IControl_Widget = Widget
    End Property
    
    Private Property Get IControl_Widgets() As vbRichClient5.cWidgets
        Set IControl_Widgets = Widgets
    End Property
'#End Region

'#Region IControl Inheritance
    Public Property Get Children() As UIElementCollection
        Set Children = m_Children
    End Property
    
    Public Property Get Widget() As cWidgetBase
        Set Widget = W
    End Property
    
    Public Property Get Widgets() As cWidgets
        Set Widgets = W.Widgets
    End Property
'#End Region

'#Region IDependencyObject Inheritance
    Public Property Get DependencyProperties() As DependencyProperties
        Set DependencyProperties = m_DependencyProperties
    End Property
'#End Region

'#Region IDependencyObject Interface Implementation
    Private Property Get IDependencyObject_Children() As Object
        Set IDependencyObject_Children = Children
    End Property
    
    Private Property Get IDependencyObject_DependencyProperties() As DependencyProperties
        Set IDependencyObject_DependencyProperties = DependencyProperties
    End Property

    Private Property Get IDependencyObject_Parent() As IDependencyObject
        Set IDependencyObject_Parent = Parent
    End Property
'#End Region

'#Region IUIElement Interface Implementation
    Private Property Get IUIElement_AttachedProperties() As ObservableDictionary
        Set IUIElement_AttachedProperties = m_BaseUIElement.AttachedProperties
    End Property
    
    Public Property Set DataContext(ByVal Value As Object)
        Call m_DependencyProperties.SetValue("DataContext", Value)
        
        ' TO-DO: Recreate the Bindings!!!
    End Property
    
    Public Property Get DataContext() As Object
        Call API.CopyVariable(m_DependencyProperties.GetValue("DataContext"), DataContext)
    End Property
    
    Private Property Get IUIElement_Base() As UIElementBase
        Set IUIElement_Base = m_BaseUIElement
    End Property

    Private Property Set IUIElement_DataContext(ByVal RHS As Object)
        '
    End Property
    
    Private Property Get IUIElement_DataContext() As Object
        '
    End Property
    
    Private Property Let IUIElement_DesignHeight(ByVal RHS As Double)
        DesignHeight = RHS
    End Property
    
    Private Property Get IUIElement_DesignHeight() As Double
        IUIElement_DesignHeight = DesignHeight
    End Property
    
    Private Property Let IUIElement_DesignLeft(ByVal RHS As Double)
        DesignLeft = RHS
    End Property
    
    Private Property Get IUIElement_DesignLeft() As Double
        IUIElement_DesignLeft = DesignLeft
    End Property
    
    Private Property Let IUIElement_DesignTop(ByVal RHS As Double)
        DesignTop = RHS
    End Property
    
    Private Property Get IUIElement_DesignTop() As Double
        IUIElement_DesignTop = DesignTop
    End Property
    
    Private Property Let IUIElement_DesignWidth(ByVal RHS As Double)
        DesignWidth = RHS
    End Property
    
    Private Property Get IUIElement_DesignWidth() As Double
        IUIElement_DesignWidth = DesignWidth
    End Property
    
    Private Sub IUIElement_Move(Left As Single, Optional Top As Single = 0#, Optional Width As Single, Optional Height As Single)
        Move Left, Top, Width, Height
    End Sub

    Private Property Set IUIElement_Parent(RHS As IControl)
        Set Parent = RHS
    End Property
    
    Private Property Get IUIElement_Parent() As IControl
        Set IUIElement_Parent = Parent
    End Property
'#End Region

'#Region UIElementCollection Events
    Private Sub m_Children_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
        On Error GoTo Catch
            
        Dim Child As Object
        
Try:
                       
        If Not e.OldItems Is Nothing Then
            For Each Child In e.OldItems
                If TypeOf Child Is IControl Then If W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Remove Child.Widget.Key
            Next
        End If
        
        MoveChildren
        If Not W.LockRefresh Then W.Refresh
        
        Exit Sub
        
Catch:
        '
    End Sub
'#End Region

'#Region Private Methods
    Private Sub MoveChildren()
        Dim MAX_ITEMS As Long
        Dim Index As Long
        Dim ItemIndex As Long
        Dim RowID As Long, ColID As Long
        
        MAX_ITEMS = m_Rows * m_Columns
                            
        Dim Cells
        ReDim Cells(0 To MAX_ITEMS - 1) As Long
        
        Dim Child As Object
        Dim CellPosition As Coordinates
        
        W.Widgets.RemoveAll
        For Index = 0 To MAX_ITEMS - 1
            If Cells(Index) = 0 Then
                If ItemIndex > m_Children.Count - 1 Then Exit For
                Set Child = m_Children(ItemIndex)
            
                CellPosition = GetGridPosition(Index)
            
                Call MoveChild(m_Children(ItemIndex), GetCoordinates(CellPosition.T, CellPosition.L), Index, CellPosition, Cells)
                ItemIndex = ItemIndex + 1
            End If
        Next
End Sub
    
    Private Function GetGridPosition(ByVal Index As Long) As Coordinates
        
        With GetGridPosition
            .L = Index Mod m_Columns
            .T = (Index + m_Columns - .L) / m_Columns - 1
        End With
    End Function
    
    Private Function GetCoordinates(ByVal RowID As Long, ByVal ColID As Long) As Coordinates
        With GetCoordinates
            .W = W.Width / m_Columns
            .H = W.Height / m_Rows
        
            .L = .W * ColID
            .T = .H * RowID
        End With
    End Function
    
    Private Sub MoveChild(Child As Object, Position As Coordinates, StartIndex As Long, CellPosition As Coordinates, Cells)
        On Error Resume Next
        
        'If W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Remove Child.Widget.Key
        'If Not W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Add Child, "_" & ObjPtr(Child)
        
        Dim Obj As IUIElement
        Set Obj = Child
        
        Dim ColSpan As Integer
        Dim RowSpan As Long
        ColSpan = 1
        RowSpan = 1
        
        Dim StartRow As Long
        Dim StartCol As Long
        StartRow = CellPosition.T
        StartCol = CellPosition.L
        
        Dim GridProps As ObservableDictionary
        If Obj.AttachedProperties.ContainsKey("Grid") Then
            Set GridProps = Obj.AttachedProperties("Grid")
            
            If GridProps.ContainsKey("ColumnSpan") Then ColSpan = GridProps("ColumnSpan")
            If GridProps.ContainsKey("RowSpan") Then RowSpan = GridProps("RowSpan")
        
            If RowSpan > (m_Rows - StartRow) Then RowSpan = m_Rows - StartRow
            If ColSpan > (m_Columns - StartCol) Then ColSpan = m_Columns - StartCol
        End If
        
       Dim i As Long
       Dim j As Long
       For i = StartIndex To StartIndex + ColSpan - 1
            For j = 0 To RowSpan - 1
                Cells((j * m_Columns) + i) = 1
            Next j
       Next
       
        Dim LL As Single, TT As Single, WW As Single, HH As Single
                    
        With Position
            
            LL = .L + Padding.Left
            TT = .T + Padding.Top
            WW = .W * ColSpan - (Padding.Left + Padding.Right)
            HH = .H * RowSpan - (Padding.Top + Padding.Bottom)
            
            'If Not W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Add Child, "_" & ObjPtr(Child), CDbl(LL), CDbl(TT), CDbl(WW), CDbl(HH)
        End With
    
        If W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Remove (Child.Widget.Key)
        W.Widgets.Add Child, "_" & ObjPtr(Child), , , , , Child.Widget.Visible
        Child.Move LL, TT, WW, HH
    End Sub
    
    Private Sub DrawGridLines(CC As cCairoContext)
        Dim RowID As Long, ColID As Long
        
        For RowID = 0 To m_Rows - 1
            For ColID = 0 To m_Columns - 1
                With GetCoordinates(RowID, ColID)
                    CC.Rectangle .L + 1, .T + 1, .W - 2, .H - 2, True
                    CC.SetLineWidth 0.25
                    CC.SetDashes 0, 0, 1, 1
                    CC.SetSourceColor vbWhite, 1
                    CC.Stroke
                End With
            Next
        Next
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        Select Case Property.Name
            Case "Visible":         W.Visible = Property.GetValue
            Case "Style":           ApplyStyle
        End Select
        
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
        If Property.Metadata.AffectsRender Then W.Refresh
    End Sub
'#End Region

'#Region Widget Events
    Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
        Dim dx As Double, dy As Double
        
        dx = dx_Aligned
        dy = dy_Aligned
        
        
        CC.Save

'        CC.Rectangle 0, 0, dx, dy
'        CC.SetSourceColor W.BackColor, W.Alpha
'        CC.Fill
        
        If DependencyProperties.GetValue("ShowGridLines") = True Then Call DrawGridLines(CC)
               
        CC.Restore
    End Sub
    
    Private Sub W_Resize()
'        If Not W.Visible Or W.IsHidden Then Exit Sub
        
        Call MoveChildren
    
    End Sub
'#End Region

Private Function GetBaseStyle() As Style
    Dim Style
    API.CopyVariable m_BaseUIElement.TryFindResource(Typename(Me) & ".BaseStyle"), Style
    
    If Not IsObject(Style) Then Exit Function
    If TypeOf Style Is Style Then Set GetBaseStyle = Style
End Function

Private Sub ApplyStyle()
    If Not ActiveStyle Is Style Then Set ActiveStyle = Style
    If Style Is Nothing Then Exit Sub
    With New StyleManager
        .ApplyStyle Style, Me
    End With
End Sub


    Private Sub ActiveStyle_StyleChanged()
        ApplyStyle
    End Sub

