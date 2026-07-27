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
Private m_Arranging As Boolean
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private m_ParentPtr As Long
Private m_Rows As Long
Private m_Columns As Long
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase
Private m_SyncingVisibility As Boolean
Private m_DesiredWidth As Double
Private m_DesiredHeight As Double
Private m_ActualWidth As Double
Private m_ActualHeight As Double

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

    Public Property Get Width() As Double
        If m_DependencyProperties.Exists("Width") Then
            Width = CDbl(m_DependencyProperties.GetValue("Width"))
        End If
    End Property

    Public Property Let Width(ByVal Value As Double)
        Call m_DependencyProperties.SetValue("Width", Value)
    End Property

    Public Property Get Height() As Double
        If m_DependencyProperties.Exists("Height") Then
            Height = CDbl(m_DependencyProperties.GetValue("Height"))
        End If
    End Property

    Public Property Let Height(ByVal Value As Double)
        Call m_DependencyProperties.SetValue("Height", Value)
    End Property

    Public Property Get Margin() As Thickness
        Set Margin = m_DependencyProperties.GetValue("Margin")
    End Property

    Public Property Set Margin(ByVal Value As Thickness)
        Call m_DependencyProperties.SetValue("Margin", Value)
    End Property

    Public Property Get Visible() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("Visible"), Visible)
    End Property

    Public Property Let Visible(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("Visible", Value)
    End Property

    Public Property Get Visibility() As Visibility
        If m_DependencyProperties.Exists("Visibility") Then
            Visibility = CLng(m_DependencyProperties.GetValue("Visibility"))
        Else
            Visibility = VisibilityVisible
        End If
    End Property

    Public Property Let Visibility(ByVal Value As Visibility)
        If m_DependencyProperties.Exists("Visibility") Then
            Call m_DependencyProperties.SetValue("Visibility", CLng(Value))
        End If
    End Property

    Public Property Get DesiredWidth() As Double
        DesiredWidth = m_DesiredWidth
    End Property

    Public Property Get DesiredHeight() As Double
        DesiredHeight = m_DesiredHeight
    End Property

    Public Property Get ActualWidth() As Double
        ActualWidth = m_ActualWidth
    End Property

    Public Property Get ActualHeight() As Double
        ActualHeight = m_ActualHeight
    End Property

    ' WPF-ish: max child cell size x Columns/Rows when Width/Height unset.
    Public Sub MeasureOverride(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
        Dim Child As Object
        Dim Measured As LayoutRect
        Dim MaxCW As Double
        Dim MaxCH As Double
        Dim ContentW As Double
        Dim ContentH As Double
        Dim DesiredW As Double
        Dim DesiredH As Double
        Dim ChildVis As Visibility
        Dim P As Panel

        MaxCW = 0#
        MaxCH = 0#

        If Not m_Children Is Nothing Then
            For Each Child In m_Children
                ChildVis = ReadElementVisibility(Child)
                If IsLayoutCollapsed(ChildVis) Then GoTo NextMeasureChild

                Measured = MeasureElementSize(Child, 0#, 0#)
                If Measured.Width <= 0! Or Measured.Height <= 0! Then
                    On Error Resume Next
                    Err.Clear
                    If TypeOf Child Is Panel Then
                        Set P = Child
                        If Measured.Width <= 0! Then Measured.Width = CSng(P.Width)
                        If Measured.Height <= 0! Then Measured.Height = CSng(P.Height)
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
                If CDbl(Measured.Width) > MaxCW Then MaxCW = CDbl(Measured.Width)
                If CDbl(Measured.Height) > MaxCH Then MaxCH = CDbl(Measured.Height)

NextMeasureChild:
            Next
        End If

        ContentW = MaxCW * CDbl(m_Columns)
        ContentH = MaxCH * CDbl(m_Rows)

        DesiredW = Width
        DesiredH = Height
        If DesiredW <= 0 Then DesiredW = ContentW
        If DesiredH <= 0 Then DesiredH = ContentH
        If AvailableWidth > 0 And DesiredW > AvailableWidth Then DesiredW = AvailableWidth
        If AvailableHeight > 0 And DesiredH > AvailableHeight Then DesiredH = AvailableHeight

        DesiredW = ClampElementWidth(Me, DesiredW)
        DesiredH = ClampElementHeight(Me, DesiredH)
        m_DesiredWidth = DesiredW
        m_DesiredHeight = DesiredH
    End Sub

    Public Sub MeasureLayout(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
        MeasureOverride AvailableWidth, AvailableHeight
    End Sub

    ' ItemsPanelTemplate clone ? set metrics without MoveChildren (grid not hosted yet).
    Friend Sub ApplyItemsPanelMetrics(ByVal NewRows As Long, ByVal NewColumns As Long, ByVal Pad As Thickness)
        Dim PadCopy As Thickness

        If NewRows < 1 Then NewRows = 1
        If NewColumns < 1 Then NewColumns = 1
        m_Rows = NewRows
        m_Columns = NewColumns
        If Not Pad Is Nothing Then
            Set PadCopy = New Thickness
            PadCopy.Left = Pad.Left
            PadCopy.Top = Pad.Top
            PadCopy.Right = Pad.Right
            PadCopy.Bottom = Pad.Bottom
            m_DependencyProperties.SetCurrentValue "Padding", PadCopy
        End If
    End Sub
'#End Region

'#Region IUIElement Inheritance
    Public Property Get AttachedProperties() As ObservableDictionary
        Set AttachedProperties = m_BaseUIElement.AttachedProperties
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
            Call .Register("Visibility", vbLong, , , , NewDependencyPropertyMetadata(False, True, False))
            .SetCurrentValue "Visible", True
            .SetCurrentValue "Visibility", CLng(VisibilityVisible)
            Call .Register("Padding", vbObject, "Thickness")
            Call .Register("Width", vbDouble, , , , NewDependencyPropertyMetadata(True, False, False))
            Call .Register("Height", vbDouble, , , , NewDependencyPropertyMetadata(True, False, False))
            Call .Register("Margin", vbObject, "Thickness", , NewDependencyPropertyMetadata(True, False, False))
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            .SetValue "Style", Nothing
            .SetCurrentValue "Width", 0#
            .SetCurrentValue "Height", 0#
            .SetCurrentValue "Margin", NewThickness(0)
        End With
        
            
        m_Columns = 3
        m_Rows = 3
        
        ' Cell inset default ? VCF UniformGrid (not WPF Panel). Keep 2; apps may set Padding="4".
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
        Dim Child As Object
        Dim ChildControl As IControl
        Dim Key As String

        On Error GoTo Catch

        If Not e.OldItems Is Nothing Then
            For Each Child In e.OldItems
                If TypeOf Child Is IControl Then
                    Set ChildControl = Child
                    Key = modLayoutEngine.ControlWidgetKey(Child)
                    If Not ChildControl.Widget Is Nothing Then
                        If W.Widgets.Exists(Key) Then W.Widgets.Remove Key
                    End If
                End If
            Next
        End If

        ' Honor LockRefresh ? ItemsControl batches AddSilent under lock then ArrangeChildren once.
        If W.LockRefresh Then Exit Sub
        MoveChildren
        If Not W.LockRefresh Then W.Refresh
        Exit Sub

Catch:
        '
    End Sub
'#End Region

'#Region Private Methods
    ' ItemsControl post-rebuild arrange (Friend ? same project).
    Friend Sub ArrangeChildren()
        MoveChildren
    End Sub

    Public Sub ArrangeOverride(ByVal FinalWidth As Double, ByVal FinalHeight As Double)
        MeasureOverride FinalWidth, FinalHeight
        MoveChildren
    End Sub

    Public Sub RelayoutChildren()
        MoveChildren
    End Sub

    Private Sub MoveChildren()
        Dim MAX_ITEMS As Long
        Dim Index As Long
        Dim ItemIndex As Long
        Dim Child As Object
        Dim CellPosition As Coordinates
        Dim Cells

        If W Is Nothing Then Exit Sub
        If W.LockRefresh Then Exit Sub
        If m_Arranging Then Exit Sub
        m_Arranging = True

        On Error GoTo Clean

        MeasureOverride CDbl(W.Width), CDbl(W.Height)

        MAX_ITEMS = m_Rows * m_Columns
        If MAX_ITEMS < 1 Then GoTo Clean
        ReDim Cells(0 To MAX_ITEMS - 1) As Long

        ' Single RemoveAll per arrange (not per child add) ? avoids O(N^2) IDE hang.
        W.Widgets.RemoveAll
        ItemIndex = 0
        For Index = 0 To MAX_ITEMS - 1
            If Cells(Index) = 0 Then
                Do While ItemIndex <= m_Children.Count - 1
                    Set Child = m_Children(ItemIndex)
                    If IsLayoutCollapsed(ReadElementVisibility(Child)) Then
                        DetachCollapsedChild Child, W
                        ItemIndex = ItemIndex + 1
                    Else
                        Exit Do
                    End If
                Loop
                If ItemIndex > m_Children.Count - 1 Then Exit For
                Set Child = m_Children(ItemIndex)

                CellPosition = GetGridPosition(Index)

                Call MoveChild(m_Children(ItemIndex), GetCoordinates(CellPosition.T, CellPosition.L), Index, CellPosition, Cells)
                ItemIndex = ItemIndex + 1
            End If
        Next

Clean:
        m_ActualWidth = CDbl(W.Width)
        m_ActualHeight = CDbl(W.Height)
        m_Arranging = False
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
        
        Dim Obj As IUIElement
        Dim Key As String
        Dim ChildVis As Visibility
        Set Obj = Child
        
        Dim ColSpan As Long
        Dim RowSpan As Long
        ColSpan = GetGridAttachedLong(Obj, "ColumnSpan", 1)
        RowSpan = GetGridAttachedLong(Obj, "RowSpan", 1)
        If ColSpan < 1 Then ColSpan = 1
        If RowSpan < 1 Then RowSpan = 1

        Dim StartRow As Long
        Dim StartCol As Long
        StartRow = CellPosition.T
        StartCol = CellPosition.L

        If RowSpan > (m_Rows - StartRow) Then RowSpan = m_Rows - StartRow
        If ColSpan > (m_Columns - StartCol) Then ColSpan = m_Columns - StartCol
        
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
            
        End With

        If WW < 0! Then WW = 0!
        If HH < 0! Then HH = 0!
        WW = CSng(ClampElementWidth(Child, CDbl(WW)))
        HH = CSng(ClampElementHeight(Child, CDbl(HH)))
    
        Key = modLayoutEngine.ControlWidgetKey(Child)
        ChildVis = ReadElementVisibility(Child)
        If Not W.Widgets.Exists(Key) Then
            W.Widgets.Add Child, Key, , , , , IsWidgetVisible(ChildVis)
        End If
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
            Case "Visible"
                If Not m_SyncingVisibility Then
                    m_SyncingVisibility = True
                    If CBool(Property.GetValue) Then
                        m_DependencyProperties.SetValue "Visibility", CLng(VisibilityVisible)
                    Else
                        m_DependencyProperties.SetValue "Visibility", CLng(VisibilityCollapsed)
                    End If
                    m_SyncingVisibility = False
                End If
            Case "Visibility"
                If Not m_SyncingVisibility Then
                    m_SyncingVisibility = True
                    m_DependencyProperties.SetCurrentValue "Visible", (CLng(Property.GetValue) = VisibilityVisible)
                    m_SyncingVisibility = False
                End If
                ApplyVisibility W, CLng(Property.GetValue)
                InvalidateParentLayout Me
            Case "Style":           ApplyStyle
            Case "Width", "Height", "Margin", "Padding"
                If Not W.LockRefresh Then MoveChildren
        End Select
        
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
        If Not Property.Metadata Is Nothing Then
            If Property.Metadata.AffectsRender Then W.Refresh
        End If
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
        If W.LockRefresh Then Exit Sub
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

    Friend Sub PrepareForUnload()
        On Error Resume Next
        Set ActiveStyle = Nothing
        Set W = Nothing
    End Sub

