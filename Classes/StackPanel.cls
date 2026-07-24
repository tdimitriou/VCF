VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StackPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents m_Children As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private m_ParentPtr As Long
Private m_TemplatedParentPtr As Long
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase
Private m_Layout As FrameworkElement
Private m_Orientation As Orientation

Public Name As String

Public Property Get Orientation() As Orientation
    Orientation = m_Orientation
End Property

Public Property Let Orientation(ByVal Value As Orientation)
    If m_Orientation = Value Then Exit Property
    m_Orientation = Value
    If m_DependencyProperties.Exists("Orientation") Then
        If Value = OrientationHorizontal Then
            m_DependencyProperties.SetCurrentValue "Orientation", "Horizontal"
        Else
            m_DependencyProperties.SetCurrentValue "Orientation", "Vertical"
        End If
    End If
    If Not m_Children Is Nothing Then ArrangePanel
End Property

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

'#Region IUIElement Inheritance
    Public Property Get AttachedProperties() As ObservableDictionary
        Set AttachedProperties = m_BaseUIElement.AttachedProperties
    End Property
    
    Public Property Set DataContext(ByVal Value As Object)
        Call m_DependencyProperties.SetValue("DataContext", Value)
        
        ' TO-DO: Recreate the Bindings!!!
    End Property
    
    Public Property Get DataContext() As Object
        Call API.CopyVariable(m_DependencyProperties.GetValue("DataContext"), DataContext)
    End Property
    
    
    
    
    
    
    
    
    Public Property Get Width() As Double
        Width = m_Layout.Width
    End Property

    Public Property Let Width(ByVal Value As Double)
        m_Layout.Width = Value
    End Property

    Public Property Get Height() As Double
        Height = m_Layout.Height
    End Property

    Public Property Let Height(ByVal Value As Double)
        m_Layout.Height = Value
    End Property

    Public Property Get Margin() As Thickness
        Set Margin = m_Layout.Margin
    End Property

    Public Property Set Margin(ByVal Value As Thickness)
        Set m_Layout.Margin = Value
    End Property

    Public Property Let Visibility(ByVal Value As Visibility)
        m_Layout.Visibility = Value
    End Property
    
    Public Property Get Visibility() As Visibility
        Visibility = m_Layout.Visibility
    End Property

    Public Property Get IsHitTestVisible() As Boolean
        IsHitTestVisible = m_Layout.IsHitTestVisible
    End Property

    Public Property Let IsHitTestVisible(ByVal Value As Boolean)
        m_Layout.IsHitTestVisible = Value
    End Property

    Public Property Get MinWidth() As Double
        MinWidth = m_Layout.MinWidth
    End Property

    Public Property Let MinWidth(ByVal Value As Double)
        m_Layout.MinWidth = Value
    End Property

    Public Property Get MinHeight() As Double
        MinHeight = m_Layout.MinHeight
    End Property

    Public Property Let MinHeight(ByVal Value As Double)
        m_Layout.MinHeight = Value
    End Property

    Public Property Get MaxWidth() As Double
        MaxWidth = m_Layout.MaxWidth
    End Property

    Public Property Let MaxWidth(ByVal Value As Double)
        m_Layout.MaxWidth = Value
    End Property

    Public Property Get MaxHeight() As Double
        MaxHeight = m_Layout.MaxHeight
    End Property

    Public Property Let MaxHeight(ByVal Value As Double)
        m_Layout.MaxHeight = Value
    End Property

    Public Property Get DesiredWidth() As Double
        DesiredWidth = m_Layout.DesiredWidth
    End Property

    Public Property Get DesiredHeight() As Double
        DesiredHeight = m_Layout.DesiredHeight
    End Property

    Public Property Get ActualWidth() As Double
        ActualWidth = m_Layout.ActualWidth
    End Property

    Public Property Get ActualHeight() As Double
        ActualHeight = m_Layout.ActualHeight
    End Property

    ' WPF-named measure entry (duck-typed via CallByName from MeasureElementSize).
    Public Sub MeasureOverride(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
        Dim ContentW As Double
        Dim ContentH As Double
        Dim DesiredW As Double
        Dim DesiredH As Double

        If m_Layout Is Nothing Then Exit Sub

        MeasureStackPanelContent m_Children, m_Orientation, AvailableWidth, AvailableHeight, ContentW, ContentH

        DesiredW = Width
        DesiredH = Height
        If DesiredW <= 0 Then DesiredW = ContentW
        If DesiredH <= 0 Then DesiredH = ContentH
        If AvailableWidth > 0 And DesiredW > AvailableWidth Then DesiredW = AvailableWidth
        If AvailableHeight > 0 And DesiredH > AvailableHeight Then DesiredH = AvailableHeight

        DesiredW = ClampElementWidth(Me, DesiredW)
        DesiredH = ClampElementHeight(Me, DesiredH)
        m_Layout.ReportDesiredSize DesiredW, DesiredH
    End Sub

    ' Compat alias for MeasureOverride.
    Public Sub MeasureLayout(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
        MeasureOverride AvailableWidth, AvailableHeight
    End Sub

    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        If Parent Is Nothing Then Exit Sub
        
        'If Parent.Widgets.Exists(Widget.Key) Then Parent.Widgets.Remove Widget.Key
        'Parent.Widgets.Add Me, "_" & ObjPtr(Me), Left, Top, Width, Height
        
        W.Move Left, Top, Width, Height
        
        ApplyVisibility W, Visibility
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        m_ParentPtr = ObjPtr(Value)
        DependencyPropertiesStatic.InheritPropertyValues Me
    End Property
    
    Public Property Get Parent() As IControl
        Call API.CopyVariable(API.ObjFromPtr(m_ParentPtr), Parent)
    End Property

    Public Property Set TemplatedParent(ByVal Value As Object)
        m_TemplatedParentPtr = ObjPtr(Value)
    End Property

    Public Property Get TemplatedParent() As Object
        Call API.CopyVariable(API.ObjFromPtr(m_TemplatedParentPtr), TemplatedParent)
    End Property
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Bindings = New List
        Set m_DependencyProperties = NewDependencyProperties(Me)
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ShowGridLines", vbBoolean)
            Call .Register("Orientation", vbString)
            .SetCurrentValue "Orientation", "Vertical"
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            .SetValue "Style", Nothing
        End With
        
        Set W = Cairo.WidgetBase
        W.CanGetFocus = False
        
        Set m_Children = NewUIElementCollection(Me)
        
        Set m_Layout = New FrameworkElement
        m_Layout.Initialize Me, Me, W
        m_Layout.RegisterLayoutProperties "StackPanel"
        
        m_Orientation = OrientationVertical
        
        m_Layout.Width = 0
        m_Layout.Height = 0
        m_Layout.Visibility = VisibilityVisible
        
        NewUIElementBase Me, m_BaseUIElement
        
        'EnumerateProperties Me
        
        Set Style = GetBaseStyle
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


'#Region IUIElement Interface Implementation
    Private Property Get IUIElement_AttachedProperties() As ObservableDictionary
        Set IUIElement_AttachedProperties = m_BaseUIElement.AttachedProperties
    End Property
    
    Private Property Get IUIElement_Base() As UIElementBase
        Set IUIElement_Base = m_BaseUIElement
    End Property
    
    Private Property Set IUIElement_DataContext(ByVal RHS As Object)
        Set DataContext = RHS
    End Property
    
    Private Property Get IUIElement_DataContext() As Object
        Set IUIElement_DataContext = DataContext
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

        If Not e.OldItems Is Nothing Then
            For Each Child In e.OldItems
                If TypeOf Child Is IControl Then
                    If W.Widgets.Exists(modLayoutEngine.ControlWidgetKey(Child)) Then
                        W.Widgets.Remove modLayoutEngine.ControlWidgetKey(Child)
                    End If
                End If
            Next
        End If

        If Not e.NewItems Is Nothing Then
            For Each Child In e.NewItems
                MoveChild Child
            Next
        End If

        If Not W.LockRefresh Then W.Refresh
End Sub
'#End Region

'#Region Private Methods
Private Sub MoveChild(Child As Object)
    If W.LockRefresh Then Exit Sub
    ArrangePanel
End Sub

    Public Sub ArrangeOverride(ByVal FinalWidth As Double, ByVal FinalHeight As Double)
        ' Orientation DP is set in Class_Initialize before W/m_Layout exist;
        ' old ArrangeStackPanelChildren no-op'd on HostWidget Is Nothing — keep that.
        If W Is Nothing Then Exit Sub
        If m_Layout Is Nothing Then Exit Sub
        If m_Children Is Nothing Then Exit Sub

        MeasureOverride FinalWidth, FinalHeight
        ArrangeStackPanelChildren m_Children, W, m_Orientation, CSng(FinalWidth), CSng(FinalHeight)
        m_Layout.ReportActualSize FinalWidth, FinalHeight
    End Sub

    Private Sub ArrangePanel()
        If W Is Nothing Then Exit Sub
        ArrangeOverride CDbl(W.Width), CDbl(W.Height)
    End Sub

    Friend Sub RelayoutChildren()
        ArrangePanel
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        Select Case Property.Name
            Case "Style":           ApplyStyle
            Case "Orientation"
                m_Orientation = ParseOrientation(CStr(Property.GetValue))
                ArrangePanel
            Case Else
                If Not m_Children Is Nothing Then
                    If m_Layout.HandlesLayoutProperty(Property.Name) Then
                        If Property.Name = "Visibility" Then
                            m_Layout.OnLayoutPropertyChanged Property.Name, m_Children, W
                        Else
                            ArrangePanel
                        End If
                    End If
                End If
        End Select
        
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Sub
'#End Region

'#Region Widget Events
    Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
        
        Dim dx As Double, dy As Double
        
        dx = dx_Aligned
        dy = dy_Aligned
        
        
        CC.Save
        
        CC.Restore
    End Sub
    
    Private Sub W_Resize()
        ArrangePanel
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

