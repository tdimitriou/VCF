VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Button"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

Private WithEvents W                            As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents m_Children                   As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private m_DesignLeft                            As Single
Private m_DesignTop                             As Single
Private m_ParentPtr                             As Long

Private m_IsMouseOver                           As Boolean
Private IsMouseDown                             As Boolean

Private m_CornerRadius                          As Double
Private m_GradientBackground                    As Boolean
Private m_BorderWidth                           As Double
Private m_ClickMode                             As ClickMode
Private WithEvents m_DependencyProperties       As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings                              As VCF.List
Private m_BaseUIElement                         As UIElementBase
Private m_OverlayWidget                         As OverlayWidget
Private m_Layout                                As FrameworkElement

Private Const OVERLAY_ALPHA                     As Single = 0.25
Private Const OVERLAY_COLOR                     As Long = vbBlack

Public Name As String

Private WithEvents TmrRelease As cTimer
Attribute TmrRelease.VB_VarHelpID = -1

'#Region Public Properties
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Get Selected() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("Selected"), Selected)
    End Property

    Public Property Let Selected(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("Selected", Value)
    End Property
    
    Public Property Get Style() As Style
        Call API.CopyVariable(m_DependencyProperties.GetValue("Style"), Style)
    End Property

    Public Property Set Style(ByVal Value As Style)
        Call m_DependencyProperties.SetValue("Style", Value)
    End Property
    
    Public Property Get Margin() As Thickness
        Call API.CopyVariable(m_DependencyProperties.GetValue("Margin"), Margin)
    End Property

    Public Property Set Margin(ByVal Value As Thickness)
        Call m_DependencyProperties.SetValue("Margin", Value)
    End Property
    
    Public Property Get Command() As ICommand
        Call API.CopyVariable(m_DependencyProperties.GetValue("Command"), Command)
    End Property
    
    Public Property Set Command(ByVal Value As ICommand)
        Call m_DependencyProperties.SetValue("Command", Value)
    End Property
    
    Public Property Get CommandParameter() As Variant
        Call API.CopyVariable(m_DependencyProperties.GetValue("CommandParameter"), CommandParameter)
    End Property
    
    Public Property Let CommandParameter(ByVal Value As Variant)
        Call m_DependencyProperties.SetValue("CommandParameter", Value)
    End Property
    
    Public Property Set CommandParameter(ByVal Value As Variant)
        Call m_DependencyProperties.SetValue("CommandParameter", Value)
    End Property

    Public Property Get Content() As Variant
        Call API.CopyVariable(m_DependencyProperties.GetValue("Content"), Content)
    End Property

    Public Property Let Content(ByVal Value As Variant)
        Call m_DependencyProperties.SetValue("Content", Value)
    End Property

    Public Property Set Content(ByVal Value As Object)
        Call m_DependencyProperties.SetValue("Content", Value)
    End Property
    
    Public Property Get ClickMode() As ClickMode
        ClickMode = m_ClickMode
    End Property
    
    Public Property Let ClickMode(ByVal Value As ClickMode)
        m_ClickMode = Value
    End Property

    Public Property Get IsMouseOver() As Boolean
        IsMouseOver = m_IsMouseOver
    End Property

    Public Property Let IsMouseOver(ByVal Value As Boolean)
        If m_IsMouseOver = Value Then Exit Property
        m_IsMouseOver = Value
        ApplyStyle
    End Property

    Public Property Get BorderWidth() As Double
        BorderWidth = m_BorderWidth
    End Property
    
    Public Property Let BorderWidth(ByVal Value As Double)
        If m_BorderWidth = Value Then Exit Property
        
        m_BorderWidth = Value
        
        W.Refresh
    End Property
    
    Public Property Get GradientBackground() As Boolean
        GradientBackground = m_GradientBackground
    End Property
    
    Public Property Let GradientBackground(ByVal Value As Boolean)
        If m_GradientBackground = Value Then Exit Property
        
        m_GradientBackground = Value
        
        W.Refresh
    End Property
    
    Public Property Get CornerRadius() As Double
        CornerRadius = m_CornerRadius
    End Property
    
    Public Property Let CornerRadius(ByVal Value As Double)
        If m_CornerRadius = Value Then Exit Property
        
        m_CornerRadius = Value
        m_OverlayWidget.CornerRadious = m_CornerRadius
        
        W.Refresh
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
    
    Public Property Let DesignHeight(ByVal Value As Double)
        m_Layout.Height = Value
    End Property
    
    Public Property Get DesignHeight() As Double
        DesignHeight = m_Layout.Height
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
        m_Layout.Width = Value
    End Property
    
    Public Property Get DesignWidth() As Double
        DesignWidth = m_Layout.Width
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

    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        
        'If Parent.Widgets.Exists(Widget.Key) Then Parent.Widgets.Remove Widget.Key
        'Parent.Widgets.Add Me, "_" & ObjPtr(Me), Left + Margin.Left, Top + Margin.Top, Width - (Margin.Left + Margin.Right), Height - (Margin.Top + Margin.Bottom), DependencyProperties.GetValue("Visible")
            
        'Debug.Print "Start Move"
        W.Move Left, Top, Width, Height
        'Debug.Print "End Move"
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
        On Error Resume Next
                
        Set m_Bindings = New List
        Set m_Children = NewUIElementCollection(Me)
        Set m_DependencyProperties = NewDependencyProperties(Me)
                
        ' Initialize the WidgetBase (W) just in case we want to use it as
        ' the source or target for the dependency properties
        Set W = Cairo.WidgetBase
                
        W.LockRefresh = True
                
        ' Add the overlay widget inside the button control
        ' to use for applying various drawing effects.
        ' It should finally be removed and the effects
        ' applied directly to the widget
        Set m_OverlayWidget = New OverlayWidget
        W.Widgets.Add m_OverlayWidget, ObjPtr(m_OverlayWidget) & "_OLW"
        m_OverlayWidget.Alpha = OVERLAY_ALPHA
        m_OverlayWidget.Color = OVERLAY_COLOR
                        
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("Visible", vbBoolean)
            Call .Register("Margin", vbObject, "Thickness")
            Call .Register("Command", vbObject, "ICommand")
            Call .Register("CommandParameter", vbVariant)
            Call .Register("Selected", vbBoolean, , , , NewDependencyPropertyMetadata(False, True, True))
            Call .Register("BackColor", vbLong, , , , NewDependencyPropertyMetadata(False, True, True))
            Call .Register("BorderColor", vbLong, , , , NewDependencyPropertyMetadata(False, True, True))
            Call .Register("ToolTip", vbString, , , , NewDependencyPropertyMetadata(False, True, False))
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            Call .Register("Content", vbVariant, , , , NewDependencyPropertyMetadata(False, True, False, OneWay))

            .SetValue "Style", Nothing
            
            ' Initialize the default CurrentValues of the Dependency Properties
            ' just in case the default style is not present.
            ' This should finally be replaced by built-in default styles
            ' inside the framework, which can be overwritten by the consumer application
            .SetCurrentValue "Visible", True
            .SetCurrentValue "Margin", NewThickness(0)
            .SetCurrentValue "Selected", False
            .SetCurrentValue "BackColor", W.BackColor
            .SetCurrentValue "BorderColor", W.BorderColor
            .SetCurrentValue "ToolTip", W.ToolTip

        
        End With
                        
        m_BorderWidth = 1
        W.HoverColor = -1
        W.CanGetFocus = False
        
        Set m_Layout = New FrameworkElement
        m_Layout.Initialize Me, Me, W
        m_Layout.RegisterLayoutProperties "Button"
        
        m_DesignLeft = 0
        m_DesignTop = 0
        m_Layout.Width = 100
        m_Layout.Height = 30
                
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

'#Region IDependencyObject Inheritance
    Public Property Get DependencyProperties() As DependencyProperties
        Set DependencyProperties = m_DependencyProperties
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
                
        If Not e.NewItems Is Nothing Then
            For Each Child In e.NewItems
                MoveChild Child
            Next
        End If
                
        W.Refresh
        
        Exit Sub
        
Catch:
        '
    End Sub
'#End Region

'#Region Private Methods
    Private Sub MoveChild(Child As Object)
        On Error Resume Next
        
        'Debug.Print "Start Move Child"
        
        Dim xFactor As Double, yFactor As Double
        Dim Offs As Single
        
        
        Offs = m_BorderWidth
        If IsMouseDown Then Offs = BorderWidth + 1
        
        xFactor = W.Width / m_Layout.Width
        yFactor = W.Height / m_Layout.Height
        
        
        'If W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Remove Child.Widget.Key
        
        Dim LL As Single, TT As Single, WW As Single, HH As Single
        
        With Child
            
            LL = (.DesignLeft + Offs) * xFactor
            TT = (.DesignTop + Offs) * yFactor
            WW = (.DesignWidth - 2 * Offs) * xFactor
            HH = (.DesignHeight - 2 * Offs) * yFactor
            
            If Not W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Add Child, "_" & ObjPtr(Child), CDbl(LL), CDbl(TT), CDbl(WW), CDbl(HH)
            
            .Move LL, TT, WW, HH
        
        End With
                    
        'Debug.Print "End Move Child"
                    
    End Sub
    
    Private Function IsInside(ByVal x As Single, ByVal y As Single) As Boolean
        IsInside = (x >= 0 And y >= 0 And x <= W.Width And y <= W.Height)
    End Function
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        On Error Resume Next
                
        ' Pass the changed property value to the matching widget
        ' property (if any) to use in the drawing routines
        Select Case Property.Name
            Case "Visible":         W.Visible = Property.GetValue
            Case "Selected":        modRenderCoalescer.RequestWidgetRefresh W
            Case "BackColor":       W.BackColor = Property.GetValue
            Case "BorderColor":     W.BorderColor = Property.GetValue
            Case "ToolTip":         W.ToolTip = Property.GetValue
            Case "Style":           ApplyStyle
            Case "Content":         modRenderCoalescer.RequestWidgetRefresh W
        End Select
        
        ' Pass the new property value to the contained widgets
        ' The 'PassPropertyValue' Method is using the 'SetCurrentValue',
        ' instead of the SetValue Operation. This practice enables both
        ' inheritance (for inheritable properties, and also individual
        ' behaviour / styling by customizing the properties by
        ' exclusively calling the 'SetValue' Method either from XAML or Code.
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Sub
'#End Region

Private Sub TmrRelease_Timer()
    Set TmrRelease = Nothing
    IsMouseDown = False
    Call W_Resize
    W.Refresh
End Sub

'#Region Widget Events
    'Private Sub W_ExitFocus()
    '    IsMouseOver = False
    '    IsMouseDown = False
    '    Call W_Resize
    '    W.Refresh
    'End Sub
    '
    'Private Sub W_Hide()
    '    IsMouseOver = False
    '    IsMouseDown = False
    '    Call W_Resize
    '    W.Refresh
    'End Sub
    
    Private Sub W_LostFocus()
        IsMouseOver = False
        IsMouseDown = False
        Call W_Resize
        W.Refresh
    End Sub
    
    Private Sub W_MouseDown(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Not W.Enabled Then Exit Sub
        
        IsMouseDown = (Button = 1)
        Call W_Resize
        W.Refresh
        
        If Not (Button = 1 And m_ClickMode = ClickModePress) Then Exit Sub
                
        
        Set TmrRelease = New_c.Timer(750, True)
        
        OnClick
        
        
        'IsMouseDown = False
        'Call W_Resize
        'W.Refresh
    End Sub
    
    Private Sub W_MouseEnter(ByVal MouseLeaveWidget As vbRichClient5.cWidgetBase)
        If Not W.Enabled Then Exit Sub
        
        IsMouseOver = True
        
        If m_ClickMode = ClickModeHover Then If Not IsMouseDown Then OnClick
    End Sub
    
    Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
        If Not W.Enabled Then Exit Sub
        
        IsMouseOver = False
        W.Refresh
    End Sub
    
    Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Not W.Enabled Then Exit Sub
        '
    End Sub

    Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
        If Not W.Enabled Then Exit Sub
                
        Set TmrRelease = Nothing
        
        IsMouseDown = False
        
        Call W_Resize
        W.Refresh
        
        If Button = 1 And IsInside(x, y) And m_ClickMode = ClickModeRelease Then OnClick
    End Sub

    Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
        
        If (Not W.Visible) Or W.IsHidden Then Exit Sub
        
        Dim dx As Double, dy As Double
        
        dx = dx_Aligned
        dy = dy_Aligned
        
        CC.Save
        
        ' Cache the value of the dependency property
        ' to avoid querying again and again
        Dim IsSelected As Boolean
        IsSelected = Selected
        
        Call DrawBackground(CC, dx, dy, IsSelected)
        Call DrawBorder(CC, dx, dy, IsSelected)
        Call DrawContent(CC, dx, dy)
        
        Call SetupOverlay(IsSelected)
        CC.Restore
                
    End Sub
    
    Private Sub W_Resize()
                
        Dim Child As Object
        
        For Each Child In m_Children
            Call MoveChild(Child)
        Next
                
    End Sub
'#End Region

'#Region Drawing Functions
    Private Function ContentCaption() As String
        Dim V As Variant

        Call API.CopyVariable(m_DependencyProperties.GetValue("Content"), V)
        If IsObject(V) Then Exit Function
        If IsNull(V) Or IsEmpty(V) Then Exit Function
        ContentCaption = CStr(V)
    End Function

    Private Sub DrawContent(CC As cCairoContext, dx As Double, dy As Double)
        Dim Cap As String
        Dim Offs As Single
        Dim L As Double
        Dim T As Double
        Dim Wd As Double
        Dim Ht As Double

        If m_Children.Count > 0 Then Exit Sub

        Cap = ContentCaption()
        If Len(Cap) = 0 Then Exit Sub

        Offs = CSng(m_BorderWidth)
        If IsMouseDown Then Offs = CSng(BorderWidth + 1)

        L = Offs
        T = Offs
        Wd = dx - 2# * Offs
        Ht = dy - 2# * Offs

        CC.SelectFont W.FontName, W.FontSize, W.ForeColor, W.FontBold, W.FontItalic, W.FontUnderline, W.FontStrikeThrough
        CC.DrawText L, T, L + Wd, T + Ht, Cap, False, AlignmentConstants.vbCenter, , 1, dtHasAccelerators
    End Sub

    Private Sub DrawBackground(CC As cCairoContext, dx As Double, dy As Double, ByVal Selected As Boolean)
        If Selected Or Not m_GradientBackground Then
            Call DrawBackgroundFlat(CC, dx, dy, Selected)
        Else
            Call DrawBackgroundGradient(CC, dx, dy, Selected)
        End If
    End Sub
    
    Private Sub DrawBackgroundGradient(CC As cCairoContext, dx As Double, dy As Double, ByVal Selected As Boolean)
        Dim Pat As cCairoPattern
        
        Set Pat = Cairo.CreateLinearPattern(0, 0, 0, dy)
        
        Dim f1 As Single, f2 As Single
        
        f1 = 1.25
        f2 = 0.75
        
        Dim Color As Long
        Color = W.BackColor
        If IsMouseOver And W.HoverColor <> -1 Then Color = W.HoverColor
        
        With Pat
            .AddColorStop 0, modStaticClasses.Color.Multiply(Color, f1), W.Alpha
            .AddColorStop 0.5, modStaticClasses.Color.Multiply(Color, f2), W.Alpha
            .AddColorStop 1, modStaticClasses.Color.Multiply(Color, f1), W.Alpha
        End With
        
        CC.RoundedRect 0, 0, dx, dy, m_CornerRadius
        CC.Fill , Pat
        
    End Sub
    
    Private Sub DrawBackgroundFlat(CC As cCairoContext, dx As Double, dy As Double, ByVal Selected As Boolean)
        CC.RoundedRect 0, 0, dx, dy, m_CornerRadius
        
        Dim Color As Long
        Color = W.BackColor
        If IsMouseOver And W.HoverColor <> -1 Then Color = W.HoverColor
        If Selected Then Color = W.SelectionColor
        
        CC.SetSourceColor Color, W.Alpha
        
        CC.Fill
    End Sub
        
    Private Sub SetupOverlay(ByVal Selected As Boolean)
        m_OverlayWidget.Visible = Not (Selected Or Not IsMouseDown)
        m_OverlayWidget.Alpha = IIf(W.Alpha > 0, OVERLAY_ALPHA, 0)
        m_OverlayWidget.Widget.MoveToFront
    End Sub

    Private Sub DrawBorder(CC As cCairoContext, dx As Double, dy As Double, ByVal Selected As Boolean)
        CC.RoundedRect 0, 0, dx, dy, CornerRadius, True
        CC.SetSourceColor W.BorderColor, W.Alpha * 0.5
        
        Dim LineWidth As Double
        LineWidth = m_BorderWidth
        
        If IsMouseOver And Not IsMouseDown Then LineWidth = 2 * LineWidth
        
        CC.SetLineWidth LineWidth
        CC.Stroke
    
    End Sub
'#End Region

Private Sub OnClick()
    On Error Resume Next
        
    If Not Command Is Nothing Then
        If Command.CanExecute Then Call Command.Execute(CommandParameter)
    End If
End Sub

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

