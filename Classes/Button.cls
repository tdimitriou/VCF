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

Private m_DesignLeft                            As Single
Private m_DesignTop                             As Single
Private m_DesignWidth                           As Single
Private m_DesignHeight                          As Single
Private m_ParentPtr                             As Long

Private IsMouseOver                             As Boolean
Private IsMouseDown                             As Boolean

Private m_CornerRadius                          As Double
Private m_GradientBackground                    As Boolean
Private m_BorderWidth                           As Double
Private m_ClickMode                             As ClickMode
Private WithEvents m_DependencyProperties       As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings                              As VCF.List
Private m_BaseUIElement                         As UIElementBase

Public Name As String

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
    
    Public Property Get ClickMode() As ClickMode
        ClickMode = m_ClickMode
    End Property
    
    Public Property Let ClickMode(ByVal Value As ClickMode)
        m_ClickMode = Value
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
        
        If Parent.Widgets.Exists(Widget.Key) Then Parent.Widgets.Remove Widget.Key
        Parent.Widgets.Add Me, "_" & ObjPtr(Me), Left + Margin.Left, Top + Margin.Top, Width - (Margin.Left + Margin.Right), Height - (Margin.Top + Margin.Bottom), DependencyProperties.GetValue("Visible")
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
        Set m_Bindings = New List
        Set m_Children = NewUIElementCollection(Me)
        Set m_DependencyProperties = NewDependencyProperties(Me)
        
        Set W = Cairo.WidgetBase
        
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("Visible", vbBoolean)
            .SetValue "Visible", True
            Call .Register("Margin", vbObject, "Thickness")
            Set Margin = NewThickness(0)
            Call .Register("Command", vbObject, "ICommand")
            Call .Register("CommandParameter", vbVariant)
            Call .Register("Selected", vbBoolean, , , , NewDependencyPropertyMetadata(False, True, True))
            .SetValue "Selected", False
            Call .Register("BackColor", vbLong, , , , NewDependencyPropertyMetadata(False, True, True))
            .SetValue "BackColor", W.BackColor
        End With
                
        m_DesignLeft = 0
        m_DesignTop = 0
        m_DesignWidth = 100
        m_DesignHeight = 30
            
        m_BorderWidth = 1
        W.HoverColor = -1
        W.CanGetFocus = False
        
        NewUIElementBase Me, m_BaseUIElement
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
        
        Dim xFactor As Double, yFactor As Double
        Dim Offs As Single
        
        
        Offs = m_BorderWidth
        If IsMouseDown Then Offs = BorderWidth + 1
        
        xFactor = W.Width / DesignWidth
        yFactor = W.Height / DesignHeight
        
        
        If W.Widgets.Exists(Child.Widget.Key) Then W.Widgets.Remove Child.Widget.Key
        
        
        With Child
            
            .Move (.DesignLeft + Offs) * xFactor, (.DesignTop + Offs) * yFactor, (.DesignWidth - 2 * Offs) * xFactor, (.DesignHeight - 2 * Offs) * yFactor
        
        End With
    End Sub
    
    Private Function IsInside(ByVal x As Single, ByVal y As Single) As Boolean
        IsInside = (x >= 0 And y >= 0 And x <= W.Width And y <= W.Height)
    End Function
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        On Error Resume Next
        
        Select Case Property.Name
            Case "Visible":         W.Visible = Property.GetValue
            Case "Selected":        W.Refresh
            Case "BackColor":       W.BackColor = Property.GetValue
        End Select
        
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Sub
'#End Region

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
        IsMouseDown = (Button = 1)
        Call W_Resize
        W.Refresh
        
        If Not (Button = 1 And m_ClickMode = ClickModePress) Then Exit Sub
                
        OnClick
        
        'IsMouseDown = False
        'Call W_Resize
        'W.Refresh
    End Sub
    
    Private Sub W_MouseEnter(ByVal MouseLeaveWidget As vbRichClient5.cWidgetBase)
        IsMouseOver = True
        W.Refresh
        
        If m_ClickMode = ClickModeHover Then If Not IsMouseDown Then OnClick
    End Sub
    
    Private Sub W_MouseLeave(ByVal MouseEnterWidget As vbRichClient5.cWidgetBase)
        IsMouseOver = False
        W.Refresh
    End Sub
    
    Private Sub W_MouseMove(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
        '
    End Sub

    Private Sub W_MouseUp(Button As Integer, Shift As Integer, ByVal x As Single, ByVal y As Single)
        IsMouseDown = False
        
        Call W_Resize
        W.Refresh
        
        If Button = 1 And IsInside(x, y) And m_ClickMode = ClickModeRelease Then OnClick
    End Sub

    Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
        Dim dx As Double, dy As Double
        
        dx = dx_Aligned
        dy = dy_Aligned
        
        CC.Save
        
        Call DrawBackground(CC, dx, dy)
        Call DrawVisualChildren(CC)
        Call DrawOverlay(CC, dx, dy)
        Call DrawBorder(CC, dx, dy)
        
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
    Private Sub DrawBackground(CC As cCairoContext, dx As Double, dy As Double)
        If Selected Or Not m_GradientBackground Then
            Call DrawBackgroundFlat(CC, dx, dy)
        Else
            Call DrawBackgroundGradient(CC, dx, dy)
        End If
    End Sub
    
    Private Sub DrawBackgroundGradient(CC As cCairoContext, dx As Double, dy As Double)
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
    
    Private Sub DrawBackgroundFlat(CC As cCairoContext, dx As Double, dy As Double)
        CC.RoundedRect 0, 0, dx, dy, m_CornerRadius
        
        Dim Color As Long
        Color = W.BackColor
        If IsMouseOver And W.HoverColor <> -1 Then Color = W.HoverColor
        If Selected Then Color = W.SelectionColor
        
        CC.SetSourceColor Color, W.Alpha
        
        CC.Fill
    End Sub
    
    Private Sub DrawVisualChildren(CC As cCairoContext)
        Dim Child As Object
        Dim Visual As IVisualChild
               
        For Each Child In m_Children
            If TypeOf Child Is IVisualChild Then
                Set Visual = Child
                Visual.DrawOn CC
            End If
        Next
    
    End Sub
    
    Private Sub DrawOverlay(CC As cCairoContext, dx As Double, dy As Double)
                
        If Selected Or Not IsMouseDown Then Exit Sub
        
        Dim Alpha As Double
        
        Alpha = W.Alpha
        If W.Alpha > 0 Then Alpha = 0.25
                
        CC.RoundedRect 0, 0, dx, dy, CornerRadius
        CC.SetSourceColor vbBlack, Alpha
        CC.Fill
        
    End Sub

    Private Sub DrawBorder(CC As cCairoContext, dx As Double, dy As Double)
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

