VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UnboundListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

Event OwnerDrawRowSelector(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event OwnerDrawHeader(CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event OwnerDrawItem(ByVal Index As Long, CC As cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
Event Click()
Event HeaderClick(ByVal ColIdx As Long, State As ColumnSortState)
Event DblClick()
Event DeleteKeyPressed()
Event ScrollIndexChange()
Event HScrollChange()
Event MouseUpClick()
Event MouseMoveOnListItem(ByVal HoverIndex As Long, ByVal RelX As Single, ByVal RelY As Single)
Event SelectedAll() 'Jürgen
Event DimensionsAdjusted()
Event ListIndexChanged()

Private WithEvents W                            As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private m_DesignLeft                            As Single
Private m_DesignTop                             As Single
Private m_DesignWidth                           As Single
Private m_DesignHeight                          As Single
Private m_Visibility                            As Visibility
Private m_ParentPtr                             As Long
Private m_Bindings                              As List
Private m_BaseUIElement                         As UIElementBase
Private m_Background                            As Long
Private m_SelectedBackground                    As Long
Private m_SelectedForeground                    As Long

Private WithEvents m_DependencyProperties       As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1

Private WithEvents m_Base                       As ListViewBase
Attribute m_Base.VB_VarHelpID = -1

Private m_Name                                  As String

'#Region Public Properties
    Public Property Get Base() As ListViewBase
        Set Base = m_Base
    End Property
    
    Public Property Get Resources() As ObservableDictionary
        Set Resources = m_BaseUIElement.Resources
    End Property
    
    Public Property Get Background() As Long
        Background = m_Background
    End Property
    
    Public Property Let Background(ByVal Value As Long)
        m_Background = Value
    End Property
    
    Public Property Get SelectedBackground() As Long
        SelectedBackground = m_SelectedBackground
    End Property
    
    Public Property Let SelectedBackground(ByVal Value As Long)
        m_SelectedBackground = Value
    End Property
    
    Public Property Get SelectedForeground() As Long
        SelectedForeground = m_SelectedForeground
    End Property
    
    Public Property Let SelectedForeground(ByVal Value As Long)
        m_SelectedForeground = Value
    End Property
    
    Public Property Get RowHeight() As Long
        RowHeight = m_Base.RowHeight
    End Property
    
    Public Property Let RowHeight(ByVal Value As Long)
        m_Base.RowHeight = Value
    End Property
    
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Get Name() As String
        Name = m_Name
    End Property
    
    Public Property Let Name(ByVal Value As String)
        m_Name = Value
    End Property
'#End Region

'#Region Public Methods
    Public Sub Refresh()
        m_Base.Widget.Refresh
    End Sub
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

    Public Property Let Visibility(ByVal Value As Visibility)
        m_Visibility = Value
        
        SetVisibility W, m_Visibility
    End Property
    
    Public Property Get Visibility() As Visibility
        Visibility = m_Visibility
    End Property

    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        If Parent Is Nothing Then Exit Sub
        
        If Parent.Widgets.Exists(Widget.Key) Then Parent.Widgets.Remove Widget.Key
        Parent.Widgets.Add Me, "_" & ObjPtr(Me), Left, Top, Width, Height
        
        SetVisibility W, m_Visibility
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
        Set m_DependencyProperties = NewDependencyProperties(Me)
        
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
        End With
        
        Set m_Base = New ListViewBase
        Set W = m_Base.Widget
        W.CanGetFocus = False
        
        DesignLeft = 0
        DesignTop = 0
        DesignWidth = 100
        DesignHeight = 30
        
        m_SelectedBackground = vbBlue
        m_SelectedForeground = vbWhite
        m_Background = vbWhite
        
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
        Set Children = Nothing
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

'#Region ListViewBase Events
    Private Sub m_Base_Click()
        RaiseEvent Click
    End Sub
    
    Private Sub m_Base_DblClick()
        RaiseEvent DblClick
    End Sub
    
    Private Sub m_Base_DeleteKeyPressed()
        RaiseEvent DeleteKeyPressed
    End Sub
    
    Private Sub m_Base_DimensionsAdjusted()
        RaiseEvent DimensionsAdjusted
    End Sub
    
    Private Sub m_Base_HeaderClick(ByVal ColIdx As Long, State As ColumnSortState)
        RaiseEvent HeaderClick(ColIdx, State)
    End Sub
    
    Private Sub m_Base_HScrollChange()
        RaiseEvent HScrollChange
    End Sub
    
    Private Sub m_Base_ListIndexChanged()
        RaiseEvent ListIndexChanged
    End Sub
    
    Private Sub m_Base_MouseMoveOnListItem(ByVal HoverIndex As Long, ByVal RelX As Single, ByVal RelY As Single)
        RaiseEvent MouseMoveOnListItem(HoverIndex, RelX, RelY)
    End Sub
    
    Private Sub m_Base_MouseUpClick()
        RaiseEvent MouseUpClick
    End Sub
    
    Private Sub m_Base_OwnerDrawHeader(CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
        RaiseEvent OwnerDrawHeader(CC, dx, dy, Alpha)
    End Sub
    
    Private Sub m_Base_OwnerDrawItem(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
        RaiseEvent OwnerDrawItem(Index, CC, dx, dy, Alpha)
    End Sub
    
    Private Sub m_Base_OwnerDrawRowSelector(CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
        RaiseEvent OwnerDrawRowSelector(CC, dx, dy, Alpha)
    End Sub
    
    Private Sub m_Base_ScrollIndexChange()
        RaiseEvent ScrollIndexChange
    End Sub
    
    Private Sub m_Base_SelectedAll()
        RaiseEvent SelectedAll
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        '
    End Sub
'#End Region

'#Region Private Methods
    Private Sub MoveChild(Child As Object)
        On Error Resume Next
        
        Dim xFactor As Double, yFactor As Double
        
        xFactor = W.Width / DesignWidth
        yFactor = W.Height / DesignHeight
        
        
        With Child
            
            .Move (.DesignLeft) * xFactor, (.DesignTop) * yFactor, (.DesignWidth) * xFactor, (.DesignHeight) * yFactor
        
        End With
    End Sub
'#End Region

