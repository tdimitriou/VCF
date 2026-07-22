VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' §2.11 / P6i — ContentPresenter as visual-tree element (WPF-aligned).
' Still exposes Draw() for host paint path when not parented in a template tree.

Implements IDependencyObject
Implements IUIElement
Implements IControl

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private m_Content As Variant
Private m_SuppressContent As Boolean
Private m_HorizontalContentAlignment As Long
Private m_VerticalContentAlignment As Long

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_DesignWidth As Single
Private m_DesignHeight As Single
Private m_ParentPtr As Long
Private m_TemplatedParentPtr As Long
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1

Public Name As String

'#Region Class Events
Private Sub Class_Initialize()
    Set m_Bindings = New List
    Set m_DependencyProperties = NewDependencyProperties(Me)
    Set W = Cairo.WidgetBase
    W.LockRefresh = True
    W.CanGetFocus = False
    W.ImplementsHitTest = True
    W.Visible = True

    With m_DependencyProperties
        Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
    End With

    m_Content = Empty
    m_SuppressContent = False
    m_HorizontalContentAlignment = AlignmentConstants.vbCenter
    m_VerticalContentAlignment = 2

    m_DesignLeft = 0
    m_DesignTop = 0
    m_DesignWidth = 100
    m_DesignHeight = 30

    NewUIElementBase Me, m_BaseUIElement
    W.LockRefresh = False
End Sub
'#End Region

'#Region Content API
Public Property Get Content() As Variant
    Call API.CopyVariable(m_Content, Content)
End Property

Public Property Let Content(ByVal Value As Variant)
    If IsObject(Value) Then
        Set m_Content = Value
    Else
        m_Content = Value
    End If
    RequestPaint
End Property

Public Property Set Content(ByVal Value As Object)
    Set m_Content = Value
    RequestPaint
End Property

Public Property Get SuppressContent() As Boolean
    SuppressContent = m_SuppressContent
End Property

Public Property Let SuppressContent(ByVal Value As Boolean)
    m_SuppressContent = Value
    RequestPaint
End Property

Public Property Get HorizontalContentAlignment() As Long
    HorizontalContentAlignment = m_HorizontalContentAlignment
End Property

Public Property Let HorizontalContentAlignment(ByVal Value As Long)
    m_HorizontalContentAlignment = Value
    RequestPaint
End Property

Public Property Get VerticalContentAlignment() As Long
    VerticalContentAlignment = m_VerticalContentAlignment
End Property

Public Property Let VerticalContentAlignment(ByVal Value As Long)
    m_VerticalContentAlignment = Value
    RequestPaint
End Property

Public Function ContentCaption() As String
    Dim V As Variant

    Call API.CopyVariable(m_Content, V)
    If IsObject(V) Then Exit Function
    If IsNull(V) Or IsEmpty(V) Then Exit Function
    ContentCaption = CStr(V)
End Function

Public Property Get WouldDrawCaption() As Boolean
    WouldDrawCaption = (Not m_SuppressContent) And (Len(ContentCaption()) > 0)
End Property

' Paint caption into [Left,Top]-[Right,Bottom]. Used by host DrawContent and W_Paint.
Public Sub Draw( _
    ByVal CC As cCairoContext, _
    ByVal Left As Double, _
    ByVal Top As Double, _
    ByVal Right As Double, _
    ByVal Bottom As Double, _
    ByVal FontName As String, _
    ByVal FontSize As Double, _
    ByVal ForeColor As Long, _
    Optional ByVal FontBold As Boolean = False, _
    Optional ByVal FontItalic As Boolean = False, _
    Optional ByVal FontUnderline As Boolean = False, _
    Optional ByVal FontStrikeThrough As Boolean = False)

    Dim Cap As String

    If m_SuppressContent Then Exit Sub
    Cap = ContentCaption()
    If Len(Cap) = 0 Then Exit Sub

    CC.SelectFont FontName, FontSize, ForeColor, FontBold, FontItalic, FontUnderline, FontStrikeThrough
    CC.DrawText Left, Top, Right, Bottom, Cap, False, _
        m_HorizontalContentAlignment, , MapCairoVAlign(m_VerticalContentAlignment), dtHasAccelerators
End Sub
'#End Region

'#Region IDependencyObject
Public Property Get DependencyProperties() As DependencyProperties
    Set DependencyProperties = m_DependencyProperties
End Property

Public Property Get Bindings() As VCF.List
    Set Bindings = m_Bindings
End Property

Private Property Get IDependencyObject_Children() As Object
    Set IDependencyObject_Children = Nothing
End Property

Private Property Get IDependencyObject_DependencyProperties() As DependencyProperties
    Set IDependencyObject_DependencyProperties = DependencyProperties
End Property

Private Property Get IDependencyObject_Parent() As IDependencyObject
    Set IDependencyObject_Parent = Parent
End Property
'#End Region

'#Region IControl
Private Property Get IControl_Children() As UIElementCollection
    Set IControl_Children = Children
End Property

Private Property Get IControl_Widget() As vbRichClient5.cWidgetBase
    Set IControl_Widget = Widget
End Property

Private Property Get IControl_Widgets() As vbRichClient5.cWidgets
    Set IControl_Widgets = Widgets
End Property

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

'#Region IUIElement
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

Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
    If W Is Nothing Then Exit Sub
    W.Move Left, Top, Width, Height
End Sub

Public Property Set Parent(ByVal Value As IControl)
    If Parent Is Value Then Exit Property
    m_ParentPtr = ObjPtr(Value)
    If Value Is Nothing Then Exit Property
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

'#Region Widget Events
Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean)
    ' Pass hits through so Button chrome still receives clicks (lookless caption).
    HitResultHit = False
End Sub

Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    If (Not W.Visible) Or W.IsHidden Then Exit Sub
    Draw CC, 0, 0, CDbl(dx_Aligned), CDbl(dy_Aligned), _
        W.FontName, W.FontSize, W.ForeColor, _
        W.FontBold, W.FontItalic, W.FontUnderline, W.FontStrikeThrough
End Sub
'#End Region

'#Region Helpers
Private Function MapCairoVAlign(ByVal Value As Long) As Long
    Select Case Value
        Case 0
            MapCairoVAlign = 0
        Case 1
            MapCairoVAlign = 2
        Case Else
            MapCairoVAlign = 1
    End Select
End Function

Private Sub RequestPaint()
    On Error Resume Next
    If Not W Is Nothing Then
        If Not W.LockRefresh Then W.Refresh
    End If
    Err.Clear
End Sub

' Copy typography from templated parent widget (Button) before nesting.
Friend Sub SyncTypographyFromWidget(ByVal Src As cWidgetBase)
    On Error Resume Next
    If Src Is Nothing Then Exit Sub
    If W Is Nothing Then Exit Sub
    W.FontName = Src.FontName
    W.FontSize = Src.FontSize
    W.ForeColor = Src.ForeColor
    W.FontBold = Src.FontBold
    W.FontItalic = Src.FontItalic
    W.FontUnderline = Src.FontUnderline
    W.FontStrikeThrough = Src.FontStrikeThrough
    Err.Clear
End Sub
'#End Region
