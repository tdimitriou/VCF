VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

Event Click()
Event Scroll()
Event Change()
Event OwnerDrawBackGround(CC As cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)
Event KeyDown(KeyCode As Integer, Shift As Integer)
Event KeyPress(KeyAscii As Integer)
Event KeyUp(KeyCode As Integer, Shift As Integer)
Event Validate(Cancel As Boolean)
Event MaxLengthViolation() ' An insert operation was cancelled because it would violate the MaxLength property value
Event BeforePaste(TextToPaste As String, Cancel As Boolean)
Event BeforeSelChange(NewSelStart As Long, NewSelLength As Long, SelBackwards As Boolean, Cancel As Boolean)
Event SelChanged(ByVal OldSelStart As Long, ByVal OldSelLength As Long, ByVal OldCaretPosition As Long)

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_DesignWidth As Single
Private m_DesignHeight As Single
Private m_Visibility As Visibility
Private m_ParentPtr As Long
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase

Private WithEvents m_Base As TextBoxBase
Attribute m_Base.VB_VarHelpID = -1
Private m_Focused As Boolean

Public Name As String

'#Region Public Properties
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property

    Public Property Let Text(ByVal Value As String)
        Call m_DependencyProperties.SetValue("Text", Value)
    End Property
    
    Public Property Get Text() As String
        Call API.CopyVariable(m_DependencyProperties.GetValue("Text"), Text)
    End Property

    Public Property Get Alignment() As AlignmentConstants
        Call API.CopyVariable(m_DependencyProperties.GetValue("Alignment"), Alignment)
    End Property
    
    Public Property Let Alignment(ByVal Value As AlignmentConstants)
        Call m_DependencyProperties.SetValue("Alignment", Value)
    End Property
    
    Public Property Get VCenter() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("VCenter"), Text)
    End Property
    
    Public Property Let VCenter(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("VCenter", Value)
    End Property
    
    Public Property Get PasswordChar() As String
        Call API.CopyVariable(m_DependencyProperties.GetValue("PasswordChar"), PasswordChar)
    End Property
    
    Public Property Let PasswordChar(ByVal Value As String)
        Call m_DependencyProperties.SetValue("PasswordChar", Value)
    End Property

    Public Property Get CueBanner() As String
        Call API.CopyVariable(m_DependencyProperties.GetValue("CueBanner"), CueBanner)
    End Property
    
    Public Property Let CueBanner(ByVal Value As String)
        Call m_DependencyProperties.SetValue("CueBanner", Value)
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
    
        SetFocus
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
            Call .Register("Text", vbString, , , "", NewDependencyPropertyMetadata(, True, , VCF.BindingMode.TwoWay))
            Call .Register("Alignment", vbLong, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("VCenter", vbString, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("PasswordChar", vbString, , , "", NewDependencyPropertyMetadata(, True))
            Call .Register("CueBanner", vbString, , , "", NewDependencyPropertyMetadata(, True))
        End With
        
        Set m_Base = New TextBoxBase
        Set W = m_Base.Widget
        m_Base.VCenter = True
        
        DesignLeft = 0
        DesignTop = 0
        DesignWidth = 100
        DesignHeight = 30
        
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

Public Property Get Border() As Boolean
    Border = m_Base.Border
End Property

Public Property Let Border(ByVal Value As Boolean)
    m_Base.Border = Value
End Property

Public Property Get Focused() As Boolean
    Focused = m_Focused
End Property

Public Property Let Focused(ByVal Value As Boolean)
    m_Focused = Value
    
    SetFocus
End Property

Private Sub SetFocus()
    On Error Resume Next
    
    If m_Focused Then W.SetFocus
End Sub

Private Sub m_Base_BeforePaste(TextToPaste As String, Cancel As Boolean)
    RaiseEvent BeforePaste(TextToPaste, Cancel)
End Sub

Private Sub m_Base_BeforeSelChange(NewSelStart As Long, NewSelLength As Long, SelBackwards As Boolean, Cancel As Boolean)
    RaiseEvent BeforeSelChange(NewSelStart, NewSelLength, SelBackwards, Cancel)
End Sub

Private Sub m_Base_Change()
    Me.Text = m_Base.Text
    RaiseEvent Change
End Sub

Private Sub m_Base_Click()
    RaiseEvent Click
End Sub

Private Sub m_Base_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub m_Base_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub m_Base_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub m_Base_MaxLengthViolation()
    RaiseEvent MaxLengthViolation
End Sub

Private Sub m_Base_OwnerDrawBackGround(CC As vbRichClient5.cCairoContext, ByVal x As Single, ByVal y As Single, ByVal dx As Single, ByVal dy As Single, CancelBuiltIn As Boolean)
    RaiseEvent OwnerDrawBackGround(CC, x, y, dx, dy, CancelBuiltIn)
End Sub

Private Sub m_Base_Scroll()
    RaiseEvent Scroll
End Sub

Private Sub m_Base_SelChanged(ByVal OldSelStart As Long, ByVal OldSelLength As Long, ByVal OldCaretPosition As Long)
    RaiseEvent SelChanged(OldSelStart, OldSelLength, OldCaretPosition)
End Sub

Private Sub m_Base_Validate(Cancel As Boolean)
    RaiseEvent Validate(Cancel)
End Sub

Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
    If Property.Name = "Text" Then m_Base.Text = Property.GetValue
    If Property.Name = "Alignment" Then m_Base.Alignment = Property.GetValue
    If Property.Name = "VCenter" Then m_Base.VCenter = Property.GetValue
    If Property.Name = "CueBanner" Then m_Base.CueBannerText = Property.GetValue
    If Property.Name = "PasswordChar" Then m_Base.PasswordChar = Property.GetValue
End Sub

'<TextBoxBase Properties>
'Public Property Get SelStart() As Long
'    SelStart = mSelStart
'End Property
'
'Public Property Let SelStart(ByVal NewValue As Long)
'    SetSelection NewValue, 0
'End Property
'
'Public Property Get SelLength() As Long
'    SelLength = mSelLength
'End Property
'
'Public Property Let SelLength(ByVal NewValue As Long)
'    SetSelection Me.SelStart, NewValue
'End Property
'
'Public Property Get SelText() As String
'    If mSelLength Then SelText = Mid$(mText, mSelStart + 1, mSelLength)
'End Property
'
'Public Property Let SelText(ByVal NewValue As String)
'    InsertText NewValue
'    W.Refresh
'End Property
'
'Public Property Get AutoSelectAll() As Boolean
'    AutoSelectAll = mAutoSelectAll
'End Property
'
'Public Property Let AutoSelectAll(ByVal AutoSelectAll As Boolean)
'    mAutoSelectAll = AutoSelectAll
'End Property
'
'Public Property Get HideSelection() As Boolean
'    HideSelection = mHideSelection
'End Property
'
'Public Property Let HideSelection(ByVal HideSelection As Boolean)
'    mHideSelection = HideSelection
'End Property
'</TextBoxBase Properties>

