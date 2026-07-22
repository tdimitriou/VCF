VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Image"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

'Private m_Left As Double, m_Top As Double, m_Width As Double, m_Height As Double

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private m_ParentPtr As Long
Private m_BaseUIElement As UIElementBase

Private m_ImageKey As String
Private m_KeepAspectRatio As Boolean
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List

Public Name As String

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Bindings = New List
        Set m_DependencyProperties = NewDependencyProperties(Me)
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
        End With
    
        Set W = Cairo.WidgetBase
        W.ImplementsHitTest = True
        W.Visible = True
        
        m_KeepAspectRatio = True
        
        NewUIElementBase Me, m_BaseUIElement
    End Sub
'#End Region

'#Region IDependencyObject Interface Implementation
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

'#Region IDependencyObject Inheritance
    Public Property Get DependencyProperties() As DependencyProperties
        Set DependencyProperties = m_DependencyProperties
    End Property
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
    
    
    
    
    
    
    
    
    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        
        If Parent Is Nothing Then Exit Sub
        
        If Parent.Widgets.Exists(Widget.Key) Then Parent.Widgets.Remove Widget.Key
        Parent.Widgets.Add Me, "_" & ObjPtr(Me), , , , , W.Visible
        W.Move Left, Top, Width, Height
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        If Parent Is Value Then Exit Property
        
        m_ParentPtr = ObjPtr(Value)
                
        If Value Is Nothing Then Exit Property
        
        DependencyPropertiesStatic.InheritPropertyValues Me
        
        Parent.Widget.Refresh
    End Property
    
    Public Property Get Parent() As IControl
        Call API.CopyVariable(API.ObjFromPtr(m_ParentPtr), Parent)
    End Property
'#End Region

'#Region Public Properties
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Let ImageKey(ByVal Value As String)
        If m_ImageKey = Value Then Exit Property
        
        m_ImageKey = Value
        
        If Parent Is Nothing Then Exit Property
        
        Parent.Widget.Refresh
    End Property
    
    Public Property Get ImageKey() As String
        ImageKey = m_ImageKey
    End Property

    Public Property Get KeepAspectRatio() As Boolean
        KeepAspectRatio = m_KeepAspectRatio
    End Property
    
    Public Property Let KeepAspectRatio(ByVal Value As Boolean)
        m_KeepAspectRatio = Value
    End Property
'#End Region

'#Region IVIsualChild Interface Implementation
    Private Sub IVisualChild_DrawOn(CC As vbRichClient5.cCairoContext, Optional ForeColor As Long = -1)
        DrawOn CC, ForeColor
    End Sub
'#End Region

'#Region IVIsualChild Inheritance
    Public Sub DrawOn(CC As cCairoContext, Optional ForeColor As Long = -1)
        If Len(m_ImageKey) = 0 Then Exit Sub
        If Not Cairo.ImageList.Exists(m_ImageKey) Then Exit Sub
    
        Dim LL As Double, TT As Double, WW As Double, HH As Double
        
             
        LL = 0
        TT = 0
        WW = W.Width
        HH = W.Height
    
        If m_KeepAspectRatio Then
            Dim NatW As Double, NatH As Double
            Dim xFactor As Double, yFactor As Double
            Dim ScaleFactor As Double
            NatW = Cairo.ImageList(m_ImageKey).Width
            NatH = Cairo.ImageList(m_ImageKey).Height
            If NatW <= 0 Then NatW = W.Width
            If NatH <= 0 Then NatH = W.Height
            xFactor = W.Width / NatW
            yFactor = W.Height / NatH
            ScaleFactor = IIf(xFactor > yFactor, yFactor, xFactor)
            WW = NatW * ScaleFactor
            HH = NatH * ScaleFactor
            LL = (W.Width - WW) / 2
            TT = (W.Height - HH) / 2
        End If
        
        CC.RenderSurfaceContent Cairo.ImageList(m_ImageKey), LL, TT, WW, HH
End Sub
'#End Region

    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        If Parent Is Nothing Then Exit Sub
        If Property.Metadata.AffectsRender Then W.Refresh ' Parent.Widget.Refresh
    End Sub

Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean)
    HitResultHit = False
End Sub

Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    DrawOn CC
End Sub
