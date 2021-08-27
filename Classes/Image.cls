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
Implements IUIElement
Implements IVisualChild

Private m_Left As Double, m_Top As Double, m_Width As Double, m_Height As Double

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_DesignWidth As Single
Private m_DesignHeight As Single
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
            Call .Register("ShowGridLines", vbBoolean)
        End With
    
        m_DesignLeft = 0
        m_DesignTop = 0
        m_DesignWidth = 30
        m_DesignHeight = 30
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
    
        If Not m_KeepAspectRatio Then
        
            ' Just Pass the values given from the parent
             
            m_Left = Left
            m_Top = Top
            m_Width = Width
            m_Height = Height
        
        Else
                    
            'Scale to the smallest (x or y) scale factor
            
            Dim xFactor As Double, yFactor As Double
    
            xFactor = Width / DesignWidth
            yFactor = Height / DesignHeight
    
            Dim ScaleFactor As Double
            ' Choose the smallest Scale Factor
            ScaleFactor = IIf(xFactor > yFactor, yFactor, xFactor)
    
            Dim Cnt As IUIElement
            Set Cnt = Parent
            Dim VAlignCenter As Boolean
            VAlignCenter = (Cnt.DesignHeight - DesignHeight = 2 * DesignTop)
    
            m_Left = Left
            m_Width = DesignWidth * ScaleFactor
            m_Height = DesignHeight * ScaleFactor
    
            If VAlignCenter Then
                m_Top = (Parent.Widget.Height - m_Height) / 2
            Else
                m_Top = Top
            End If
    
        End If
    
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
    
        CC.RenderSurfaceContent Cairo.ImageList(m_ImageKey), m_Left, m_Top, m_Width, m_Height
End Sub
'#End Region

    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        If Parent Is Nothing Then Exit Sub
        If Property.Metadata.AffectsRender Then Parent.Widget.Refresh
    End Sub
