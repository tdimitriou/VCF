VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentControl"
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

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_ParentPtr As Long
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase
Private m_Layout As FrameworkElement

Public Property Get Content() As Object
    If m_Children.Count > 0 Then
        Set Content = m_Children(0)
    End If
End Property

Public Property Set Content(ByVal Value As Object)
    Dim i As Long

    For i = m_Children.Count - 1 To 0 Step -1
        m_Children.RemoveAt i
    Next

    If Not Value Is Nothing Then
        If TypeOf Value Is IUIElement Then
            m_Children.Add Value
        End If
    End If
End Property

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
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Bindings = New List
        Set m_DependencyProperties = NewDependencyProperties(Me)
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ShowGridLines", vbBoolean)
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            .SetValue "Style", Nothing
        End With
        
        Set W = Cairo.WidgetBase
        W.CanGetFocus = False
        
        Set m_Children = NewUIElementCollection(Me)
        
        Set m_Layout = New FrameworkElement
        m_Layout.Initialize Me, Me, W
        m_Layout.RegisterLayoutProperties "ContentControl"
        m_Layout.LegacyScaleLayout = False
        
        DesignLeft = 0
        DesignTop = 0
        m_Layout.Width = 300
        m_Layout.Height = 300
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
        
        If Not W.LockRefresh Then W.Refresh
        
        Exit Sub
        
Catch:
        '
End Sub
'#End Region

'#Region Private Methods
    Private Sub MoveChild(Child As Object)
        ArrangeContent
    End Sub

    Private Sub ArrangeContent()
        If m_Children.Count > 0 Then
            ArrangeDecoratorChild m_Children(0), W
        End If
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        Select Case Property.Name
            Case "Style":           ApplyStyle
            Case Else
                If Not m_Children Is Nothing Then
                    If m_Layout.HandlesLayoutProperty(Property.Name) Then
                        ArrangeContent
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
        ArrangeContent
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

