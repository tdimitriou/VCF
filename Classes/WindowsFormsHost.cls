VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "WindowsFormsHost"
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
Private WithEvents m_Children                   As ObservableCollection
Attribute m_Children.VB_VarHelpID = -1

Private m_DesignLeft                            As Single
Private m_DesignTop                             As Single
Private m_DesignWidth                           As Single
Private m_DesignHeight                          As Single
Private m_Visibility                            As Visibility
Private m_ParentPtr                             As Long
Private WithEvents m_DependencyProperties       As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings                              As List
Private m_BaseUIElement                         As UIElementBase
Private m_AutomaticallyUnloadContent            As Boolean

Public Name As String

'#Region Public Properties
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Get AutomaticallyUnloadContent() As Boolean
        AutomaticallyUnloadContent = m_AutomaticallyUnloadContent
    End Property
    
    Public Property Let AutomaticallyUnloadContent(ByVal Value As Boolean)
        m_AutomaticallyUnloadContent = Value
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
        End With
        
        Set W = Cairo.WidgetBase
        W.CanGetFocus = False
        
        DesignLeft = 0
        DesignTop = 0
        DesignWidth = 300
        DesignHeight = 300
        
        NewUIElementBase Me, m_BaseUIElement
                
        Set m_Children = New ObservableCollection
    End Sub

    Private Sub Class_Terminate()
        On Error Resume Next
        
        m_Children.Clear
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
    Public Property Get Children() As ObservableCollection
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
            
Try:
        On Error GoTo Catch

        If m_AutomaticallyUnloadContent Then
            UnloadObjects e.OldItems
        Else
            HideObjects e.OldItems
        End If
        
        W.Refresh
        
        Exit Sub
        
Catch:
        '
End Sub
'#End Region

'#Region Helper Functions
    Private Sub UnloadObjects(ByVal List As VCF.List)
        On Error Resume Next
        
        If List Is Nothing Then Exit Sub
        
        Dim Obj As Object
                
        For Each Obj In List
            UnloadObj Obj
        Next
    End Sub
    
    Private Sub HideObjects(ByVal List As VCF.List)
        On Error Resume Next
        
        If List Is Nothing Then Exit Sub
        
        Dim Obj As Object
                
        For Each Obj In List
            modWindowsFormsHostHelper.ShowWindow Obj, False
        Next
    End Sub
    
    Private Sub UnloadObj(ByVal Obj As Object)
        On Error Resume Next
        
        If TypeOf Obj Is VB.Form Then
            Unload Obj
        ElseIf TypeOf Obj Is vbRichClient5.cWidgetForm Then
            Obj.Unload
        Else
            ' Try to unload the object
            Unload Obj
        End If
    End Sub
'#End Region
    
'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        '
    End Sub
'#End Region

'#Region Widget Events
    Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
        Dim dx As Double, dy As Double
        
        dx = dx_Aligned
        dy = dy_Aligned
        
        CC.Rectangle 0, 0, dx, dy
        CC.SetSourceColor W.BackColor, W.Alpha
        CC.Fill
        
        If m_Children.Count = 0 Then Exit Sub
        modWindowsFormsHostHelper.SetChild m_Children(0), Me
    End Sub
'#End Region



