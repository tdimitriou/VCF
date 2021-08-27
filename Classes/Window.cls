VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Window"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement

Public WithEvents Form As cWidgetForm
Attribute Form.VB_VarHelpID = -1
Private WithEvents WRoot As cWidgetRoot
Attribute WRoot.VB_VarHelpID = -1

Private WithEvents m_Children As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_DesignWidth As Single
Private m_DesignHeight As Single
Private m_NamedChildren As ObservableDictionary
Private m_Bindings As List
Private m_DialogResult As Boolean
Private m_BasePtr As Long
Private m_BaseUIElement As UIElementBase

Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1

Private WithEvents m_Subclass As cSubClass
Attribute m_Subclass.VB_VarHelpID = -1

Public Event WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                        ByVal Subclass As cSubClass, ByRef CancelDefaultProc As Boolean)

'#Region Form Events
    Private Sub Form_ResizeWithDimensions(ByVal NewWidth As Long, ByVal NewHeight As Long)
        Dim Child As Object
                
        For Each Child In m_Children
            Call MoveChild(Child)
        Next
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        On Error Resume Next
        
        m_Subclass.UnHook
        If m_BasePtr <> 0 Then Application.Current.Base.Windows.Remove m_BasePtr
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
    
    Public Property Get Widget() As vbRichClient5.cWidgetBase
        Set Widget = Form.WidgetRoot
    End Property
    
    Public Property Get Widgets() As vbRichClient5.cWidgets
        Set Widgets = Form.Widgets
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
        Form.Move Left, Top, Width, Height
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        '
    End Property
    
    Public Property Get Parent() As IControl
        Set Parent = Nothing
    End Property
'#End Region

'#Region Public Properties
    Public Property Get DialogResult() As Boolean
        DialogResult = m_DialogResult
    End Property
    
    Public Property Let DialogResult(ByVal Value As Boolean)
        m_DialogResult = Value
    End Property
    
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Get NamedChildren() As ObservableDictionary
        Set NamedChildren = m_NamedChildren
    End Property
'#End Region

'#Region Public Methods
    Public Sub Show(Optional Modal As FormShowConstants = vbModeless, Optional Owner As Object, Optional Focused As Boolean = True)
        Form.Show Modal, Owner, Not Focused
    End Sub
    
    Public Function ShowDialog(Optional Owner As Object) As Boolean
        Form.Show vbModal, Owner
        
        ShowDialog = m_DialogResult
    End Function
    
    Public Sub Dispose()
        On Error Resume Next
        
        If m_BasePtr <> 0 Then If Application.Current.Base.Windows.Exists(m_BasePtr) Then Application.Current.Base.Windows.Remove m_BasePtr
    
    End Sub
'#End Region

'#Region UIElementCollection Events
    Private Sub m_Children_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
        On Error GoTo Catch
            
        Dim Child As Object
        
Try:
        
       WRoot.LockRefresh = True
       
        If Not e.OldItems Is Nothing Then
            For Each Child In e.OldItems
                If TypeOf Child Is IControl Then If WRoot.Widgets.Exists(Child.Widget.Key) Then WRoot.Widgets.Remove Child.Widget.Key
            Next
        End If
                
        If Not e.NewItems Is Nothing Then
            For Each Child In e.NewItems
                MoveChild Child
            Next
        End If
        
        WRoot.LockRefresh = False
        
        Exit Sub
        
Catch:
        '
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Sub
'#End Region

'#Region Subclass Events
    Private Sub m_Subclass_WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
        
        'Enable F10 Key Functionality
        If Msg = 260 And wParam = 121 Then m_Subclass.CallWindowProc 256, wParam, lParam
        
        ' Pass the event to consumers to allow subclassing
        Dim Cancel As Boolean
        RaiseEvent WindowProc(Result, Msg, wParam, lParam, m_Subclass, Cancel)
        If Cancel = False Then m_Subclass.CallWindowProc Msg, wParam, lParam
    
        Result = 1
    End Sub
'#End Region

'#Region WidgetRoot Events
    Private Sub WRoot_Paint(CC As vbRichClient5.cCairoContext)
        DrawVisualChildren CC
    End Sub
'#End Region

'#Region Private Methods
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
    
    Private Sub MoveChild(Child As Object)
        On Error Resume Next
        
        Dim xFactor As Double, yFactor As Double
        
        xFactor = Form.ScaleWidth / DesignWidth
        yFactor = Form.ScaleHeight / DesignHeight
        
        With Child
            .Move .DesignLeft * xFactor, .DesignTop * yFactor, .DesignWidth * xFactor, .DesignHeight * yFactor
        End With
    End Sub

    Public Sub SetFocus()
        On Error Resume Next
        
        Form.WidgetRoot.SetFocus Nothing
        
        If m_Children.Count = 0 Then Exit Sub
        
        Dim FocusedObject As Object
        Dim FocusedControl As IControl
        
        Set FocusedObject = GetFocusedObject(m_Children(0))
                        
        If FocusedObject Is Nothing Or Not TypeOf FocusedObject Is IControl Then Exit Sub
                        
        Set FocusedControl = FocusedObject
        
        Form.WidgetRoot.SetFocus FocusedControl.Widget
    End Sub
    
    Private Function GetFocusedObject(ByVal o As Object) As Object
        On Error Resume Next
        
        Dim c As Object
        Dim IsFocused As Boolean
        
        If Err.Number <> 0 Then Err.Clear
        IsFocused = o.Focused
        If Err.Number <> 0 Then Err.Clear
        
        If IsFocused Then
            Set GetFocusedObject = o
            
            Exit Function
        End If
        
        For Each c In o.Widgets
            If Err.Number <> 0 Then Err.Clear
            IsFocused = c.Focused
            If Err.Number <> 0 Then Err.Clear
            
            If IsFocused Then
                Set GetFocusedObject = c
                Exit Function
            Else
                Set GetFocusedObject = GetFocusedObject(c)
                If Not GetFocusedObject Is Nothing Then
                    Exit Function
                End If
            End If
        Next
    End Function
    
    Friend Sub OnChildElementsChanged(ByVal Sender As Object, e As CollectionChangedEventArgs)
        BuildNamedItemsList
        SetFocus
    End Sub
    
    Private Sub BuildNamedItemsList()
        Dim Obj As Object
        
        m_NamedChildren.Clear
                        
        For Each Obj In m_Children
            AddNamedItem Obj
        Next
    End Sub
    
    Private Sub AddNamedItem(Obj As Object)
    
    On Error Resume Next
        
    Dim Name As String
    Dim Children As Object
                
    Err.Clear
    
    Name = Obj.Name
    Set Children = Obj.Children
    
    Err.Clear
        
Try:
    On Error GoTo Catch
     
    If Len(Name) Then m_NamedChildren.Add Name, Obj
          
    If Children Is Nothing Then Exit Sub
     
    For Each Obj In Children
        AddNamedItem Obj
    Next
        
    Exit Sub

Catch:
        
        Err.Raise Err.Number, TypeName(Me), Err.Description, Err.HelpFile, Err.HelpContext
    
    End Sub
'#End Region

'#Region Constructor
    Friend Sub Initialize(Superclass As IWindow, Optional ParentWindow As Object = Nothing, Optional CanGetFocus As Boolean = True)
        InitializeObject ParentWindow, CanGetFocus
        
        m_BasePtr = ObjPtr(Superclass)
        Application.Current.Base.Windows.Add m_BasePtr, Superclass
        Superclass.InitializeComponent
    End Sub
'#End Region

'#Region Constructor Helpers
    Private Sub InitializeObject(ParentWindow As Object, CanGetFocus As Boolean)
            Set m_Bindings = New VCF.List
            
            Set m_DependencyProperties = NewDependencyProperties(Me)
            With m_DependencyProperties
                Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
                Call .Register("ShowGridLines", vbBoolean)
            End With
            
            DesignWidth = 1024
            DesignHeight = 768
            
            If ParentWindow Is Nothing Then
                Set Form = Cairo.WidgetForms.Create(, , , DesignWidth, DesignHeight)
            Else
                Set Form = Cairo.WidgetForms.CreateChild(ParentWindow.hWnd, , CanGetFocus)
            End If
            
            ' Enable subclassing & F10 functionality support
            Set m_Subclass = New_c.Subclass
            m_Subclass.Hook Form.hWnd
            
            Set WRoot = Form.WidgetRoot
            WRoot.BackColor = vbWhite
            WRoot.ImageKeyRenderBehaviour = ImgKeyRenderStretched
            
            Set m_Children = NewUIElementCollection(Me)
            Set m_NamedChildren = New ObservableDictionary
            
            NewUIElementBase Me, m_BaseUIElement
    End Sub
'#End Region

