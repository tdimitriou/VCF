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
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private WithEvents m_Children As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1

Private m_NamedChildren As ObservableDictionary
Private m_Bindings As List
Private m_DialogResult As Boolean
Private m_BasePtr As Long
Private m_BaseUIElement As UIElementBase
Private m_Layout As FrameworkElement

Private m_W As Long, m_H As Long
Private m_DeferredLayout As Boolean

Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1

Private WithEvents m_Subclass As cSubClass
Attribute m_Subclass.VB_VarHelpID = -1

Public Event WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, _
                        ByVal Subclass As cSubClass, ByRef CancelDefaultProc As Boolean)

'#Region Form Events
    Private Sub Form_ResizeWithDimensions(ByVal NewWidth As Long, ByVal NewHeight As Long)
        Dim Child As Object
                
        If m_Children Is Nothing Then Exit Sub
        
        m_Layout.OnHostResize m_Children, WRoot, Form.ScaleWidth, Form.ScaleHeight
    End Sub
    
    Private Sub Form_Unload(Cancel As Integer)
        ' Single teardown path for Window. Callers use Window.Unload -> Form.Unload.
        ' Order: detach bindings while sinks/tree are valid, then unhook, then registry.
        On Error Resume Next

        If Not WRoot Is Nothing Then WRoot.LockRefresh = True

        modBindingExpressions.DetachBindingsTree Me

        If Not m_Subclass Is Nothing Then
            m_Subclass.UnHook
            Set m_Subclass = Nothing
        End If

        If m_BasePtr <> 0 Then
            If Not Application.Current Is Nothing Then
                If Not Application.Current.Base Is Nothing Then
                    If Application.Current.Base.Windows.Exists(m_BasePtr) Then
                        Application.Current.Base.Windows.Remove m_BasePtr
                    End If
                End If
            End If
            m_BasePtr = 0
        End If
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
    
    Public Property Get Style() As Style
        Call API.CopyVariable(m_DependencyProperties.GetValue("Style"), Style)
    End Property

    Public Property Set Style(ByVal Value As Style)
        Call m_DependencyProperties.SetValue("Style", Value)
    End Property
    
    Public Property Get NamedChildren() As ObservableDictionary
        Set NamedChildren = m_NamedChildren
    End Property
'#End Region

'#Region Public Methods
    Public Sub Show(Optional Modal As FormShowConstants = vbModeless, Optional Owner As Object, Optional Focused As Boolean = True)
        SyncFormBorderStyle
        Form.Show Modal, Owner, Not Focused
    End Sub
    
    Public Function ShowDialog(Optional Owner As Object) As Boolean
        SyncFormBorderStyle
        Form.Show vbModal, Owner
        
        ShowDialog = m_DialogResult
    End Function
    
    ' Preferred shutdown vs compiled DLL:
    ' 1) Detach bindings while tree+Form still live
    ' 2) Release children NOW (Form still alive) � post-unload Release AVs;
    '    hold-bag only deferred that hang to IDE End after MsgBox
    ' 3) Form.Unload -> Form_Unload (unhook, Windows.Remove)
    ' 4) Drop Form/WRoot so Set Shell = Nothing is cheap
    ' Prepare sinks -> transfer children to hold -> Form.Unload -> drain hold.
    ' (Direct Release before/after Form.Unload freezes vs compiled DLL on P7c Buttons.)
    Public Sub Unload()
        On Error Resume Next

        modWindowTeardown.LogStep "Window.Unload: enter"
        Debug.Print "Window.Unload: enter"

        If Not WRoot Is Nothing Then WRoot.LockRefresh = True

        modWindowTeardown.LogStep "Window.Unload: detach"
        modBindingExpressions.DetachBindingsTree Me

        modWindowTeardown.LogStep "Window.Unload: prepare sinks"
        modWindowTeardown.PrepareWindowChildren Me

        modWindowTeardown.LogStep "Window.Unload: transfer children"
        If Not m_Children Is Nothing Then modWindowTeardown.TransferAll m_Children
        m_DeferredLayout = False

        ' Disarm+Release while Form/Cairo still alive. After Form.Unload this hangs
        ' (compiled DLL); deferring to IDE End only moves the same hang.
        modWindowTeardown.LogStep "Window.Unload: DrainHoldForce (pre-Form)"
        modWindowTeardown.DrainHoldForce
        modWindowTeardown.LogStep "Window.Unload: DrainHoldPermanent (pre-Form)"
        modWindowTeardown.DrainHoldPermanent

        modWindowTeardown.LogStep "Window.Unload: Form.Unload"
        If Not Form Is Nothing Then Form.Unload

        modWindowTeardown.LogStep "Window.Unload: drop Form"
        Set WRoot = Nothing
        Set Form = Nothing

        modWindowTeardown.LogStep "Window.Unload: done"
        Debug.Print "Window.Unload: done"
    End Sub
'#End Region

'#Region UIElementCollection Events
    Private Sub m_Children_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
        On Error GoTo Catch
            
        Dim Child As Object
        
Try:
               
        If Not e.OldItems Is Nothing Then
            For Each Child In e.OldItems
                If TypeOf Child Is IControl Then If WRoot.Widgets.Exists(Child.Widget.Key) Then WRoot.Widgets.Remove Child.Widget.Key
            Next
        End If
                
        If Not e.NewItems Is Nothing Then
            If WRoot.LockRefresh Then
                m_DeferredLayout = True
            Else
                For Each Child In e.NewItems
                    MoveChild Child
                Next
            End If
        End If
        
        If Not WRoot.LockRefresh Then WRoot.Refresh
        
        Exit Sub
        
Catch:
        '
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        On Error Resume Next
        
        Select Case Property.Name
            Case "Style":                ApplyStyle
            Case "BorderStyle":
                
                If Not Form Is Nothing Then
                    Dim Desired As Long
                    Desired = CLng(Property.GetValue)
                    If Form.IsChildWindow Then
                        Property.SetCurrentValue Form.BorderStyle
                    ElseIf Form.BorderStyle <> Desired Then
                        Form.BorderStyle = Desired
                    End If
                End If
            Case Else
                If Not m_Children Is Nothing Then
                    If m_Layout.HandlesLayoutProperty(Property.Name) Then
                        m_Layout.OnLayoutPropertyChanged Property.Name, m_Children, WRoot, Form.ScaleWidth, Form.ScaleHeight
                    End If
                End If
        End Select
        
        If Not m_Children Is Nothing Then DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Sub
'#End Region

'#Region Subclass Events
    Private Sub m_Subclass_WindowProc(Result As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long)
        
        'Enable F10 Key Functionality
        If Msg = 260 And wParam = 121 Then Result = m_Subclass.CallWindowProc(256, wParam, lParam)

        ' Pass the event to consumers to allow subclassing
        Dim Cancel As Boolean
        RaiseEvent WindowProc(Result, Msg, wParam, lParam, m_Subclass, Cancel)
        If Cancel = False Then Result = m_Subclass.CallWindowProc(Msg, wParam, lParam)
    
    End Sub
'#End Region

'#Region Private Methods
    Private Sub MoveChild(Child As Object)
        m_Layout.ArrangeChildren m_Children, WRoot, Form.ScaleWidth, Form.ScaleHeight
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
        If WRoot.LockRefresh Then Exit Sub
        BuildNamedItemsList
    End Sub
    
    ' Phase 2a.1 ? Visibility swap on a direct Window child (view host) should not require
    ' the app to call RelayoutChildren / RebuildNamedItemsList manually.
    Friend Sub OnChildVisibilityChanged()
        On Error Resume Next
        If Form Is Nothing Then Exit Sub
        If Not WRoot Is Nothing Then
            If WRoot.LockRefresh Then Exit Sub
        End If
        RelayoutChildren
        RebuildNamedItemsList
    End Sub
    
    Public Sub ApplyDeferredChildLayout()
        If Not m_DeferredLayout Then Exit Sub
        m_DeferredLayout = False
        RelayoutChildren
    End Sub
    
    Public Sub RelayoutChildren()
        On Error Resume Next
        m_DeferredLayout = False
        If m_Children Is Nothing Then Exit Sub
        If Form Is Nothing Then Exit Sub
        
        m_Layout.OnHostResize m_Children, WRoot, Form.ScaleWidth, Form.ScaleHeight
        
        Dim Child As Object
        For Each Child In m_Children
            ApplyChildHostLayout Child
        Next
    End Sub
    
    Private Sub ApplyChildHostLayout(ByVal Child As Object)
        On Error Resume Next
        If TypeOf Child Is IUserControl Then
            Dim IU As IUserControl
            Set IU = Child
            IU.Base.ApplyDeferredHostLayout
        ElseIf TypeOf Child Is UserControl Then
            Child.ApplyDeferredHostLayout
        End If
    End Sub
    
    Public Sub RebuildNamedItemsList()
        BuildNamedItemsList
    End Sub
    
    Private Sub BuildNamedItemsList()
        Dim Obj As Object
        
        m_NamedChildren.Clear
                        
        For Each Obj In m_Children
            On Error Resume Next
            If TypeOf Obj Is IUserControl Then
                If Obj.Visibility = VisibilityCollapsed Then GoTo NextChild
            End If
            AddNamedItem Obj
NextChild:
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
     
    If Len(Name) Then
        If Not m_NamedChildren.ContainsKey(Name) Then m_NamedChildren.Add Name, Obj
    End If
          
    If Children Is Nothing Then Exit Sub
     
    For Each Obj In Children
        AddNamedItem Obj
    Next
        
    Exit Sub

Catch:
        
        Err.Clear
        If Children Is Nothing Then Exit Sub
        For Each Obj In Children
            AddNamedItem Obj
        Next
    End Sub
'#End Region

'#Region Constructor
    Friend Sub Initialize(Superclass As IWindow, Optional ParentWindow As Object = Nothing, Optional CanGetFocus As Boolean = True)
        InitializeObject ParentWindow, CanGetFocus
        
        m_BasePtr = ObjPtr(Superclass)
        Application.Current.Base.Windows.Add m_BasePtr, Superclass
        Superclass.InitializeComponent
        
        Set Style = GetBaseStyle
        
        SyncFormBorderStyle
    End Sub
'#End Region

'#Region Constructor Helpers
    Private Sub InitializeObject(ParentWindow As Object, CanGetFocus As Boolean)
            Set m_Bindings = New VCF.List
            
            Set m_DependencyProperties = NewDependencyProperties(Me)
            With m_DependencyProperties
                Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
                Call .Register("ShowGridLines", vbBoolean)
                Call .Register("BorderStyle", vbLong, , , , NewDependencyPropertyMetadata(False, True, False, OneWay, vbSizable))
                Call .Register("Style", vbObject, "Style", , , NewDependencyPropertyMetadata(False, True, False))
                .SetValue "Style", Nothing
                ' BorderStyle default via metadata DefaultValue (vbSizable), not SetCurrentValue.
            End With
            
            Dim DefaultW As Double
            Dim DefaultH As Double
            DefaultW = 1024
            DefaultH = 768
            
            If ParentWindow Is Nothing Then
                Set Form = Cairo.WidgetForms.Create(0, , , DefaultW, DefaultH)
                On Error Resume Next
                Form.BorderStyle = 0
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
            
            Set m_Layout = New FrameworkElement
            m_Layout.Initialize Me, Me, WRoot
            m_Layout.RegisterLayoutProperties "Window"
            m_Layout.Width = DefaultW
            m_Layout.Height = DefaultH
            
            Form.Width = DefaultW
            Form.Height = DefaultH
            
            NewUIElementBase Me, m_BaseUIElement
            
            'EnumerateProperties Me
    End Sub
'#End Region

Private Sub SyncFormBorderStyle()
    On Error Resume Next
    If Form Is Nothing Then Exit Sub
    If Form.IsChildWindow Then Exit Sub
    
    Dim Desired As Long
    Desired = CLng(m_DependencyProperties.GetValue("BorderStyle"))
    If Form.BorderStyle <> Desired Then Form.BorderStyle = Desired
End Sub

Private Function GetBaseStyle() As Style
    Dim Style
    API.CopyVariable m_BaseUIElement.TryFindResource(TypeName(Me) & ".BaseStyle"), Style
    
    If Not IsObject(Style) Then Exit Function
    If TypeOf Style Is Style Then Set GetBaseStyle = Style
End Function

Private Sub ApplyStyle()
    On Error Resume Next
    If Not ActiveStyle Is Style Then Set ActiveStyle = Style
    If Style Is Nothing Then Exit Sub
    With New StyleManager
        .ApplyStyle Style, Me
    End With
End Sub


    Private Sub ActiveStyle_StyleChanged()
        On Error Resume Next
        ApplyStyle
    End Sub
