VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ItemsControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement
Implements IItemsControl

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents m_Children As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private WithEvents m_ItemsSource As ObservableCollection
Attribute m_ItemsSource.VB_VarHelpID = -1

Private m_ItemsHost As Object
Private m_ItemsPanel As ItemsPanelTemplate

Private m_ParentPtr As Long
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase
Private m_Layout As FrameworkElement
Private m_ItemTemplate As DataTemplate
Private m_ItemVisuals As List

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

Public Property Get ItemsSource() As Object
    Dim Value As Variant
    Call API.CopyVariable(m_DependencyProperties.GetValue("ItemsSource"), Value)
    If Not IsObject(Value) Then Exit Property
    Set ItemsSource = Value
End Property

Public Property Set ItemsSource(ByVal Value As Object)
    Call m_DependencyProperties.SetValue("ItemsSource", Value)
End Property

Public Property Get ItemTemplate() As DataTemplate
    Set ItemTemplate = m_ItemTemplate
End Property

Public Property Set ItemTemplate(ByVal Value As DataTemplate)
    Set m_ItemTemplate = Value
    If Not m_ItemsSource Is Nothing Then RebuildAllItems
End Property

' Panel that lays out generated item visuals. Default: vertical StackPanel.
Public Property Get ItemsPanel() As ItemsPanelTemplate
    Set ItemsPanel = m_ItemsPanel
End Property

Public Property Set ItemsPanel(ByVal Value As ItemsPanelTemplate)
    Set m_ItemsPanel = Value
    InstallItemsHost CreateItemsHostInstance()
End Property

' Live panel instance (StackPanel by default, or clone of ItemsPanel root).
Public Property Get ItemsHost() As Object
    Set ItemsHost = m_ItemsHost
End Property

Public Property Get ItemCount() As Long
    If m_ItemsSource Is Nothing Then Exit Property
    ItemCount = m_ItemsSource.Count
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
End Property

Public Property Get DataContext() As Object
    Call API.CopyVariable(m_DependencyProperties.GetValue("DataContext"), DataContext)
End Property

Public Property Let Visibility(ByVal Value As Visibility)
    m_Layout.Visibility = Value
End Property

Public Property Get Visibility() As Visibility
    Visibility = m_Layout.Visibility
End Property

Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
    W.Move Left, Top, Width, Height
    ApplyVisibility W, Visibility
    ArrangeItemsHost
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
    Set m_ItemVisuals = New List
    Set m_DependencyProperties = NewDependencyProperties(Me)

    With m_DependencyProperties
        Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
        Call .Register("ItemsSource", vbObject)
        Call .Register("ItemTemplate", vbObject, "DataTemplate")
        Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
        .SetValue "Style", Nothing
    End With

    Set W = Cairo.WidgetBase
    W.CanGetFocus = False

    Set m_Children = NewUIElementCollection(Me)

    Set m_Layout = New FrameworkElement
    m_Layout.Initialize Me, Me, W
    m_Layout.RegisterLayoutProperties "ItemsControl"

    m_Layout.Width = 0
    m_Layout.Height = 0
    m_Layout.Visibility = VisibilityVisible

    NewUIElementBase Me, m_BaseUIElement

    InstallItemsHost CreateItemsHostInstance()

    Set Style = GetBaseStyle
    Set m_ItemsSource = Nothing
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

'#Region IItemsControl Interface Implementation
Private Property Set IItemsControl_ItemsSource(ByVal RHS As Object)
    Set ItemsSource = RHS
End Property

Private Property Get IItemsControl_ItemsSource() As Object
    Set IItemsControl_ItemsSource = ItemsSource
End Property

Private Property Set IItemsControl_ItemTemplate(ByVal RHS As DataTemplate)
    Set ItemTemplate = RHS
End Property

Private Property Get IItemsControl_ItemTemplate() As DataTemplate
    Set IItemsControl_ItemTemplate = ItemTemplate
End Property
'#End Region

'#Region ItemsSource / template engine
Private Sub SetItemsSource(ByVal Value As ObservableCollection)
    Set m_ItemsSource = Nothing
    Set m_ItemsSource = Value
    If Value Is Nothing Then
        ClearGeneratedItems
    Else
        RebuildAllItems
    End If
End Sub

' Drop generated visuals when ItemsSource is cleared (RebuildAllItems no-ops if source is Nothing).
Private Sub ClearGeneratedItems()
    Dim HostChildren As UIElementCollection
    Dim WasLocked As Boolean
    Dim HostWasLocked As Boolean

    m_ItemVisuals.Clear
    If m_ItemsHost Is Nothing Then Exit Sub
    If W Is Nothing Then Exit Sub

    HostWasLocked = m_ItemsHost.Widget.LockRefresh
    WasLocked = W.LockRefresh
    W.LockRefresh = True
    m_ItemsHost.Widget.LockRefresh = True

    On Error Resume Next
    Set HostChildren = HostChildrenCollection
    ' Non-silent Clear so the host panel detaches child widgets (ClearSilent leaves them).
    HostChildren.Clear
    On Error GoTo 0

    W.LockRefresh = WasLocked
    If Not m_ItemsHost Is Nothing Then m_ItemsHost.Widget.LockRefresh = HostWasLocked
    If Not W.LockRefresh Then W.Refresh
End Sub

Private Sub RebuildAllItems()
    Dim i As Long
    Dim Item As Variant
    Dim WasLocked As Boolean
    Dim HostWasLocked As Boolean
    Dim ErrNum As Long
    Dim ErrDesc As String
    Dim ErrSource As String
    Dim HostChildren As UIElementCollection

    If m_ItemsHost Is Nothing Then Exit Sub
    If m_ItemsSource Is Nothing Then Exit Sub
    If m_ItemTemplate Is Nothing Then Exit Sub
    If W Is Nothing Then Exit Sub

    HostWasLocked = m_ItemsHost.Widget.LockRefresh

    WasLocked = W.LockRefresh
    W.LockRefresh = True
    m_ItemsHost.Widget.LockRefresh = True

    On Error GoTo CleanUp

    ' Early-bind Children ? Friend ClearSilent/AddSilent are not visible via Object late-bind (438).
    Set HostChildren = HostChildrenCollection
    HostChildren.ClearSilent
    m_ItemVisuals.Clear

    For i = 0 To m_ItemsSource.Count - 1
        Call API.CopyVariable(m_ItemsSource(i), Item)
        AddItemAt i, Item, True
    Next

    ArrangeItemsHost
    ArrangeGeneratedItems
    GoTo Done

CleanUp:
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source

Done:
    W.LockRefresh = WasLocked
    If Not m_ItemsHost Is Nothing Then m_ItemsHost.Widget.LockRefresh = HostWasLocked
    If Not W.LockRefresh Then W.Refresh
    If ErrNum <> 0 Then Err.Raise ErrNum, ErrSource, ErrDesc
End Sub

Private Sub AddItemAt(ByVal Index As Long, ByVal Item As Variant, Optional ByVal Silent As Boolean = False)
    Dim Root As Object
    Dim HostChildren As UIElementCollection
    Dim UseSilent As Boolean

    Set Root = modItemTemplateEngine.CloneItemVisualForItem(m_ItemTemplate, Item, m_BaseUIElement)
    If Root Is Nothing Then
        Err.Raise vbObjectError + 10, TypeName(Me), "ItemTemplate produced no visual for index " & Index
    End If

    Set HostChildren = HostChildrenCollection

    ' Silent batch under host LockRefresh; UniformGrid ArrangeChildren runs once after unlock.
    UseSilent = Silent

    If Index >= m_ItemVisuals.Count Then
        m_ItemVisuals.Add Root
        If UseSilent Then
            HostChildren.AddSilent Root
        Else
            HostChildren.Add Root
        End If
    Else
        m_ItemVisuals.Insert Index, Root
        HostChildren.Insert Index, Root
    End If
End Sub

Private Sub RemoveItemAt(ByVal Index As Long)
    If Index < 0 Or Index >= m_ItemVisuals.Count Then Exit Sub
    m_ItemVisuals.RemoveAt Index
    HostChildrenCollection.RemoveAt Index
End Sub

Private Sub ReplaceItemAt(ByVal Index As Long, ByVal Item As Variant)
    If Index < 0 Or Index >= m_ItemVisuals.Count Then Exit Sub
    HostChildrenCollection.RemoveAt Index
    m_ItemVisuals.RemoveAt Index
    AddItemAt Index, Item
End Sub

Private Sub MoveItemVisual(ByVal OldIndex As Long, ByVal NewIndex As Long)
    Dim Visual As Object
    Dim HostChildren As UIElementCollection

    Set Visual = m_ItemVisuals(OldIndex)
    m_ItemVisuals.RemoveAt OldIndex
    m_ItemVisuals.Insert NewIndex, Visual
    Set HostChildren = HostChildrenCollection
    HostChildren.RemoveAt OldIndex
    HostChildren.Insert NewIndex, Visual
End Sub
'#End Region

'#Region DependencyProperties Events
Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
    Dim Value As Variant

    Select Case Property.Name
        Case "ItemsSource"
            Call API.CopyVariable(Property.GetValue, Value)
            If Value Is Nothing Then
                SetItemsSource Nothing
            ElseIf TypeOf Value Is ObservableCollection Then
                SetItemsSource Value
            End If
        Case "ItemTemplate"
            Call API.CopyVariable(Property.GetValue, Value)
            If IsObject(Value) Then
                Set m_ItemTemplate = Value
            Else
                Set m_ItemTemplate = Nothing
            End If
            If Not m_ItemsSource Is Nothing Then RebuildAllItems
        Case "Style"
            ApplyStyle
        Case Else
            If m_Layout.HandlesLayoutProperty(Property.Name) Then ArrangeItemsHost
            DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Select
End Sub
'#End Region

'#Region ItemsSource Events
Private Sub m_ItemsSource_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
    Dim Index As Long
    Dim Value As Variant

    Select Case e.Action
        Case CollectionChangedActionAdd
            If e.NewItems Is Nothing Then Exit Sub
            Index = e.NewStartingIndex
            For Each Value In e.NewItems
                AddItemAt Index, Value
                Index = Index + 1
            Next
        Case CollectionChangedActionRemove
            If e.OldItems Is Nothing Then Exit Sub
            Index = e.OldStartingIndex + e.OldItems.Count - 1
            Do While Index >= e.OldStartingIndex
                RemoveItemAt Index
                Index = Index - 1
            Loop
        Case CollectionChangedActionReplace
            If e.NewItems Is Nothing Then Exit Sub
            ReplaceItemAt e.NewStartingIndex, e.NewItems(0)
        Case CollectionChangedActionMove
            MoveItemVisual e.OldStartingIndex, e.NewStartingIndex
        Case CollectionChangedActionReset
            RebuildAllItems
    End Select

    ArrangeItemsHost
    If Not W.LockRefresh Then W.Refresh
End Sub
'#End Region

'#Region UIElementCollection Events
Private Sub m_Children_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
    Dim Child As Object

    On Error Resume Next
    If W Is Nothing Then Exit Sub

    If Not e.OldItems Is Nothing Then
        For Each Child In e.OldItems
            If TypeOf Child Is IControl Then
                If W.Widgets.Exists(modLayoutEngine.ControlWidgetKey(Child)) Then
                    W.Widgets.Remove modLayoutEngine.ControlWidgetKey(Child)
                End If
            End If
        Next
    End If

    If Not e.NewItems Is Nothing Then
        For Each Child In e.NewItems
            If TypeOf Child Is IControl Then
                modLayoutEngine.AttachChildWidget Child, W, VisibilityVisible
            End If
        Next
    End If

    ArrangeItemsHost
    If Not W.LockRefresh Then W.Refresh
End Sub
'#End Region

'#Region Widget Events
Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    CC.Save
    CC.Restore
End Sub

Private Sub W_Resize()
    ArrangeItemsHost
    ArrangeGeneratedItems
End Sub
'#End Region

'#Region Private Methods
Private Function CreateItemsHostInstance() As Object
    Dim Root As Object

    If m_ItemsPanel Is Nothing Then
        Set CreateItemsHostInstance = CreateDefaultStackHost()
        Exit Function
    End If
    If m_ItemsPanel.Children Is Nothing Then
        Set CreateItemsHostInstance = CreateDefaultStackHost()
        Exit Function
    End If
    If m_ItemsPanel.Children.Count = 0 Then
        Set CreateItemsHostInstance = CreateDefaultStackHost()
        Exit Function
    End If

    Set Root = m_ItemsPanel.Children(0)
    Set CreateItemsHostInstance = modItemTemplateEngine.CloneItemsPanelRoot(Root)
    If CreateItemsHostInstance Is Nothing Then
        Err.Raise vbObjectError + 11, TypeName(Me), _
            "ItemsPanel root must be StackPanel or UniformGrid (got " & TypeName(Root) & ")"
    End If
End Function

Private Function CreateDefaultStackHost() As StackPanel
    Dim Sp As StackPanel
    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Set CreateDefaultStackHost = Sp
End Function

Private Sub InstallItemsHost(ByVal NewHost As Object)
    Dim HadItems As Boolean

    If NewHost Is Nothing Then Exit Sub

    HadItems = (Not m_ItemsSource Is Nothing) And (Not m_ItemTemplate Is Nothing)

    If Not m_ItemsHost Is Nothing Then
        If m_Children.Contains(m_ItemsHost) Then m_Children.Remove m_ItemsHost
    End If

    Set m_ItemsHost = NewHost
    m_Children.Add m_ItemsHost
    ArrangeItemsHost

    If HadItems Then RebuildAllItems
End Sub

Private Sub ArrangeItemsHost()
    If m_ItemsHost Is Nothing Then Exit Sub
    ' Use Widget.Move ? UniformGrid/StackPanel.Move no-op when Parent is Nothing.
    If W.Width <= 0! Or W.Height <= 0! Then Exit Sub
    m_ItemsHost.Widget.Move 0, 0, W.Width, W.Height
End Sub

' StackPanel / UniformGrid: one arrange after silent item batch.
Private Sub ArrangeGeneratedItems()
    Dim Sp As StackPanel
    Dim Ug As UniformGrid

    If m_ItemsHost Is Nothing Then Exit Sub
    If W.Width <= 0! Or W.Height <= 0! Then Exit Sub

    If TypeOf m_ItemsHost Is StackPanel Then
        Set Sp = m_ItemsHost
        Call modLayoutEngine.ArrangeStackPanelChildren( _
            Sp.Children, Sp.Widget, Sp.Orientation, W.Width, W.Height)
    ElseIf TypeOf m_ItemsHost Is UniformGrid Then
        Set Ug = m_ItemsHost
        Ug.ArrangeChildren
    End If
End Sub

' Resolve Children through a typed variable so Friend AddSilent/ClearSilent work (not via Object).
Private Function HostChildrenCollection() As UIElementCollection
    Dim HostChildren As UIElementCollection
    Set HostChildren = m_ItemsHost.Children
    Set HostChildrenCollection = HostChildren
End Function

Private Function GetBaseStyle() As Style
    Dim StyleObj As Variant
    Call API.CopyVariable(m_BaseUIElement.TryFindResource(TypeName(Me) & ".BaseStyle"), StyleObj)
    If IsObject(StyleObj) Then
        If TypeOf StyleObj Is Style Then Set GetBaseStyle = StyleObj
    End If
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

' Drop WithEvents sinks while Cairo widget is still valid (compiled-DLL teardown).
' Do not Set m_Children = Nothing here — that Releases the host tree early and hangs.
Friend Sub PrepareForUnload()
    On Error Resume Next
    Set m_ItemsSource = Nothing
    Set ActiveStyle = Nothing
    Set W = Nothing
End Sub

' Never Release item visuals or ItemsHost after Form.Unload — that hangs.
' Park them in HoldPermanent; only the ItemsControl shell is Released by DrainHoldForce.
Friend Sub DisarmForRelease()
    Dim Kids As UIElementCollection
    Dim Ch As Object
    Dim Ug As UniformGrid
    Dim Sp As StackPanel

    On Error Resume Next

    modWindowTeardown.LogStep "Disarm enter"
    Set m_ItemsSource = Nothing
    Set ActiveStyle = Nothing
    Set W = Nothing
    modWindowTeardown.LogStep "Disarm sinks cleared"

    If m_ItemsHost Is Nothing Then GoTo ClearSelf

    If TypeOf m_ItemsHost Is UniformGrid Then
        Set Ug = m_ItemsHost
        Set Kids = Ug.Children
        Do While Not Kids Is Nothing
            If Kids.Count <= 0 Then Exit Do
            Set Ch = Kids.DetachAt(0)
            modWindowTeardown.LogStep "Disarm item park " & TypeName(Ch)
            ' No PrepareTree here — sinks were cleared before Form.Unload; doing it after hangs.
            modWindowTeardown.HoldPermanent Ch
            Set Ch = Nothing
            modWindowTeardown.LogStep "Disarm item parked"
        Loop
        modWindowTeardown.LogStep "Disarm items done"
        ' Do not Ug.PrepareForUnload here — hangs after item park (VCF_Unload.log).
    ElseIf TypeOf m_ItemsHost Is StackPanel Then
        Set Sp = m_ItemsHost
        Set Kids = Sp.Children
        Do While Not Kids Is Nothing
            If Kids.Count <= 0 Then Exit Do
            Set Ch = Kids.DetachAt(0)
            modWindowTeardown.LogStep "Disarm item park " & TypeName(Ch)
            modWindowTeardown.HoldPermanent Ch
            Set Ch = Nothing
            modWindowTeardown.LogStep "Disarm item parked"
        Loop
        modWindowTeardown.LogStep "Disarm items done"
    End If

ClearSelf:
    modWindowTeardown.LogStep "Disarm ClearSelf ItemVisuals"
    Set m_ItemVisuals = Nothing
    Set m_ItemTemplate = Nothing
    modWindowTeardown.LogStep "Disarm ClearSelf ItemVisuals done"

    If Not m_Children Is Nothing Then
        Do While m_Children.Count > 0
            Set Ch = m_Children.DetachAt(0)
            modWindowTeardown.LogStep "Disarm host park " & TypeName(Ch)
            ' Park host as-is; PrepareForUnload on emptied UniformGrid hangs.
            modWindowTeardown.HoldPermanent Ch
            Set Ch = Nothing
            modWindowTeardown.LogStep "Disarm host parked"
        Loop
    End If

    modWindowTeardown.LogStep "Disarm ClearSelf refs"
    Set m_ItemsHost = Nothing
    Set m_Children = Nothing
    Set m_Layout = Nothing
    Set m_Bindings = Nothing
    modWindowTeardown.LogStep "Disarm ClearSelf done"
End Sub
'#End Region
