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

Private m_ItemsHost As StackPanel

Private m_DesignLeft As Single
Private m_DesignTop As Single
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

Public Property Get ItemsHost() As StackPanel
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
    m_Layout.LegacyScaleLayout = False

    DesignLeft = 0
    DesignTop = 0
    m_Layout.Width = 300
    m_Layout.Height = 300
    m_Layout.Visibility = VisibilityVisible

    NewUIElementBase Me, m_BaseUIElement

    Set m_ItemsHost = New StackPanel
    m_ItemsHost.Orientation = OrientationVertical
    m_Children.Add m_ItemsHost

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
    RebuildAllItems
End Sub

Private Sub RebuildAllItems()
    Dim i As Long
    Dim Item As Variant
    Dim WasLocked As Boolean
    Dim HostWasLocked As Boolean
    Dim ErrNum As Long
    Dim ErrDesc As String
    Dim ErrSource As String

    If m_ItemsHost Is Nothing Then Exit Sub
    If m_ItemsSource Is Nothing Then Exit Sub
    If m_ItemTemplate Is Nothing Then Exit Sub
    If W Is Nothing Then Exit Sub

    HostWasLocked = m_ItemsHost.Widget.LockRefresh

    WasLocked = W.LockRefresh
    W.LockRefresh = True
    m_ItemsHost.Widget.LockRefresh = True

    On Error GoTo CleanUp

    m_ItemsHost.Children.ClearSilent
    m_ItemVisuals.Clear

    For i = 0 To m_ItemsSource.Count - 1
        Call API.CopyVariable(m_ItemsSource(i), Item)
        AddItemAt i, Item, True
    Next

    ArrangeItemsHost
    Call modLayoutEngine.ArrangeStackPanelChildren( _
        m_ItemsHost.Children, m_ItemsHost.Widget, m_ItemsHost.Orientation, W.Width, W.Height)
    GoTo Done

CleanUp:
    ErrNum = Err.Number
    ErrDesc = Err.Description
    ErrSource = Err.Source

Done:
    W.LockRefresh = WasLocked
    m_ItemsHost.Widget.LockRefresh = HostWasLocked
    If Not W.LockRefresh Then W.Refresh
    If ErrNum <> 0 Then Err.Raise ErrNum, ErrSource, ErrDesc
End Sub

Private Sub AddItemAt(ByVal Index As Long, ByVal Item As Variant, Optional ByVal Silent As Boolean = False)
    Dim Root As Object

    Set Root = modItemTemplateEngine.CloneItemVisualForItem(m_ItemTemplate, Item, m_BaseUIElement)
    If Root Is Nothing Then
        Err.Raise vbObjectError + 10, TypeName(Me), "ItemTemplate produced no visual for index " & Index
    End If

    If Index >= m_ItemVisuals.Count Then
        m_ItemVisuals.Add Root
        If Silent Then
            m_ItemsHost.Children.AddSilent Root
        Else
            m_ItemsHost.Children.Add Root
        End If
    Else
        m_ItemVisuals.Insert Index, Root
        m_ItemsHost.Children.Insert Index, Root
    End If
End Sub

Private Sub RemoveItemAt(ByVal Index As Long)
    If Index < 0 Or Index >= m_ItemVisuals.Count Then Exit Sub
    m_ItemVisuals.RemoveAt Index
    m_ItemsHost.Children.RemoveAt Index
End Sub

Private Sub ReplaceItemAt(ByVal Index As Long, ByVal Item As Variant)
    If Index < 0 Or Index >= m_ItemVisuals.Count Then Exit Sub
    m_ItemsHost.Children.RemoveAt Index
    m_ItemVisuals.RemoveAt Index
    AddItemAt Index, Item
End Sub

Private Sub MoveItemVisual(ByVal OldIndex As Long, ByVal NewIndex As Long)
    Dim Visual As Object
    Set Visual = m_ItemVisuals(OldIndex)
    m_ItemVisuals.RemoveAt OldIndex
    m_ItemVisuals.Insert NewIndex, Visual
    m_ItemsHost.Children.RemoveAt OldIndex
    m_ItemsHost.Children.Insert NewIndex, Visual
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
            Else
                Err.Raise vbObjectError + 4, TypeName(Me), "ItemsSource must be an ObservableCollection"
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
End Sub
'#End Region

'#Region Private Methods
Private Sub ArrangeItemsHost()
    If m_ItemsHost Is Nothing Then Exit Sub
    m_ItemsHost.Move 0, 0, W.Width, W.Height
End Sub

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
'#End Region
