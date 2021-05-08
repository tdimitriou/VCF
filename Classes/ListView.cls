VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IControl
Implements IUIElement
Implements IItemsControl

Private WithEvents W                            As cWidgetBase
Attribute W.VB_VarHelpID = -1

Private m_DesignLeft                            As Single
Private m_DesignTop                             As Single
Private m_DesignWidth                           As Single
Private m_DesignHeight                          As Single
Private m_Visibility                            As Visibility
Private m_ParentPtr                             As Long
Private m_Bindings                              As List
Private m_BaseUIElement                         As UIElementBase
Private m_Handlers                              As cSortedDictionary
Private m_Background                            As Long
Private m_SelectedBackground                    As Long
Private m_SelectedForeground                    As Long

Private WithEvents m_DependencyProperties       As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1

Private WithEvents m_ItemsSource                As ObservableCollection
Attribute m_ItemsSource.VB_VarHelpID = -1
Private WithEvents m_ListCollectionView         As VCF.ListCollectionView
Attribute m_ListCollectionView.VB_VarHelpID = -1
Private WithEvents m_Base                       As ListViewBase
Attribute m_Base.VB_VarHelpID = -1

Private m_Name                                  As String
Private m_ItemTemplate                          As DataTemplate
Private ItemTemplates                           As List

'#Region Public Properties
    Public Property Get Base() As ListViewBase
        Set Base = m_Base
    End Property
    
    Public Property Get Resources() As ObservableDictionary
        Set Resources = m_BaseUIElement.Resources
    End Property
    
    Public Property Get Background() As Long
        Background = m_Background
    End Property
    
    Public Property Let Background(ByVal Value As Long)
        m_Background = Value
    End Property
    
    Public Property Get SelectedBackground() As Long
        SelectedBackground = m_SelectedBackground
    End Property
    
    Public Property Let SelectedBackground(ByVal Value As Long)
        m_SelectedBackground = Value
    End Property
    
    Public Property Get SelectedForeground() As Long
        SelectedForeground = m_SelectedForeground
    End Property
    
    Public Property Let SelectedForeground(ByVal Value As Long)
        m_SelectedForeground = Value
    End Property
    
    Public Property Get RowHeight() As Long
        RowHeight = m_Base.RowHeight
    End Property
    
    Public Property Let RowHeight(ByVal Value As Long)
        m_Base.RowHeight = Value
    End Property
    
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
    End Property
    
    Public Property Get ItemsSource() As Object
        Dim Value
        Call API.CopyVariable(m_DependencyProperties.GetValue("ItemsSource"), Value)
        If Not IsObject(Value) Then Exit Property ' return Nothing
        
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
    End Property
    
    Public Property Get Name() As String
        Name = m_Name
    End Property
    
    Public Property Let Name(ByVal Value As String)
        m_Name = Value
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
        Set m_Handlers = New_c.SortedDictionary(BinaryCompare)
        Set m_Bindings = New List
        Set m_DependencyProperties = NewDependencyProperties(Me)
        Set ItemTemplates = New List
        
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ItemsSource", vbObject)
            Set Me.ItemsSource = Nothing
        End With
        
        Set m_Base = New ListViewBase
        Set W = m_Base.Widget
        W.CanGetFocus = False
        
        DesignLeft = 0
        DesignTop = 0
        DesignWidth = 100
        DesignHeight = 30
        
        m_SelectedBackground = vbBlue
        m_SelectedForeground = vbWhite
        m_Background = vbWhite
        
        NewUIElementBase Me, m_BaseUIElement
        
'        Set m_ItemTemplate = New DataTemplate
'        Dim Tb As TextBlock
'        Set Tb = New TextBlock
'        Tb.DesignWidth = 200
'
'        Tb.Text = "Hello"
'        m_ItemTemplate.Children.Add Tb
'
'        Set Tb = New TextBlock
'        Tb.DesignLeft = 200
'        Tb.ForeColor = vbRed
'        Tb.DesignWidth = 100
'        Tb.HorizontalAlignment = 1
'        Tb.Text = "12.00"
'        m_ItemTemplate.Children.Add Tb
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

'#Region IItemsControl Interface Implementation
    Private Property Set IItemsControl_ItemsSource(ByVal RHS As Object)
        Set ItemsSource = RHS
    End Property
    
    Private Property Get IItemsControl_ItemsSource() As Object
        Set IItemsControl_ItemsSource = ItemsSource
    End Property
    
    Private Property Set IItemsControl_ItemTemplate(ByVal RHS As DataTemplate)
        '
    End Property
    
    Private Property Get IItemsControl_ItemTemplate() As DataTemplate
        '
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

'#Region ListViewBase Events
    Private Sub m_Base_DimensionsAdjusted()
            Dim Child As Object
                    
            If m_ItemTemplate Is Nothing Then Exit Sub
                    
            For Each Child In m_ItemTemplate.Children
                Call MoveChild(Child)
            Next
    End Sub
    
    Private Sub m_Base_ListIndexChanged()
        m_ListCollectionView.MoveCurrentToPosition m_Base.ListIndex
    End Sub
    
    Private Sub m_Base_OwnerDrawItem(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
        On Error Resume Next
        
        Dim ActiveTemplate As DataTemplate
        DrawItemBackground Index, CC, dx, dy, Alpha
    
        Set ActiveTemplate = FindDataTemplate(Index)
        
        If Not ActiveTemplate Is Nothing Then
            ' Draw each VisualChild
            Dim Child As IVisualChild
            Dim Obj As ICloneable
            Dim Element As IUIElement
    
            For Each Child In ActiveTemplate.Children
                Set Element = Child
    
                Call MoveChild(Element)
    
                If IsObject(m_ItemsSource(Index)) Then
                    Child.DrawOn CC, IIf(Index = m_Base.ListIndex, m_SelectedForeground, -1)
                 Else
                    DrawDefaultItemTemplate Index, CC, dx, dy, Alpha
                End If
            
            Next
        Else
            DrawDefaultItemTemplate Index, CC, dx, dy, Alpha
        End If
    End Sub

    Private Sub m_Base_OwnerDrawRowSelector(CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
        '
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        Dim Value
        If Property.Name = "ItemsSource" Then
            Call API.CopyVariable(Property.GetValue, Value)
            If Not IsObject(Value) Then Exit Sub
            
            If Value Is Nothing Then
                SetItemSource Value
            Else
                If Not TypeOf Value Is ObservableCollection Then Exit Sub
                SetItemSource Value
            End If
        End If
    End Sub
'#End Region

'#Region ItemsSource Events

    Private Sub m_ItemsSource_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
        Dim Index As Long
        Dim Value
        
        Select Case e.Action
        
            Case CollectionChangedAction.CollectionChangedActionAdd
                Index = e.NewStartingIndex
                
                For Each Value In e.NewItems
                    ItemTemplates.Insert Index, Empty
                    Index = Index + 1
                Next
            Case CollectionChangedAction.CollectionChangedActionMove
                '
            Case CollectionChangedAction.CollectionChangedActionRemove
                ItemTemplates.RemoveRange e.OldStartingIndex, e.OldItems.Count
            Case CollectionChangedAction.CollectionChangedActionReplace
                If TypeName(e.NewItems(0)) = TypeName(e.OldItems(0)) Then
                    ' No need to change anything
                Else
                    
                    ' Reset the template at the specified index
                    Set ItemTemplates(e.NewStartingIndex) = Empty
                End If
            Case CollectionChangedAction.CollectionChangedActionReset
                Set ItemTemplates = New List
        End Select
    
        m_Base.ListCount = m_ItemsSource.Count
        UpdateHandlers e
    End Sub
'#End Region

'#Region ListCollectionView Events
    Private Sub m_ListCollectionView_CurrentChanged(Sender As Object, ByVal NewPosition As Long, ByVal OldPosition As Long)
        On Error Resume Next
        
        m_Base.ListIndex = m_ListCollectionView.CurrentPosition
    End Sub
'#End Region


'#Region Private Methods
    Private Sub MoveChild(Child As Object)
        On Error Resume Next
        
        Dim xFactor As Double, yFactor As Double
        
        xFactor = W.Width / DesignWidth
        yFactor = W.Height / DesignHeight
        
        
        With Child
            
            .Move (.DesignLeft) * xFactor, (.DesignTop) * yFactor, (.DesignWidth) * xFactor, (.DesignHeight) * yFactor
        
        End With
    End Sub
'#End Region

'#Region Helper Methods
    Private Sub DrawDefaultItemTemplate(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
    ' Draw the default representation of the item
        
        Dim Data
        Dim Text As String
        
        API.CopyVariable m_ItemsSource(Index), Data
            
        ' What is done...
        ' We try to cast the value to string
        ' Even if the value is an object it will be cast successfully if there exists a default property which can be cast
        ' If the cast is not possible we return the the string representation of the object
        On Error Resume Next
        Err.Clear
        Text = Data
        If Err Then Text = "{" & TypeName(Data) & "}"
        
        W.SelectFontSettingsInto CC, IIf(Index = m_Base.ListIndex, vbWhite, -1)
            
        CC.DrawText 0, 0, CDbl(dx), CDbl(dy), Text, , , , 1
    End Sub
    
    Private Sub DrawItemBackground(ByVal Index As Long, CC As vbRichClient5.cCairoContext, ByVal dx As Single, ByVal dy As Single, ByVal Alpha As Single)
        Dim Width As Double, Height As Double
        
        Width = dx
        Height = dy
            
        Dim BgColor As Long
        Dim BgAlpha As Single
        Dim TxtColor As Long
        
        If Index = m_Base.ListIndex Then
            BgColor = m_SelectedBackground
            BgAlpha = Alpha
        Else
            If Index = m_Base.HoverIndex Then
                BgColor = m_SelectedBackground
                BgAlpha = 0.1
            Else
                BgColor = m_Background
                BgAlpha = Alpha
            End If
        End If
        
        CC.Rectangle 0, 0, Width, Height
        CC.SetSourceColor BgColor, BgAlpha
        CC.Fill
    End Sub

    Private Sub SetItemSource(ByVal Value As ObservableCollection)
        On Error Resume Next
        
        Set m_ItemsSource = Value
        
        Set m_ListCollectionView = Nothing
        
        If m_ItemsSource Is Nothing Then Exit Sub
        
        Set m_ListCollectionView = modStaticClasses.CollectionViewSource.GetDefaultView(m_ItemsSource)
        
        Dim v As Variant
        Dim Obj As VCF.INotifyPropertyChanged
             
        Set ItemTemplates = New List
        Set m_Handlers = New_c.SortedDictionary(BinaryCompare)
        
        For Each v In m_ItemsSource
            
            ItemTemplates.Add Empty
            
            If IsObject(v) Then
                If TypeOf v Is VCF.INotifyPropertyChanged Then
                    Set Obj = v
                    AddHandler Obj
                End If
            End If
        Next
            
        m_Base.ListCount = m_ItemsSource.Count
        
        m_Base.ListIndex = m_ListCollectionView.CurrentPosition
    End Sub
    
    Private Sub UpdateHandlers(e As VCF.CollectionChangedEventArgs)
        Dim v As Variant
        Dim Obj As VCF.INotifyPropertyChanged
        
        If Not e.NewItems Is Nothing Then
            For Each v In e.NewItems
                If IsObject(v) Then
                    If TypeOf v Is VCF.INotifyPropertyChanged Then
                        Set Obj = v
                        AddHandler Obj
                    End If
                End If
            Next
        End If
        
        If Not e.OldItems Is Nothing Then
            For Each v In e.OldItems
                If IsObject(v) Then
                    If TypeOf v Is VCF.INotifyPropertyChanged Then
                        Set Obj = v
                        RemoveHandler Obj
                    End If
                End If
            Next
        End If
    End Sub
    
    Private Sub AddHandler(ByVal Value As VCF.INotifyPropertyChanged)
        Dim H As VCF.ListViewPropertyChangedHandler
        
        If m_Handlers.Exists(ObjPtr(Value)) Then Exit Sub
        
        Set H = New VCF.ListViewPropertyChangedHandler
            
        H.Init Value.PropertyChangedEvent, Me
        
        m_Handlers.Add ObjPtr(Value), H
        
    End Sub
    
    Private Sub RemoveHandler(ByVal Value As VCF.INotifyPropertyChanged)
        If m_Handlers.Exists(ObjPtr(Value)) Then m_Handlers.Remove ObjPtr(Value)
    End Sub
    
    Friend Sub PropertyChangedCallback(Sender As Object, Optional e As Variant)
        On Error Resume Next
        
        Dim Index As Long
        Index = m_ItemsSource.IndexOf(Sender)
        
        ' Reset the Template for the Item, so we can recreate it.
        If Index >= 0 Then ItemTemplates(Index) = Empty
        
        m_Base.Widget.Refresh
    End Sub
    
    Private Function FindDataTemplate(ByVal Index As Long) As DataTemplate
        ' Definitions:
        ' 1. Empty: Not initialized yet. Must call CreateDataTemplate and save the returned object
        ' 2. Nothing: Already searched but nothing found. No further action
        ' 3. DataTemplateObject: Already searched, found and saved to the collection. No further action
        
        On Error Resume Next
        
        Err.Clear
        
        Set FindDataTemplate = ItemTemplates(Index)
        
        ' The Template has already been Created, so the call to the
        ' ItemTemplates(Index) has returned Object ( or Nothing).
        ' We are good to go!!!
        If Err.Number = 0 Then Exit Function
        
        Err.Clear
        
        ' The Template is still empty.
        ' We must create, save for future use and return the result.
        Set FindDataTemplate = CreateDataTemplate(m_ItemsSource(Index))
        
        If Err Then
            Err.Clear
            Exit Function
        End If
        ' Do not save in the case that an error has occured,
        ' in order to be able to create for it again.
        
        Set ItemTemplates(Index) = FindDataTemplate
    
    End Function
    
    Private Function CreateDataTemplate(Item) As DataTemplate
        Dim ActiveTemplate As DataTemplate
        
        Dim Key
        Key = "DataTemplate_" & TypeName(Item)
        
        Dim ResourceTemplate
        API.CopyVariable m_BaseUIElement.TryFindResource(Key), ResourceTemplate
        
        If Not IsEmpty(ResourceTemplate) Then Set ActiveTemplate = ResourceTemplate
        
        If ActiveTemplate Is Nothing Then Set ActiveTemplate = m_ItemTemplate
            
        If ActiveTemplate Is Nothing Then Exit Function
            
        ' CloneActiveTemplate
        Dim NewTemplate As DataTemplate
        Set NewTemplate = New DataTemplate
        
        Dim Child As IVisualChild
        Dim Obj As ICloneable
        Dim Element As IUIElement
                        
        For Each Obj In ActiveTemplate.Children
            Set Obj = Obj.Clone
            Set Child = Obj
            Set Element = Obj
            
            If IsObject(Item) Then Set Element.DataContext = Item
            NewTemplate.Children.Add Obj
        Next
        
        Set CreateDataTemplate = NewTemplate
    End Function
    
'#End Region

