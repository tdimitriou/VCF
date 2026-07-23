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
Implements IContentControl

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents m_Children As UIElementCollection
Attribute m_Children.VB_VarHelpID = -1
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private m_ParentPtr As Long
Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase
Private m_Layout As FrameworkElement
Private m_ContentHost As ContentHost
Private m_DesiredWidth As Double
Private m_DesiredHeight As Double
Private m_ActualWidth As Double
Private m_ActualHeight As Double

Public Name As String

' WPF-aligned Content (Variant): string -> paint via ContentPresenter; IUIElement -> single child.
Public Property Get Content() As Variant
    Dim V As Variant

    Call API.CopyVariable(m_DependencyProperties.GetValue("Content"), V)
    If IsEmpty(V) Or IsNull(V) Then
        If Not m_Children Is Nothing Then
            If m_Children.Count > 0 Then
                Set Content = m_Children(0)
                Exit Property
            End If
        End If
    End If
    Call API.CopyVariable(V, Content)
End Property

Public Property Let Content(ByVal Value As Variant)
    Call m_DependencyProperties.SetValue("Content", Value)
End Property

Public Property Set Content(ByVal Value As Object)
    Call m_DependencyProperties.SetValue("Content", Value)
End Property

Public Property Get ContentTemplate() As DataTemplate
    Call API.CopyVariable(m_DependencyProperties.GetValue("ContentTemplate"), ContentTemplate)
End Property

Public Property Set ContentTemplate(ByVal Value As DataTemplate)
    Call m_DependencyProperties.SetValue("ContentTemplate", Value)
End Property

' Host ContentPresenter (shared ContentHost with Button).
Public Property Get ContentPresenter() As ContentPresenter
    Set ContentPresenter = m_ContentHost.HostPresenter
End Property

Public Sub SyncContentPresenter()
    m_ContentHost.SyncContent m_DependencyProperties
    m_ContentHost.SyncSuppress (m_Children.Count > 0)
End Sub

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

'#Region IContentControl Interface Implementation
    Private Property Get IContentControl_Content() As Variant
        Call API.CopyVariable(Content, IContentControl_Content)
    End Property

    Private Property Let IContentControl_Content(ByVal Value As Variant)
        Content = Value
    End Property

    Private Property Set IContentControl_Content(ByVal Value As Object)
        Set Content = Value
    End Property

    Private Property Get IContentControl_ContentPresenter() As ContentPresenter
        Set IContentControl_ContentPresenter = ContentPresenter
    End Property

    Private Sub IContentControl_SyncContentPresenter()
        SyncContentPresenter
    End Sub
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

    Public Property Get DesiredWidth() As Double
        DesiredWidth = m_DesiredWidth
    End Property

    Public Property Get DesiredHeight() As Double
        DesiredHeight = m_DesiredHeight
    End Property

    Public Property Get ActualWidth() As Double
        ActualWidth = m_ActualWidth
    End Property

    Public Property Get ActualHeight() As Double
        ActualHeight = m_ActualHeight
    End Property

    ' IUIElement Content child: Desired = child Width/Height (content-driven when host unset).
    Public Sub MeasureOverride(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
        Dim ContentW As Double
        Dim ContentH As Double
        Dim DesiredW As Double
        Dim DesiredH As Double
        Dim Child As Object

        If m_Layout Is Nothing Then Exit Sub

        If Not m_Children Is Nothing Then
            If m_Children.Count >= 1 Then
                Set Child = m_Children(0)
                ContentW = ReadChildLayoutWidth(Child)
                ContentH = ReadChildLayoutHeight(Child)

                DesiredW = Width
                DesiredH = Height
                If DesiredW <= 0 Then DesiredW = ContentW
                If DesiredH <= 0 Then DesiredH = ContentH
                If AvailableWidth > 0 And DesiredW > AvailableWidth Then DesiredW = AvailableWidth
                If AvailableHeight > 0 And DesiredH > AvailableHeight Then DesiredH = AvailableHeight

                DesiredW = ClampElementWidth(Me, DesiredW)
                DesiredH = ClampElementHeight(Me, DesiredH)

                m_DesiredWidth = DesiredW
                m_DesiredHeight = DesiredH
                m_Layout.ReportDesiredSize DesiredW, DesiredH
                Exit Sub
            End If
        End If

        m_Layout.Measure AvailableWidth, AvailableHeight
        m_DesiredWidth = m_Layout.DesiredWidth
        m_DesiredHeight = m_Layout.DesiredHeight
    End Sub

    Public Sub MeasureLayout(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
        MeasureOverride AvailableWidth, AvailableHeight
    End Sub

    Private Function ReadChildLayoutWidth(ByVal Child As Object) As Double
        Dim Dep As IDependencyObject
        Dim P As Panel

        ReadChildLayoutWidth = 0#
        If Child Is Nothing Then Exit Function

        On Error Resume Next
        Err.Clear
        If TypeOf Child Is Panel Then
            Set P = Child
            ReadChildLayoutWidth = P.Width
            Exit Function
        End If

        Err.Clear
        If TypeOf Child Is IDependencyObject Then
            Set Dep = Child
            If Not Dep.DependencyProperties Is Nothing Then
                If Dep.DependencyProperties.Exists("Width") Then
                    ReadChildLayoutWidth = CDbl(Dep.DependencyProperties.GetValue("Width"))
                End If
            End If
        End If
        Err.Clear
    End Function

    Private Function ReadChildLayoutHeight(ByVal Child As Object) As Double
        Dim Dep As IDependencyObject
        Dim P As Panel

        ReadChildLayoutHeight = 0#
        If Child Is Nothing Then Exit Function

        On Error Resume Next
        Err.Clear
        If TypeOf Child Is Panel Then
            Set P = Child
            ReadChildLayoutHeight = P.Height
            Exit Function
        End If

        Err.Clear
        If TypeOf Child Is IDependencyObject Then
            Set Dep = Child
            If Not Dep.DependencyProperties Is Nothing Then
                If Dep.DependencyProperties.Exists("Height") Then
                    ReadChildLayoutHeight = CDbl(Dep.DependencyProperties.GetValue("Height"))
                End If
            End If
        End If
        Err.Clear
    End Function

    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        If Parent Is Nothing Then Exit Sub
        
        W.Move Left, Top, Width, Height
        
        ApplyVisibility W, Visibility
        ArrangeContent
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
        Set m_ContentHost = New ContentHost
        Set m_DependencyProperties = NewDependencyProperties(Me)
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ShowGridLines", vbBoolean)
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            Call .Register("Content", vbVariant, , , , NewDependencyPropertyMetadata(False, True, False, OneWay, ""))
            Call .Register("ContentTemplate", vbObject, "DataTemplate", , NewDependencyPropertyMetadata(False, True, False))
            .SetValue "Style", Nothing
            .SetValue "ContentTemplate", Nothing
        End With
        
        Set W = Cairo.WidgetBase
        W.CanGetFocus = False
        
        Set m_Children = NewUIElementCollection(Me)
        
        Set m_Layout = New FrameworkElement
        m_Layout.Initialize Me, Me, W
        m_Layout.RegisterLayoutProperties "ContentControl"
        
        m_Layout.Width = 0
        m_Layout.Height = 0
        m_Layout.Visibility = VisibilityVisible
        
        NewUIElementBase Me, m_BaseUIElement
        
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
        Dim Child As Object
        Dim ChildControl As IControl

        On Error GoTo Catch

        If Not e.OldItems Is Nothing Then
            For Each Child In e.OldItems
                If TypeOf Child Is IControl Then
                    Set ChildControl = Child
                    If Not ChildControl.Widget Is Nothing Then
                        If W.Widgets.Exists(ChildControl.Widget.Key) Then W.Widgets.Remove ChildControl.Widget.Key
                    End If
                End If
            Next
        End If

        If Not e.NewItems Is Nothing Then
            For Each Child In e.NewItems
                MoveChild Child
            Next
        End If

        SyncContentPresenter
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

    Public Sub ArrangeOverride(ByVal FinalWidth As Double, ByVal FinalHeight As Double)
        If W Is Nothing Then Exit Sub
        If m_Layout Is Nothing Then Exit Sub

        ' Measure Desired* even if arrange attach fails later.
        MeasureOverride FinalWidth, FinalHeight

        On Error GoTo AfterArrange
        If Not m_Children Is Nothing Then
            If m_Children.Count > 0 Then
                ArrangeDecoratorChild m_Children(0), W
            End If
        End If

AfterArrange:
        m_ActualWidth = FinalWidth
        m_ActualHeight = FinalHeight
        m_Layout.ReportActualSize m_ActualWidth, m_ActualHeight
    End Sub

    Private Sub ArrangeContent()
        If W Is Nothing Then Exit Sub
        ArrangeOverride CDbl(W.Width), CDbl(W.Height)
    End Sub

    Public Sub RelayoutChildren()
        ArrangeContent
    End Sub

    Private Sub ApplyContentValue(ByVal Value As Variant)
        ' Value unused — ApplyContentVisual reads Content + ContentTemplate from Props.
        m_ContentHost.ApplyContentVisual m_Children, m_DependencyProperties, m_BaseUIElement
    End Sub

    Private Sub ApplyContentVisual()
        m_ContentHost.ApplyContentVisual m_Children, m_DependencyProperties, m_BaseUIElement
    End Sub
'#End Region

'#Region DependencyProperties Events
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        Select Case Property.Name
            Case "Style":           ApplyStyle
            Case "Content", "ContentTemplate":
                ApplyContentVisual
                ArrangeContent
                modRenderCoalescer.RequestWidgetRefresh W
            Case Else
                If Not m_Children Is Nothing Then
                    If m_Layout.HandlesLayoutProperty(Property.Name) Then
                        If Property.Name = "Visibility" Then
                            m_Layout.OnLayoutPropertyChanged Property.Name, m_Children, W
                        Else
                            ArrangeContent
                        End If
                    End If
                End If
        End Select
        
        DependencyPropertiesStatic.PassPropertyValue m_Children, Property
    End Sub
'#End Region

'#Region Widget Events
    Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
        Dim dx As Double, dy As Double

        If (Not W.Visible) Or W.IsHidden Then Exit Sub

        dx = dx_Aligned
        dy = dy_Aligned

        CC.Save
        SyncContentPresenter
        m_ContentHost.DrawCaption CC, 0, 0, dx, dy, _
            W.FontName, W.FontSize, W.ForeColor, _
            W.FontBold, W.FontItalic, W.FontUnderline, W.FontStrikeThrough
        CC.Restore
    End Sub
    
    Private Sub W_Resize()
        ArrangeContent
    End Sub
'#End Region

Private Function GetBaseStyle() As Style
    Dim Style
    API.CopyVariable m_BaseUIElement.TryFindResource(TypeName(Me) & ".BaseStyle"), Style
    
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

