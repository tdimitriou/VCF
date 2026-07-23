VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FrameworkElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_HostPtr As Long
Private m_UIPtr As Long
Private m_WidgetPtr As Long
Private m_DesiredWidth As Double
Private m_DesiredHeight As Double
Private m_ActualWidth As Double
Private m_ActualHeight As Double
Private m_MeasureDirty As Boolean
Private m_ArrangeDirty As Boolean

Friend Sub Initialize( _
    ByVal Host As IDependencyObject, _
    ByVal UIHost As IUIElement, _
    ByVal W As cWidgetBase)

    m_HostPtr = ObjPtr(Host)
    m_UIPtr = ObjPtr(UIHost)
    m_WidgetPtr = ObjPtr(W)
    m_MeasureDirty = True
    m_ArrangeDirty = True
End Sub

Friend Sub RegisterLayoutProperties(Optional ByVal TypeName As String = "FrameworkElement")
    modStaticClasses.DependencyPropertyRegistry.EnsureBuiltInTypes
    modStaticClasses.DependencyPropertyRegistry.ApplyTo TypeName, Host.DependencyProperties
    SetDefaultLayoutValues
End Sub

Public Property Get Width() As Double
    Width = CoerceLayoutDouble(Host.DependencyProperties, "Width", 0#)
End Property

Public Property Let Width(ByVal Value As Double)
    If Host.DependencyProperties.Exists("Width") Then
        Host.DependencyProperties.SetValue "Width", Value
    End If
End Property

Public Property Get Height() As Double
    Height = CoerceLayoutDouble(Host.DependencyProperties, "Height", 0#)
End Property

Public Property Let Height(ByVal Value As Double)
    If Host.DependencyProperties.Exists("Height") Then
        Host.DependencyProperties.SetValue "Height", Value
    End If
End Property

Public Property Get Margin() As Thickness
    If Host.DependencyProperties.Exists("Margin") Then
        Set Margin = Host.DependencyProperties.GetValue("Margin")
    End If
End Property

Public Property Set Margin(ByVal Value As Thickness)
    If Host.DependencyProperties.Exists("Margin") Then
        Host.DependencyProperties.SetValue "Margin", Value
    End If
End Property

Public Property Get Visibility() As Visibility
    If Host.DependencyProperties.Exists("Visibility") Then
        Visibility = CLng(Host.DependencyProperties.GetValue("Visibility"))
    Else
        Visibility = VisibilityVisible
    End If
End Property

Public Property Let Visibility(ByVal Value As Visibility)
    If Host.DependencyProperties.Exists("Visibility") Then
        Host.DependencyProperties.SetValue "Visibility", CLng(Value)
    End If
    ApplyVisibilityToWidget
End Property

Public Property Get MinWidth() As Double
    MinWidth = CoerceLayoutDouble(Host.DependencyProperties, "MinWidth", 0#)
End Property

Public Property Let MinWidth(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MinWidth") Then
        Host.DependencyProperties.SetValue "MinWidth", Value
    End If
End Property

Public Property Get MinHeight() As Double
    MinHeight = CoerceLayoutDouble(Host.DependencyProperties, "MinHeight", 0#)
End Property

Public Property Let MinHeight(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MinHeight") Then
        Host.DependencyProperties.SetValue "MinHeight", Value
    End If
End Property

Public Property Get MaxWidth() As Double
    MaxWidth = CoerceLayoutDouble(Host.DependencyProperties, "MaxWidth", 0#)
End Property

Public Property Let MaxWidth(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MaxWidth") Then
        Host.DependencyProperties.SetValue "MaxWidth", Value
    End If
End Property

Public Property Get MaxHeight() As Double
    MaxHeight = CoerceLayoutDouble(Host.DependencyProperties, "MaxHeight", 0#)
End Property

Public Property Let MaxHeight(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MaxHeight") Then
        Host.DependencyProperties.SetValue "MaxHeight", Value
    End If
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

' Panels (StackPanel MeasureOverride) report content-driven desired size.
Friend Sub ReportDesiredSize(ByVal DesiredW As Double, ByVal DesiredH As Double)
    m_DesiredWidth = DesiredW
    m_DesiredHeight = DesiredH
    m_MeasureDirty = False
End Sub

' Sync Actual* from the final widget/host slot without re-Move (parent already positioned us).
Friend Sub ReportActualSize(ByVal ActualW As Double, ByVal ActualH As Double)
    m_ActualWidth = ActualW
    m_ActualHeight = ActualH
    m_ArrangeDirty = False
End Sub

Public Sub InvalidateMeasure()
    m_MeasureDirty = True
    m_ArrangeDirty = True
End Sub

Public Sub InvalidateArrange()
    m_ArrangeDirty = True
End Sub

Public Property Get IsMeasureDirty() As Boolean
    IsMeasureDirty = m_MeasureDirty
End Property

Public Property Get IsArrangeDirty() As Boolean
    IsArrangeDirty = m_ArrangeDirty
End Property

Public Sub Measure(ByVal AvailableWidth As Double, ByVal AvailableHeight As Double)
    Dim W As Double
    Dim H As Double

    W = Width
    H = Height

    If W <= 0 Then W = AvailableWidth
    If H <= 0 Then H = AvailableHeight
    If AvailableWidth > 0 And W > AvailableWidth Then W = AvailableWidth
    If AvailableHeight > 0 And H > AvailableHeight Then H = AvailableHeight

    W = ClampElementWidth(Host, W)
    H = ClampElementHeight(Host, H)

    m_DesiredWidth = W
    m_DesiredHeight = H
    m_MeasureDirty = False
    m_ArrangeDirty = True
End Sub

Public Sub Arrange(ByVal FinalLeft As Single, ByVal FinalTop As Single, ByVal FinalWidth As Single, ByVal FinalHeight As Single)
    UIHost.Move FinalLeft, FinalTop, FinalWidth, FinalHeight
    m_ActualWidth = CDbl(FinalWidth)
    m_ActualHeight = CDbl(FinalHeight)
    m_ArrangeDirty = False
End Sub

Public Sub ArrangeChildren( _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Dim Child As Object
    Dim R As LayoutRect
    Dim ChildUI As IUIElement
    Dim ChildVis As Visibility
    Dim HostWidth As Single
    Dim HostHeight As Single
    Dim HostObj As Object
    Dim FillWindowContent As Boolean
    Dim FillUserControlRoot As Boolean

    If Children Is Nothing Then Exit Sub
    If HostWidget Is Nothing Then Exit Sub

    If OverrideHostWidth > 0 Then
        HostWidth = OverrideHostWidth
    Else
        HostWidth = HostWidget.Width
    End If
    If OverrideHostHeight > 0 Then
        HostHeight = OverrideHostHeight
    Else
        HostHeight = HostWidget.Height
    End If

    FillWindowContent = False
    FillUserControlRoot = False
    On Error Resume Next
    Set HostObj = Host
    If Not HostObj Is Nothing Then
        If TypeOf HostObj Is Window Then FillWindowContent = True
        ' Single root under UserControl fills the host (Measure/Arrange panels).
        If TypeOf HostObj Is UserControl Then
            If Children.Count = 1 Then FillUserControlRoot = True
        End If
    End If
    Err.Clear
    On Error GoTo 0

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextChild
        Set ChildUI = Child

        ChildVis = ReadChildVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then
            If TypeOf Child Is IControl Then
                DetachCollapsedChild Child, HostWidget
            End If
            GoTo NextChild
        End If

        If Not TypeOf Child Is IControl Then GoTo NextChild

        AttachChildWidget Child, HostWidget, ChildVis

        ' Window content host (WPF): visible views fill the client area.
        If FillWindowContent Then
            If TypeOf Child Is IUserControl Or TypeOf Child Is UserControl Then
                ChildUI.Move 0!, 0!, HostWidth, HostHeight
                ApplyChildVisibility Child, ChildVis
                GoTo NextChild
            End If
        End If

        ' UserControl single root (typically Grid/Stack): fill client; no canvas scale.
        If FillUserControlRoot Then
            ChildUI.Move 0!, 0!, HostWidth, HostHeight
            ApplyChildVisibility Child, ChildVis
            GoTo NextChild
        End If

        ' Absolute Margin/Width/Height = fixed pixels on multi-child UC/Border/Panel.
        R = LayoutRectFromMargin(ReadChildMargin(Child), ReadChildWidth(Child), ReadChildHeight(Child))
        R.Width = CSng(ClampElementWidth(Child, R.Width))
        R.Height = CSng(ClampElementHeight(Child, R.Height))

        ApplyLayoutRectToElement ChildUI, R
        ApplyChildVisibility Child, ChildVis

NextChild:
    Next
End Sub

Public Sub OnHostResize( _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Dim AW As Single
    Dim AH As Single

    ArrangeChildren Children, HostWidget, OverrideHostWidth, OverrideHostHeight

    If OverrideHostWidth > 0 Then
        AW = OverrideHostWidth
    ElseIf Not HostWidget Is Nothing Then
        AW = HostWidget.Width
    End If
    If OverrideHostHeight > 0 Then
        AH = OverrideHostHeight
    ElseIf Not HostWidget Is Nothing Then
        AH = HostWidget.Height
    End If
    ReportActualSize CDbl(AW), CDbl(AH)
End Sub

Public Function HandlesLayoutProperty(ByVal PropertyName As String) As Boolean
    Select Case PropertyName
        Case "Width", "Height", "Margin", "Visibility", "MinWidth", "MinHeight", "MaxWidth", "MaxHeight", "HorizontalAlignment", "VerticalAlignment"
            HandlesLayoutProperty = True
    End Select
End Function

Public Sub OnLayoutPropertyChanged( _
    ByVal PropertyName As String, _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    On Error Resume Next

    Select Case PropertyName
        Case "Visibility"
            ApplyVisibilityToWidget
            InvalidateMeasure
            If Not Host Is Nothing Then InvalidateParentLayout Host
        Case "Width", "Height", "Margin", "MinWidth", "MinHeight", "MaxWidth", "MaxHeight", "HorizontalAlignment", "VerticalAlignment"
            InvalidateMeasure
            OnHostResize Children, HostWidget, OverrideHostWidth, OverrideHostHeight
            If Not Host Is Nothing Then InvalidateParentLayout Host
    End Select
End Sub

Private Sub SetDefaultLayoutValues()
    With Host.DependencyProperties
        ' Unset Width/Height = 0 so MeasureLayout uses content (Auto tracks).
        If .Exists("Width") Then .SetCurrentValue "Width", 0#
        If .Exists("Height") Then .SetCurrentValue "Height", 0#
        If .Exists("MinWidth") Then .SetCurrentValue "MinWidth", 0#
        If .Exists("MinHeight") Then .SetCurrentValue "MinHeight", 0#
        If .Exists("MaxWidth") Then .SetCurrentValue "MaxWidth", 0#
        If .Exists("MaxHeight") Then .SetCurrentValue "MaxHeight", 0#
        If .Exists("HorizontalAlignment") Then .SetCurrentValue "HorizontalAlignment", "Stretch"
        If .Exists("VerticalAlignment") Then .SetCurrentValue "VerticalAlignment", "Stretch"
        If .Exists("Visibility") Then .SetCurrentValue "Visibility", CLng(VisibilityVisible)
        If .Exists("Margin") Then
            Dim DefaultMargin As Thickness
            Set DefaultMargin = NewThickness(0)
            .SetCurrentValue "Margin", DefaultMargin
        End If
    End With
End Sub

Private Sub ApplyVisibilityToWidget()
    ApplyVisibility Widget, Visibility
End Sub

Private Property Get Host() As IDependencyObject
    Set Host = API.ObjFromPtr(m_HostPtr)
End Property

Private Property Get UIHost() As IUIElement
    Set UIHost = API.ObjFromPtr(m_UIPtr)
End Property

Private Property Get Widget() As cWidgetBase
    Set Widget = API.ObjFromPtr(m_WidgetPtr)
End Property

Private Function CoerceLayoutDouble(ByVal Props As DependencyProperties, ByVal Name As String, ByVal DefaultValue As Double) As Double
    Dim V As Variant

    CoerceLayoutDouble = DefaultValue
    On Error Resume Next
    If Props Is Nothing Then Exit Function
    If Not Props.Exists(Name) Then Exit Function
    Call API.CopyVariable(Props.GetValue(Name), V)
    If IsNull(V) Or IsEmpty(V) Then Exit Function
    If Not IsNumeric(V) Then Exit Function
    CoerceLayoutDouble = CDbl(V)
End Function

Private Function ReadChildVisibility(ByVal Child As Object) As Visibility
    On Error Resume Next
    If TypeOf Child Is IControl Then
        If Child.DependencyProperties.Exists("Visibility") Then
            ReadChildVisibility = CLng(Child.DependencyProperties.GetValue("Visibility"))
            Exit Function
        End If
        If Child.DependencyProperties.Exists("Visible") Then
            If CBool(Child.DependencyProperties.GetValue("Visible")) Then
                ReadChildVisibility = VisibilityVisible
            Else
                ReadChildVisibility = VisibilityCollapsed
            End If
            Exit Function
        End If
    End If
    If TypeOf Child Is IUIElement Then
        Select Case TypeName(Child)
            Case "Panel", "UserControl", "StackPanel", "Grid", "ContentControl", "Border", "Button"
                ReadChildVisibility = Child.Visibility
                Exit Function
        End Select
    End If
    ReadChildVisibility = VisibilityVisible
End Function

Private Function ReadChildMargin(ByVal Child As Object) As Thickness
    On Error Resume Next
    If Child.DependencyProperties.Exists("Margin") Then
        Set ReadChildMargin = Child.DependencyProperties.GetValue("Margin")
    End If
    If ReadChildMargin Is Nothing Then
        Set ReadChildMargin = modConstructors.NewThickness(0, 0, 0, 0)
    End If
End Function

Private Function ReadChildWidth(ByVal Child As Object) As Double
    Dim V As Variant
    On Error Resume Next
    If Child.DependencyProperties.Exists("Width") Then
        Call API.CopyVariable(Child.DependencyProperties.GetValue("Width"), V)
        If Not IsNull(V) And Not IsEmpty(V) And IsNumeric(V) Then ReadChildWidth = CDbl(V)
    End If
End Function

Private Function ReadChildHeight(ByVal Child As Object) As Double
    Dim V As Variant
    On Error Resume Next
    If Child.DependencyProperties.Exists("Height") Then
        Call API.CopyVariable(Child.DependencyProperties.GetValue("Height"), V)
        If Not IsNull(V) And Not IsEmpty(V) And IsNumeric(V) Then ReadChildHeight = CDbl(V)
    End If
End Function

Private Sub ApplyChildVisibility(ByVal Child As Object, ByVal Value As Visibility)
    ApplyChildWidgetVisibility Child, Value
End Sub
