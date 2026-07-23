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

Friend Sub Initialize( _
    ByVal Host As IDependencyObject, _
    ByVal UIHost As IUIElement, _
    ByVal W As cWidgetBase)

    m_HostPtr = ObjPtr(Host)
    m_UIPtr = ObjPtr(UIHost)
    m_WidgetPtr = ObjPtr(W)
End Sub

Friend Sub RegisterLayoutProperties(Optional ByVal TypeName As String = "FrameworkElement")
    modStaticClasses.DependencyPropertyRegistry.EnsureBuiltInTypes
    modStaticClasses.DependencyPropertyRegistry.ApplyTo TypeName, Host.DependencyProperties
    SetDefaultLayoutValues
End Sub

Public Property Get Width() As Double
    If Host.DependencyProperties.Exists("Width") Then
        Width = CDbl(Host.DependencyProperties.GetValue("Width"))
    End If
End Property

Public Property Let Width(ByVal Value As Double)
    If Host.DependencyProperties.Exists("Width") Then
        Host.DependencyProperties.SetValue "Width", Value
    End If
End Property

Public Property Get Height() As Double
    If Host.DependencyProperties.Exists("Height") Then
        Height = CDbl(Host.DependencyProperties.GetValue("Height"))
    End If
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
    If Host.DependencyProperties.Exists("MinWidth") Then
        MinWidth = CDbl(Host.DependencyProperties.GetValue("MinWidth"))
    End If
End Property

Public Property Let MinWidth(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MinWidth") Then
        Host.DependencyProperties.SetValue "MinWidth", Value
    End If
End Property

Public Property Get MinHeight() As Double
    If Host.DependencyProperties.Exists("MinHeight") Then
        MinHeight = CDbl(Host.DependencyProperties.GetValue("MinHeight"))
    End If
End Property

Public Property Let MinHeight(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MinHeight") Then
        Host.DependencyProperties.SetValue "MinHeight", Value
    End If
End Property

Public Property Get MaxWidth() As Double
    If Host.DependencyProperties.Exists("MaxWidth") Then
        MaxWidth = CDbl(Host.DependencyProperties.GetValue("MaxWidth"))
    End If
End Property

Public Property Let MaxWidth(ByVal Value As Double)
    If Host.DependencyProperties.Exists("MaxWidth") Then
        Host.DependencyProperties.SetValue "MaxWidth", Value
    End If
End Property

Public Property Get MaxHeight() As Double
    If Host.DependencyProperties.Exists("MaxHeight") Then
        MaxHeight = CDbl(Host.DependencyProperties.GetValue("MaxHeight"))
    End If
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
End Sub

Public Sub Arrange(ByVal FinalLeft As Single, ByVal FinalTop As Single, ByVal FinalWidth As Single, ByVal FinalHeight As Single)
    UIHost.Move FinalLeft, FinalTop, FinalWidth, FinalHeight
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

    For Each Child In Children
        If Not TypeOf Child Is IUIElement Then GoTo NextChild
        Set ChildUI = Child

        ChildVis = ReadChildVisibility(Child)
        If IsLayoutCollapsed(ChildVis) Then
            If TypeOf Child Is IControl Then
                If HostWidget.Widgets.Exists(Child.Widget.Key) Then HostWidget.Widgets.Remove Child.Widget.Key
            End If
            GoTo NextChild
        End If

        If Not TypeOf Child Is IControl Then GoTo NextChild

        If HostWidget.Widgets.Exists(Child.Widget.Key) Then HostWidget.Widgets.Remove Child.Widget.Key
        If Not HostWidget.Widgets.Exists(Child.Widget.Key) Then
            HostWidget.Widgets.Add Child, "_" & ObjPtr(Child), , , , , IsWidgetVisible(ChildVis)
        End If

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

    ArrangeChildren Children, HostWidget, OverrideHostWidth, OverrideHostHeight
End Sub

Public Function HandlesLayoutProperty(ByVal PropertyName As String) As Boolean
    Select Case PropertyName
        Case "Width", "Height", "Margin", "Visibility", "MinWidth", "MinHeight", "MaxWidth", "MaxHeight"
            HandlesLayoutProperty = True
    End Select
End Function

Public Sub OnLayoutPropertyChanged( _
    ByVal PropertyName As String, _
    ByVal Children As UIElementCollection, _
    ByVal HostWidget As cWidgetBase, _
    Optional ByVal OverrideHostWidth As Single = 0, _
    Optional ByVal OverrideHostHeight As Single = 0)

    Select Case PropertyName
        Case "Visibility"
            ApplyVisibilityToWidget
            InvalidateParentLayout Host
        Case "Width", "Height", "Margin", "MinWidth", "MinHeight", "MaxWidth", "MaxHeight"
            OnHostResize Children, HostWidget, OverrideHostWidth, OverrideHostHeight
            InvalidateParentLayout Host
    End Select
End Sub

Private Sub SetDefaultLayoutValues()
    With Host.DependencyProperties
        If .Exists("Width") Then .SetCurrentValue "Width", 300#
        If .Exists("Height") Then .SetCurrentValue "Height", 300#
        If .Exists("MinWidth") Then .SetCurrentValue "MinWidth", 0#
        If .Exists("MinHeight") Then .SetCurrentValue "MinHeight", 0#
        If .Exists("MaxWidth") Then .SetCurrentValue "MaxWidth", 0#
        If .Exists("MaxHeight") Then .SetCurrentValue "MaxHeight", 0#
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
    On Error Resume Next
    If Child.DependencyProperties.Exists("Width") Then
        ReadChildWidth = CDbl(Child.DependencyProperties.GetValue("Width"))
    End If
End Function

Private Function ReadChildHeight(ByVal Child As Object) As Double
    On Error Resume Next
    If Child.DependencyProperties.Exists("Height") Then
        ReadChildHeight = CDbl(Child.DependencyProperties.GetValue("Height"))
    End If
End Function

Private Sub ApplyChildVisibility(ByVal Child As Object, ByVal Value As Visibility)
    On Error Resume Next
    If TypeOf Child Is IControl Then
        ApplyVisibility Child.Widget, Value
    End If
End Sub
