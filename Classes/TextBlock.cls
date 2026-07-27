VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TextBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IDependencyObject
Implements IUIElement
Implements IControl
Implements ICloneable

Private m_Left As Double, m_Top As Double, m_Width As Double, m_Height As Double

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Private WithEvents ActiveStyle As Style
Attribute ActiveStyle.VB_VarHelpID = -1

Private m_ParentPtr As Long
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase

Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1
Private m_SyncingVisibility As Boolean

Public Name As String

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Bindings = New List
        Set m_DependencyProperties = NewDependencyProperties(Me)
        
        ' Initialize the WidgetBase (W) just in case we want to use it as
        ' the source or target for the dependency properties
        Set W = Cairo.WidgetBase
        W.LockRefresh = True
        
        ' Hit-test enabled + HitResultHit=False -> events forward to parent (Button).
        W.ImplementsHitTest = True
        W.Visible = True
        
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ForeColor", vbLong, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontName", vbString, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontSize", vbSingle, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontBold", vbBoolean, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontUnderline", vbBoolean, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontItalic", vbBoolean, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontStrikeThrough", vbBoolean, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("ScaleFont", vbBoolean, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("HorizontalAlignment", vbLong, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("VerticalAlignment", vbLong, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("Text", vbString, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("Style", vbObject, "Style", , NewDependencyPropertyMetadata(False, True, False))
            Call .Register("Width", vbDouble, , , , NewDependencyPropertyMetadata(True, False, False))
            Call .Register("Height", vbDouble, , , , NewDependencyPropertyMetadata(True, False, False))
            Call .Register("Margin", vbObject, "Thickness", , NewDependencyPropertyMetadata(True, False, False))
            Call .Register("Visible", vbBoolean)
            Call .Register("Visibility", vbLong, , , , NewDependencyPropertyMetadata(False, True, False))
            Call .Register("IsHitTestVisible", vbBoolean, , , , NewDependencyPropertyMetadata(False, False, False))
            
            .SetValue "Style", Nothing
            .SetCurrentValue "Width", 0#
            .SetCurrentValue "Height", 0#
            .SetCurrentValue "Margin", NewThickness(0)
            .SetCurrentValue "Visible", True
            .SetCurrentValue "Visibility", CLng(VisibilityVisible)
            .SetCurrentValue "IsHitTestVisible", True
        
            ' Initialize the default CurrentValues of the Dependency Properties
            ' just in case the default style is not present.
            ' This should finally be replaced by built-in default styles
            ' inside the framework, which can be overwritten by the consumer application
            .SetCurrentValue "ForeColor", vbBlack
            .SetCurrentValue "FontName", "Arial"
            .SetCurrentValue "FontSize", 8.5
            .SetCurrentValue "FontBold", False
            .SetCurrentValue "FontUnderline", False
            .SetCurrentValue "FontItalic", False
            .SetCurrentValue "FontStrikeThrough", False
            .SetCurrentValue "ScaleFont", True
            .SetCurrentValue "HorizontalAlignment", 0
            .SetCurrentValue "VerticalAlignment", 2
        End With
                
        ' Initialize the Default metrics
        ' Important for any resize operations
        ' especially in the inner widgets
        
        NewUIElementBase Me, m_BaseUIElement

        'EnumerateProperties Me
        
        Set Style = GetBaseStyle
        W.LockRefresh = False
    End Sub
'#End Region

'#Region IDependencyObject Inheritance
    Public Property Get DependencyProperties() As DependencyProperties
        Set DependencyProperties = m_DependencyProperties
    End Property
'#End Region

Private Function ICloneable_Clone() As Object
    Set ICloneable_Clone = Clone
End Function

'#Region IDependencyObject Interface Implementation
    Private Property Get IDependencyObject_Children() As Object
        Set IDependencyObject_Children = Nothing
    End Property
    
    Private Property Get IDependencyObject_DependencyProperties() As DependencyProperties
        Set IDependencyObject_DependencyProperties = DependencyProperties
    End Property

    Private Property Get IDependencyObject_Parent() As IDependencyObject
        Set IDependencyObject_Parent = Parent
    End Property
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
    
    
    
    
    
    
    
    
    Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single)
        
        m_Left = 0 'Left
        m_Top = 0 'Top
        m_Width = Width
        m_Height = Height
        ' Do not write arranged size into Width/Height DPs (WPF: Actual* vs author size).
        ' Overwriting DPs made Grid treat text HorizontalAlignment=Left as layout Left
        ' with a sticky narrow width (CalcTextRowsInfo -> one glyph per row).
        
        If W Is Nothing Then Exit Sub
        W.Move Left, Top, Width, Height
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        If Parent Is Value Then Exit Property
        
        m_ParentPtr = ObjPtr(Value)
        
        If Value Is Nothing Then Exit Property
        
        DependencyPropertiesStatic.InheritPropertyValues Me
        If Not Parent.Widget Is Nothing Then Parent.Widget.Refresh
    End Property
        
    Public Property Get Parent() As IControl
        Call API.CopyVariable(API.ObjFromPtr(m_ParentPtr), Parent)
    End Property
'#End Region

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

    Public Property Get Width() As Double
        Width = CDbl(m_DependencyProperties.GetValue("Width"))
    End Property

    Public Property Let Width(ByVal Value As Double)
        Call m_DependencyProperties.SetValue("Width", Value)
    End Property

    Public Property Get Height() As Double
        Height = CDbl(m_DependencyProperties.GetValue("Height"))
    End Property

    Public Property Let Height(ByVal Value As Double)
        Call m_DependencyProperties.SetValue("Height", Value)
    End Property

    Public Property Get Margin() As Thickness
        Set Margin = m_DependencyProperties.GetValue("Margin")
    End Property

    Public Property Set Margin(ByVal Value As Thickness)
        Call m_DependencyProperties.SetValue("Margin", Value)
    End Property

    Public Property Get Visibility() As Visibility
        Visibility = CLng(m_DependencyProperties.GetValue("Visibility"))
    End Property

    Public Property Let Visibility(ByVal Value As Visibility)
        Call m_DependencyProperties.SetValue("Visibility", CLng(Value))
    End Property

    ' Compat facade: False -> VisibilityCollapsed (layout reclaim).
    Public Property Get Visible() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("Visible"), Visible)
    End Property

    Public Property Let Visible(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("Visible", Value)
    End Property

    Public Property Get IsHitTestVisible() As Boolean
        IsHitTestVisible = CBool(m_DependencyProperties.GetValue("IsHitTestVisible"))
    End Property

    Public Property Let IsHitTestVisible(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("IsHitTestVisible", Value)
    End Property
    
    Public Property Get ForeColor() As Long
        Call API.CopyVariable(m_DependencyProperties.GetValue("ForeColor"), ForeColor)
    End Property
    
    Public Property Let ForeColor(ByVal Value As Long)
        Call m_DependencyProperties.SetValue("ForeColor", Value)
    End Property
        
    Public Property Get FontName() As String
        Call API.CopyVariable(m_DependencyProperties.GetValue("FontName"), FontName)
    End Property
    
    Public Property Let FontName(ByVal Value As String)
        Call m_DependencyProperties.SetValue("FontName", Value)
    End Property
    
    Public Property Get FontSize() As Single
        Call API.CopyVariable(m_DependencyProperties.GetValue("FontSize"), FontSize)
    End Property
    
    Public Property Let FontSize(ByVal Value As Single)
        Call m_DependencyProperties.SetValue("FontSize", Value)
    End Property
    
    Public Property Get FontBold() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("FontBold"), FontBold)
    End Property
    
    Public Property Let FontBold(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("FontBold", Value)
    End Property

    Public Property Get FontItalic() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("FontItalic"), FontItalic)
    End Property
    
    Public Property Let FontItalic(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("FontItalic", Value)
    End Property

    Public Property Get FontUnderline() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("FontUnderline"), FontUnderline)
    End Property
    
    Public Property Let FontUnderline(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("FontUnderline", Value)
    End Property

    Public Property Get FontStrikeThrough() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("FontStrikeThrough"), FontStrikeThrough)
    End Property
    
    Public Property Let FontStrikeThrough(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("FontStrikeThrough", Value)
    End Property

    Public Property Get HorizontalAlignment() As Long
        Call API.CopyVariable(m_DependencyProperties.GetValue("HorizontalAlignment"), HorizontalAlignment)
    End Property
    
    Public Property Let HorizontalAlignment(ByVal Value As Long)
        Call m_DependencyProperties.SetValue("HorizontalAlignment", Value)
    End Property
    
    Public Property Get VerticalAlignment() As Long
        Call API.CopyVariable(m_DependencyProperties.GetValue("VerticalAlignment"), VerticalAlignment)
    End Property
    
    Public Property Let VerticalAlignment(ByVal Value As Long)
        Call m_DependencyProperties.SetValue("VerticalAlignment", Value)
    End Property
    
    Public Property Get ScaleFont() As Boolean
        Call API.CopyVariable(m_DependencyProperties.GetValue("ScaleFont"), ScaleFont)
    End Property
    
    Public Property Let ScaleFont(ByVal Value As Boolean)
        Call m_DependencyProperties.SetValue("ScaleFont", Value)
    End Property
    
    Public Property Let Text(ByVal Value As String)
        Call m_DependencyProperties.SetValue("Text", Value)
    End Property
    
    Public Property Get Text() As String
        Call API.CopyVariable(m_DependencyProperties.GetValue("Text"), Text)
    End Property
'#End Region

'#Region IVIsualChild Interface Implementation
    Private Sub IVisualChild_DrawOn(CC As vbRichClient5.cCairoContext, Optional ForeColor As Long = -1)
        DrawOn CC, ForeColor
    End Sub
'#End Region

'#Region IVIsualChild Inheritance
Private Sub DrawOn(CC As cCairoContext, Optional ForeColor As Long = -1)
    On Error Resume Next
    
        If Len(Me.Text) = 0 Then Exit Sub
        
        Dim Txt As String
        Dim TextWidth As Double
        Dim MaxWidth As Double
        Dim FontHeight As Double
        Dim TextHeight As Double
        Dim TextColor As Long
        
        Dim RowsCharCount() As Long
        Dim RowsCharOffset() As Long
        Dim RowCount As Long
        Dim MaxRowExtends As Single
        Dim MaxCharsPerRow As Long
        Dim MaxRowWidth As Double
        Dim MaxRowIndex As Long
        Dim Index As Integer
                        
        TextColor = Me.ForeColor
        If ForeColor <> -1 Then TextColor = ForeColor
        
        Txt = Replace$(Me.Text, Chr(13), vbCrLf)
        
        Const RECT_OFFSET = 2
        
        CC.SelectFont Me.FontName, Me.FontSize * GetScaleFactor, TextColor, Me.FontBold, Me.FontItalic, Me.FontUnderline, Me.FontStrikeThrough
        CC.CalcTextRowsInfo Txt, m_Width - 2 * RECT_OFFSET, True, False, RowsCharCount, RowsCharOffset, RowCount, MaxRowExtends
        
        For Index = 0 To RowCount - 1
            If RowsCharCount(Index) > MaxCharsPerRow Then
                MaxCharsPerRow = RowsCharCount(Index)
                MaxRowIndex = Index
            End If
            
            TextWidth = CC.GetTextExtents(Mid$(Txt, RowsCharOffset(Index) \ 2 + 1, Abs(RowsCharCount(Index))), FontHeight, TextHeight)
                        
            If TextWidth > MaxWidth Then MaxWidth = TextWidth
        Next
                               
        Dim R As Thickness
        Set R = EvalPosition(m_Width, m_Height, MaxWidth, (FontHeight + 0) * RowCount, RECT_OFFSET)
        
        '<Draw Dotted Border Around Text>
        'CC.Save
        'CC.Rectangle m_Left + R.Left, m_Top + R.Top, R.Left + R.Right, R.Top + R.Bottom
        'CC.SetLineWidth 0.5
        'CC.SetSourceColor vbBlack
        'CC.SetDashes 0, 1, 1
        'CC.Stroke
        'CC.Restore
        '</Draw Dotted Border Around Text>

        
        CC.DrawText m_Left + R.Left, m_Top + R.Top, R.Left + R.Right, R.Top + R.Bottom, Txt, False, HorizontalAlignment, , 1, dtHasAccelerators
        
        '<Draw Dotted Border Around Available Area>
        'CC.Save
        'CC.Rectangle m_Left, m_Top, m_Width, m_Height
        'CC.SetSourceColor vbMagenta
        'CC.SetLineWidth 2
        'CC.Stroke
        'CC.Restore
        '</Draw Dotted Border Around Available Area>
    End Sub
'#End Region

'#Region Helper Functions
    Private Function GetScaleFactor() As Double
Try:
        On Error GoTo Catch
        
        GetScaleFactor = 1
        
        If ScaleFont = False Then Exit Function
        
        ' Design-canvas font scaling removed with Design*; ScaleFont is currently a no-op.
        GetScaleFactor = 1
        Exit Function
        
Catch:
        '
    End Function
    
    Private Function GetRoot(o As IDependencyObject) As IDependencyObject
        If o.Parent Is Nothing Then
            Set GetRoot = o
        Else
            Set GetRoot = GetRoot(o.Parent)
        End If
    End Function
    
    Private Function EvalPosition(ByVal MaxW As Double, _
                                    ByVal MaxH As Double, _
                                    ByVal ActW As Double, _
                                    ByVal ActH As Double, _
                                    Optional ByVal Offset As Double = 0) As Thickness
                            
        Dim LL As Double, TT As Double, WW As Double, HH As Double
        
        Select Case HorizontalAlignment
            Case AlignmentConstants.vbLeftJustify
                LL = 0
                WW = ActW + 2 * Offset
            Case AlignmentConstants.vbRightJustify
                WW = ActW + 2 * Offset
                LL = MaxW - WW
            Case AlignmentConstants.vbCenter
                WW = ActW + 2 * Offset
                LL = (MaxW - WW) / 2
        End Select
        
        Select Case VerticalAlignment
            Case 0
                TT = 0
                HH = ActH + 2 * Offset
            Case 1
                HH = ActH + 2 * Offset
                TT = MaxH - HH
            Case 2
                HH = ActH + 2 * Offset
                TT = (MaxH - HH) / 2
        End Select
        
        Set EvalPosition = modConstructors.NewThickness(LL, TT, WW - LL, HH - TT)
    End Function
'#End Region

'#Region DependencyPropertiesEvents
    Private Sub m_DependencyProperties_DependencyPropertyChanged(ByVal Property As DependencyProperty, PreviousValue As Variant)
        On Error Resume Next
        
        Select Case Property.Name
            Case "Style":           ApplyStyle
            Case "Width":           m_Width = CDbl(Property.GetValue)
            Case "Height":          m_Height = CDbl(Property.GetValue)
            Case "Visible"
                If Not m_SyncingVisibility Then
                    m_SyncingVisibility = True
                    If CBool(Property.GetValue) Then
                        m_DependencyProperties.SetValue "Visibility", CLng(VisibilityVisible)
                    Else
                        m_DependencyProperties.SetValue "Visibility", CLng(VisibilityCollapsed)
                    End If
                    m_SyncingVisibility = False
                End If
            Case "Visibility"
                If Not m_SyncingVisibility Then
                    m_SyncingVisibility = True
                    m_DependencyProperties.SetCurrentValue "Visible", (CLng(Property.GetValue) = VisibilityVisible)
                    m_SyncingVisibility = False
                End If
                ApplyVisibility W, CLng(Property.GetValue)
                InvalidateParentLayout Me
            Case "IsHitTestVisible"
                ApplyIsHitTestVisible W, CBool(Property.GetValue)
        End Select
        
        If Parent Is Nothing Then Exit Sub
        If Property.Metadata.AffectsRender Then W.Refresh ' Parent.Widget.Refresh
    
    End Sub
'#End Region

Public Function Clone() As TextBlock
    Dim Props As cProperties
    Dim TargetProps As cProperties
    Dim Prop As cProperty
    
    Set Clone = New TextBlock
    
    Set Props = New_c.Properties
    Props.BindTo Me, True
    
    Set TargetProps = New_c.Properties
    TargetProps.BindTo Clone, True
    
    For Each Prop In Props
        If Prop.AllowsLet Then
            TargetProps(Prop.Name).Value = Prop.Value
        End If
    Next
    
    Dim SrcBinding As Binding
    Dim DstBinding As Binding
    
    For Each SrcBinding In m_Bindings
        If SrcBinding.TargetProperty Is Nothing Then GoTo NextBinding
        Set DstBinding = New Binding
        
        DstBinding.Initialize SrcBinding.Source, _
                                SrcBinding.Path, _
                                Clone, _
                                Clone.DependencyProperties.GetProperty(SrcBinding.TargetProperty.Name), _
                                SrcBinding.Converter, _
                                SrcBinding.ConverterParameter, _
                                SrcBinding.StringFormat
        
        If Not SrcBinding.SrcDepObj Is Nothing Then _
                Set DstBinding.SrcDepObj = Clone.DependencyProperties.GetProperty(SrcBinding.SrcDepObj.Name)
        
        Clone.Bindings.Add DstBinding
NextBinding:
    Next
    
    If Not W Is Nothing Then W.Move 0, 0, CSng(m_Width), CSng(m_Height)
End Function

Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean)
    HitResultHit = False
End Sub

Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    DrawOn CC
End Sub

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

    Friend Sub PrepareForUnload()
        On Error Resume Next
        Set ActiveStyle = Nothing
        Set W = Nothing
    End Sub

