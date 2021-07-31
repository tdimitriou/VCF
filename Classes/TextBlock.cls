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
Implements IVisualChild
Implements ICloneable

Private m_Left As Double, m_Top As Double, m_Width As Double, m_Height As Double

Private m_DesignLeft As Single
Private m_DesignTop As Single
Private m_DesignWidth As Single
Private m_DesignHeight As Single
Private m_ParentPtr As Long
Private m_Bindings As List
Private m_BaseUIElement As UIElementBase

Private WithEvents m_DependencyProperties As DependencyProperties
Attribute m_DependencyProperties.VB_VarHelpID = -1

Public Name As String

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_Bindings = New List
        Set m_DependencyProperties = NewDependencyProperties(Me)
        With m_DependencyProperties
            Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
            Call .Register("ForeColor", vbLong, , , , NewDependencyPropertyMetadata(, True))
            Call .Register("FontName", vbString, , , "", NewDependencyPropertyMetadata(, True))
            Call .Register("FontSize", vbSingle, , , 9, NewDependencyPropertyMetadata(, True))
            Call .Register("FontBold", vbBoolean, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("FontUnderline", vbBoolean, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("FontItalic", vbBoolean, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("FontStrikeThrough", vbBoolean, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("ScaleFont", vbBoolean, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("HorizontalAlignment", vbLong, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("VerticalAlignment", vbLong, , , 0, NewDependencyPropertyMetadata(, True))
            Call .Register("Text", vbString, , , "", NewDependencyPropertyMetadata(, True))
            
            ForeColor = vbBlack
            FontName = "Arial"
            FontSize = 8.5
            FontBold = False
            FontUnderline = False
            FontItalic = False
            FontStrikeThrough = False
            ScaleFont = True
            HorizontalAlignment = 0
            VerticalAlignment = 2
        End With
    
        m_DesignLeft = 0
        m_DesignTop = 0
        m_DesignWidth = 100
        m_DesignHeight = 30
        
        NewUIElementBase Me, m_BaseUIElement
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
        m_Left = Left
        m_Top = Top
        m_Width = Width
        m_Height = Height
    End Sub

    Public Property Set Parent(ByVal Value As IControl)
        If Parent Is Value Then Exit Property
        
        m_ParentPtr = ObjPtr(Value)
        
        If Value Is Nothing Then Exit Property
        
        DependencyPropertiesStatic.InheritPropertyValues Me
        
        Parent.Widget.Refresh
    End Property
        
    Public Property Get Parent() As IControl
        Call API.CopyVariable(API.ObjFromPtr(m_ParentPtr), Parent)
    End Property
'#End Region

'#Region Public Properties
    Public Property Get Bindings() As VCF.List
        Set Bindings = m_Bindings
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
        
        Txt = Me.Text
        
        CC.SelectFont Me.FontName, Me.FontSize * GetScaleFactor, TextColor, Me.FontBold, Me.FontItalic, Me.FontUnderline, Me.FontStrikeThrough
        CC.CalcTextRowsInfo Txt, m_Width - 4, True, False, RowsCharCount, RowsCharOffset, RowCount, MaxRowExtends
        
        For Index = 0 To RowCount - 1
            If RowsCharCount(Index) > MaxCharsPerRow Then
                MaxCharsPerRow = RowsCharCount(Index)
                MaxRowIndex = Index
            End If
            
            TextWidth = CC.GetTextExtents(Mid$(Txt, RowsCharOffset(Index) / 2 + 1, RowsCharCount(Index)), FontHeight, TextHeight)
            If TextWidth > MaxWidth Then MaxWidth = TextWidth
        Next
               
        'TextWidth = CC.GetTextExtents(Mid$(Me.Text, RowsCharOffset(MaxRowIndex) / 2 + 1, RowsCharCount(MaxRowIndex)), FontHeight, TextHeight)
        
        Const RECT_OFFSET = 2
        Dim R As Thickness
        'Set R = EvalPosition(m_Width, m_Height, TextWidth, FontHeight * RowCount, RECT_OFFSET)
        Set R = EvalPosition(m_Width, m_Height, MaxWidth, FontHeight * RowCount, RECT_OFFSET)
        
        '<Draw Dotted Border Around Text>
        'CC.Save
        'CC.Rectangle m_Left + R.Left, m_Top + R.Top, R.Left + R.Right, R.Top + R.Bottom
        'CC.SetLineWidth 0.5
        'CC.SetSourceColor vbBlack
        'CC.SetDashes 0, 1, 1
        'CC.Stroke
        'CC.Restore
        '</Draw Dotted Border Around Text>


        CC.DrawText m_Left + R.Left, m_Top + R.Top, R.Left + R.Right, R.Top + R.Bottom, Txt, , HorizontalAlignment, , 1

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
        Dim Root As IDependencyObject
        Dim RootElement As IUIElement
        Dim RootControl As IControl
        
Try:
        On Error GoTo Catch
        
        GetScaleFactor = 1
        
        If ScaleFont = False Then Exit Function
        
        Set Root = GetRoot(Me)
        Set RootElement = Root
        Set RootControl = Root
        
        Dim xF As Double, yF As Double
        
        xF = RootControl.Widget.Width / RootElement.DesignWidth
        yF = RootControl.Widget.Height / RootElement.DesignHeight
        
        GetScaleFactor = IIf(yF < xF, yF, xF)
        
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
        
        If Parent Is Nothing Then Exit Sub
        If Property.Metadata.AffectsRender Then Parent.Widget.Refresh
    
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
        Set DstBinding = New Binding
        
        DstBinding.Initialize SrcBinding.Source, _
                                SrcBinding.Path, _
                                Clone, _
                                Clone.DependencyProperties.GetProperty(SrcBinding.TargetProperty.Name), _
                                SrcBinding.Converter, _
                                SrcBinding.StringFormat
        
        If Not SrcBinding.SrcDepObj Is Nothing Then _
                Set DstBinding.SrcDepObj = Clone.DependencyProperties.GetProperty(SrcBinding.SrcDepObj.Name)
        
        Clone.Bindings.Add DstBinding
    Next
    
    Clone.Move CSng(m_Left), CSng(m_Top), CSng(m_Width), CSng(m_Height)
End Function
