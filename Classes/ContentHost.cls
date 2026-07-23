VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' Shared content-host implementation for ContentControl and Button (composition).
' Owns the host ContentPresenter; optional TemplatePresenter overrides ActivePresenter
' when a lookless ControlTemplate content slot is live (Button).
'
' Content visual precedence (3.8.0):
'   1. Explicit UIElement children (not ContentTemplate-generated)
'   2. Else ContentTemplate + Content as clone DataContext/item
'   3. Else string/Content via ContentPresenter paint

Private m_HostPresenter As ContentPresenter
Private m_TemplatePresenter As ContentPresenter
Private m_GeneratedRoot As Object

Public Sub Initialize()
    If m_HostPresenter Is Nothing Then Set m_HostPresenter = New ContentPresenter
End Sub

Public Property Get HostPresenter() As ContentPresenter
    Set HostPresenter = m_HostPresenter
End Property

' Live ControlTemplate content slot (Button); Nothing = use HostPresenter.
Public Property Get TemplatePresenter() As ContentPresenter
    Set TemplatePresenter = m_TemplatePresenter
End Property

Public Property Set TemplatePresenter(ByVal Value As ContentPresenter)
    Set m_TemplatePresenter = Value
End Property

Public Property Get ActivePresenter() As ContentPresenter
    If Not m_TemplatePresenter Is Nothing Then
        Set ActivePresenter = m_TemplatePresenter
    Else
        Set ActivePresenter = m_HostPresenter
    End If
End Property

Public Sub SyncContent(ByVal Props As DependencyProperties)
    Dim V As Variant
    Dim CP As ContentPresenter

    Set CP = ActivePresenter
    If CP Is Nothing Then Exit Sub
    If Props Is Nothing Then Exit Sub

    Call API.CopyVariable(Props.GetValue("Content"), V)
    If IsObject(V) Then
        Set CP.Content = V
    Else
        CP.Content = V
    End If
End Sub

Public Sub SyncSuppress(ByVal Suppress As Boolean)
    Dim CP As ContentPresenter

    Set CP = ActivePresenter
    If CP Is Nothing Then Exit Sub
    CP.SuppressContent = Suppress
End Sub

Public Sub SyncAlignment(ByVal Horizontal As Long, ByVal Vertical As Long)
    Dim CP As ContentPresenter

    Set CP = ActivePresenter
    If CP Is Nothing Then Exit Sub
    CP.HorizontalContentAlignment = Horizontal
    CP.VerticalContentAlignment = Vertical
End Sub

' Unified Content / ContentTemplate apply (ContentControl).
Public Sub ApplyContentVisual( _
    ByVal Children As UIElementCollection, _
    ByVal Props As DependencyProperties, _
    Optional ByVal ResourceHost As UIElementBase = Nothing)

    Dim Tmpl As DataTemplate
    Dim V As Variant
    Dim Root As Object
    Dim HasExplicit As Boolean
    Dim i As Long

    If Children Is Nothing Then Exit Sub
    If Props Is Nothing Then Exit Sub

    HasExplicit = False
    For i = 0 To Children.Count - 1
        If Not IsGeneratedRoot(Children(i)) Then
            HasExplicit = True
            Exit For
        End If
    Next

    If HasExplicit Then
        RemoveGeneratedRoots Children
        SyncContent Props
        SyncSuppress True
        Exit Sub
    End If

    Set Tmpl = TryGetContentTemplate(Props)
    Call API.CopyVariable(Props.GetValue("Content"), V)

    If Not Tmpl Is Nothing Then
        ClearAllChildren Children
        Set m_GeneratedRoot = Nothing
        Set Root = modItemTemplateEngine.CloneItemVisualForItem(Tmpl, V, ResourceHost)
        If Not Root Is Nothing Then
            If TypeOf Root Is IUIElement Then
                Children.Add Root
                Set m_GeneratedRoot = Root
                SyncContent Props
                SyncSuppress True
                Exit Sub
            End If
        End If
        ' Clone failed — fall through to string/IUIElement Content path.
    End If

    Set m_GeneratedRoot = Nothing
    ApplyContentValue Children, Props, V
End Sub

' ContentControl: IUIElement Content becomes the single visual child; string clears children.
Public Sub ApplyContentValue( _
    ByVal Children As UIElementCollection, _
    ByVal Props As DependencyProperties, _
    ByVal Value As Variant)

    Dim i As Long
    Dim El As IUIElement

    If Children Is Nothing Then Exit Sub

    If IsObject(Value) Then
        If Not Value Is Nothing Then
            If TypeOf Value Is IUIElement Then
                ClearAllChildren Children
                Set m_GeneratedRoot = Nothing
                Set El = Value
                Children.Add El
                SyncContent Props
                SyncSuppress True
                Exit Sub
            End If
        End If
    End If

    ClearAllChildren Children
    Set m_GeneratedRoot = Nothing
    SyncContent Props
    SyncSuppress False
End Sub

Public Function IsNestedTemplatePresenter() As Boolean
    If m_TemplatePresenter Is Nothing Then Exit Function
    If m_TemplatePresenter.Parent Is Nothing Then Exit Function
    IsNestedTemplatePresenter = True
End Function

Public Sub DrawCaption( _
    ByVal CC As cCairoContext, _
    ByVal Left As Double, _
    ByVal Top As Double, _
    ByVal Right As Double, _
    ByVal Bottom As Double, _
    ByVal FontName As String, _
    ByVal FontSize As Double, _
    ByVal ForeColor As Long, _
    Optional ByVal FontBold As Boolean = False, _
    Optional ByVal FontItalic As Boolean = False, _
    Optional ByVal FontUnderline As Boolean = False, _
    Optional ByVal FontStrikeThrough As Boolean = False)

    Dim CP As ContentPresenter

    Set CP = ActivePresenter
    If CP Is Nothing Then Exit Sub
    If Not CP.WouldDrawCaption Then Exit Sub
    CP.Draw CC, Left, Top, Right, Bottom, FontName, FontSize, ForeColor, _
        FontBold, FontItalic, FontUnderline, FontStrikeThrough
End Sub

Private Function TryGetContentTemplate(ByVal Props As DependencyProperties) As DataTemplate
    Dim V As Variant

    On Error Resume Next
    If Props Is Nothing Then Exit Function
    If Not Props.Exists("ContentTemplate") Then Exit Function
    Call API.CopyVariable(Props.GetValue("ContentTemplate"), V)
    If IsObject(V) Then
        If Not V Is Nothing Then
            If TypeOf V Is DataTemplate Then Set TryGetContentTemplate = V
        End If
    End If
    Err.Clear
End Function

Private Function IsGeneratedRoot(ByVal Child As Object) As Boolean
    If Child Is Nothing Then Exit Function
    If m_GeneratedRoot Is Nothing Then Exit Function
    IsGeneratedRoot = (Child Is m_GeneratedRoot)
End Function

Private Sub RemoveGeneratedRoots(ByVal Children As UIElementCollection)
    Dim i As Long

    If Children Is Nothing Then Exit Sub
    If m_GeneratedRoot Is Nothing Then Exit Sub

    For i = Children.Count - 1 To 0 Step -1
        If Children(i) Is m_GeneratedRoot Then Children.RemoveAt i
    Next
    Set m_GeneratedRoot = Nothing
End Sub

Private Sub ClearAllChildren(ByVal Children As UIElementCollection)
    Dim i As Long

    If Children Is Nothing Then Exit Sub
    For i = Children.Count - 1 To 0 Step -1
        Children.RemoveAt i
    Next
End Sub

Private Sub Class_Initialize()
    Initialize
End Sub
