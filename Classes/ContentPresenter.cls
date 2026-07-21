VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "ContentPresenter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' §2.11 — paint-only ContentPresenter (first slice).
' No cWidgetBase / Children in this release — avoids Cairo widget attach churn.
' Button.DrawContent delegates string Content painting here.

Private m_Content As Variant
Private m_SuppressContent As Boolean

Public Name As String

Public Property Get Content() As Variant
    Call API.CopyVariable(m_Content, Content)
End Property

Public Property Let Content(ByVal Value As Variant)
    If IsObject(Value) Then
        Set m_Content = Value
    Else
        m_Content = Value
    End If
End Property

Public Property Set Content(ByVal Value As Object)
    Set m_Content = Value
End Property

' When True, string Content is not painted (e.g. host has visual children).
Public Property Get SuppressContent() As Boolean
    SuppressContent = m_SuppressContent
End Property

Public Property Let SuppressContent(ByVal Value As Boolean)
    m_SuppressContent = Value
End Property

' Non-object Content coerced to string; Empty/Null/object → "".
Public Function ContentCaption() As String
    Dim V As Variant

    Call API.CopyVariable(m_Content, V)
    If IsObject(V) Then Exit Function
    If IsNull(V) Or IsEmpty(V) Then Exit Function
    ContentCaption = CStr(V)
End Function

' True when a string caption would be drawn (not suppressed, non-empty).
Public Property Get WouldDrawCaption() As Boolean
    WouldDrawCaption = (Not m_SuppressContent) And (Len(ContentCaption()) > 0)
End Property

' Paint centered caption into [Left,Top]-[Right,Bottom]. No-op when suppressed/empty.
Public Sub Draw( _
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

    Dim Cap As String

    If m_SuppressContent Then Exit Sub
    Cap = ContentCaption()
    If Len(Cap) = 0 Then Exit Sub

    CC.SelectFont FontName, FontSize, ForeColor, FontBold, FontItalic, FontUnderline, FontStrikeThrough
    CC.DrawText Left, Top, Right, Bottom, Cap, False, AlignmentConstants.vbCenter, , 1, dtHasAccelerators
End Sub

Private Sub Class_Initialize()
    m_Content = Empty
    m_SuppressContent = False
End Sub
