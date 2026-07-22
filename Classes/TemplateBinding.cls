VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "TemplateBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' WPF shorthand: {TemplateBinding Path} ==
' {Binding Path, RelativeSource={RelativeSource TemplatedParent}, Mode=OneWay}

Implements IMarkupExtension

Public Path As String

Private m_Target As IDependencyObject
Private m_TargetProperty As String

Public Function InitializeFromMarkup( _
    ByVal Target As IDependencyObject, _
    ByVal TargetProperty As String, _
    ByVal MarkupProperties As String) As TemplateBinding

    Set m_Target = Target
    m_TargetProperty = TargetProperty
    Path = ParsePath(MarkupProperties)
    If Len(Path) = 0 Then Path = TargetProperty
    Set InitializeFromMarkup = Me
End Function

' Imperative helper (Phase0 / code).
Public Function Attach( _
    ByVal Target As IDependencyObject, _
    ByVal TargetPropertyName As String, _
    Optional ByVal SourcePath As String) As BindingExpression

    Dim Expr As BindingExpression
    Dim RS As RelativeSource
    Dim Src As Object
    Dim P As String

    If Target Is Nothing Then Exit Function
    P = SourcePath
    If Len(P) = 0 Then P = TargetPropertyName

    Set RS = New RelativeSource
    RS.Mode = RelativeSourceTemplatedParent
    Set Src = RS.Resolve(API.CObj(Target))
    If Src Is Nothing Then Exit Function

    Set Expr = New BindingExpression
    If Expr.Attach(Target, TargetPropertyName, Src, P, OneWay) Is Nothing Then Exit Function
    Set Attach = Expr
End Function

Private Function ParsePath(ByVal MarkupProperties As String) As String
    Dim Parts
    Dim P As Variant
    Dim Body
    Dim Key As String
    Dim Val As String

    MarkupProperties = Trim$(MarkupProperties)
    If Len(MarkupProperties) = 0 Then Exit Function

    Parts = Split(MarkupProperties, ",")
    For Each P In Parts
        If InStr(1, P, "=") > 0 Then
            Body = Split(Trim$(P), "=", 2)
            Key = LCase$(Trim$(Body(0)))
            Val = Trim$(Body(1))
            If Key = "path" Or Key = "" Then
                ParsePath = Val
                Exit Function
            End If
        Else
            ParsePath = Trim$(P)
            Exit Function
        End If
    Next
End Function

Private Function IMarkupExtension_ProvideValue() As Variant
    Dim B As Binding
    Dim RS As RelativeSource
    Dim Src As Object
    Dim TargetProp As DependencyProperty

    On Error GoTo Handler

    If m_Target Is Nothing Then Exit Function
    If Len(m_TargetProperty) = 0 Then Exit Function
    If Not m_Target.DependencyProperties.Exists(m_TargetProperty) Then Exit Function

    Set TargetProp = m_Target.DependencyProperties.GetProperty(m_TargetProperty)

    Set RS = New RelativeSource
    RS.Mode = RelativeSourceTemplatedParent
    Set Src = RS.Resolve(API.CObj(m_Target))
    If Src Is Nothing Then Exit Function

    Set B = New Binding
    B.Mode = OneWay
    B.Path = Path
    Set B.TargetProperty = TargetProp
    Set B.Source = Src
    Set B.Target = m_Target
    API.CObj(m_Target).Bindings.Add B

    Call API.CopyVariable(B.ProvideValue, IMarkupExtension_ProvideValue)
    Exit Function

Handler:
    Err.Clear
End Function
