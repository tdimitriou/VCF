VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "RelativeSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IMarkupExtension

Public Enum RelativeSourceMode
    RelativeSourceSelf = 0
    RelativeSourceTemplatedParent = 1
    RelativeSourceFindAncestor = 2
End Enum

Public Mode As RelativeSourceMode
Public AncestorType As String
Public AncestorLevel As Long

Private Sub Class_Initialize()
    Mode = RelativeSourceSelf
    AncestorType = ""
    AncestorLevel = 1
End Sub

Public Function InitializeFromMarkup(ByVal MarkupProperties As String) As RelativeSource
    Dim Parts
    Dim P As Variant
    Dim Body
    Dim Key As String
    Dim Val As String

    If Len(Trim$(MarkupProperties)) = 0 Then
        Set InitializeFromMarkup = Me
        Exit Function
    End If

    Parts = Split(MarkupProperties, ",")
    For Each P In Parts
        If InStr(1, P, "=") > 0 Then
            Body = Split(Trim$(P), "=", 2)
            Key = Trim$(Body(0))
            Val = Trim$(Body(1))
            ApplyMarkupPair Key, Val
        Else
            ApplyMarkupPair "Mode", Trim$(P)
        End If
    Next

    Set InitializeFromMarkup = Me
End Function

Private Sub ApplyMarkupPair(ByVal Key As String, ByVal Val As String)
    Select Case LCase$(Key)
        Case "mode", ""
            Mode = ParseMode(Val)
        Case "ancestortype"
            AncestorType = Val
        Case "ancestorlevel"
            If IsNumeric(Val) Then AncestorLevel = CLng(Val)
    End Select
End Sub

Private Function ParseMode(ByVal Val As String) As RelativeSourceMode
    Select Case LCase$(Trim$(Val))
        Case "self", "0", "relativesourceself"
            ParseMode = RelativeSourceSelf
        Case "templatedparent", "1", "relativesourcetemplatedparent"
            ParseMode = RelativeSourceTemplatedParent
        Case "findancestor", "2", "relativesourcefindancestor"
            ParseMode = RelativeSourceFindAncestor
        Case Else
            ParseMode = RelativeSourceSelf
    End Select
End Function

Private Function IMarkupExtension_ProvideValue() As Variant
    ' Descriptor - BindingsManager resolves against the binding Target.
    Set IMarkupExtension_ProvideValue = Me
End Function

' Resolve this RelativeSource against a binding Target (Phase0 / imperative use).
Public Function Resolve(ByVal Target As Object) As Object
    Set Resolve = ResolveRelativeSource(Target, Me)
End Function
