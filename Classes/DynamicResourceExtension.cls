VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DynamicResourceExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMarkupExtension

Public ResourceKey As String
Public Key As String
Public Target As IUIElement

Private Sub Class_Initialize()
    ResourceKey = ""
    Key = ""
End Sub

Private Function IMarkupExtension_ProvideValue() As Variant
    Call API.CopyVariable(ProvideValue, IMarkupExtension_ProvideValue)
End Function

Public Function ProvideValue() As Variant
    Dim LookupKey As String
    Dim Found As Variant

    LookupKey = ResourceKey
    If Len(LookupKey) = 0 Then LookupKey = Key
    If Len(LookupKey) = 0 Then Exit Function

    If Not Target Is Nothing Then
        API.CopyVariable Target.Base.TryFindResource(LookupKey), Found
        If Not IsEmpty(Found) Then
            API.CopyVariable Found, ProvideValue
            Exit Function
        End If
    ElseIf Not Application.Current Is Nothing Then
        API.CopyVariable Application.Current.TryFindResource(LookupKey), Found
        If Not IsEmpty(Found) Then
            API.CopyVariable Found, ProvideValue
            Exit Function
        End If
    End If

    API.CopyVariable GetThemeResource(LookupKey), ProvideValue
End Function

Private Function GetThemeResource(ByVal LookupKey As String)
    On Error GoTo Catch

    Dim ActiveTheme As ObservableDictionary
    Set ActiveTheme = ThemesManager.ActiveTheme

    If ActiveTheme Is Nothing Then Exit Function

    API.CopyVariable ActiveTheme(LookupKey), GetThemeResource
    Exit Function

Catch:
    GetThemeResource = Empty
End Function

Private Property Get ThemesManager() As ThemesManager
    On Error GoTo Catch
    Set ThemesManager = Application.Current.Base.ThemesManager
    Exit Property
Catch:
    Set ThemesManager = Nothing
End Property

Public Function InitializeFromMarkup(ByVal Target As IDependencyObject, _
                                            ByVal MarkupProperties As String) As IMarkupExtension

    Dim ArgList
    Dim PropList
    Dim vProp As Variant
    Dim Prop As KeyValuePair
    Dim Properties As cSortedDictionary

    If ParseMarkupPropertiesString(MarkupProperties, ArgList, PropList) <> 0 Then Exit Function
    If IsEmpty(ArgList) And IsEmpty(PropList) Then Exit Function

    If Not Target Is Nothing Then
        If TypeOf Target Is IUIElement Then
            Set Me.Target = Target
        ElseIf TypeOf Target Is IWindow Then
            Dim TrgWin As IWindow
            Set TrgWin = Target
            Set Me.Target = TrgWin.Base
        ElseIf TypeOf Target Is IUserControl Then
            Dim TrgUC As IUserControl
            Set TrgUC = Target
            Set Me.Target = TrgUC.Base
        End If
    End If

    If Not IsEmpty(ArgList) Then
        Me.ResourceKey = ArgList(0)
    End If

    Set Properties = New_c.SortedDictionary(TextCompare)

    If Not IsEmpty(PropList) Then
        For Each vProp In PropList
            Prop = vProp
            Properties.Add Prop.Key, Prop.Value
        Next
    End If

    SetProperties Me, Properties

    If Len(Me.ResourceKey) = 0 And Len(Me.Key) > 0 Then
        Me.ResourceKey = Me.Key
    End If

    Set InitializeFromMarkup = Me
End Function

Private Function ParseMarkupPropertiesString(ByVal MarkupProperties As String, _
                                            ByRef ArgList As Variant, _
                                            ByRef PropList As Variant) As Long

    On Error GoTo Catch

    Dim Props
    Dim Prop
    Dim Key As String
    Dim Value
    Dim Body
    Dim Pair As KeyValuePair
    Dim Literals
    Dim Ext As MarkupExtensions

    Set Ext = New MarkupExtensions
    Ext.ParseLiterals MarkupProperties, Literals
    Props = Split(MarkupProperties, ",")

    For Each Prop In Props
        Ext.RestoreLiterals Prop, Literals
        Body = Split(Trim$(Prop), "=")

        If UBound(Body) = 0 Then
            If IsEmpty(ArgList) Then
                ReDim ArgList(0 To 0) As Variant
            Else
                ReDim Preserve ArgList(0 To UBound(ArgList) + 1) As Variant
            End If
            API.CopyVariable Ext.GetMarkupValue(Body(0), Me.Target), ArgList(UBound(ArgList))
        ElseIf UBound(Body) = 1 Then
            If IsEmpty(PropList) Then
                ReDim PropList(0 To 0) As Variant
            Else
                ReDim Preserve PropList(0 To UBound(PropList) + 1) As Variant
            End If
            Key = Trim$(Body(0))
            Value = Trim$(Body(1))
            Pair.Key = Key
            API.CopyVariable Ext.GetMarkupValue(Value, Me.Target), Pair.Value
            PropList(UBound(PropList)) = Pair
        End If
    Next

    ParseMarkupPropertiesString = 0
    Exit Function

Catch:
    On Error Resume Next
    ArgList = Empty
    PropList = Empty
    ParseMarkupPropertiesString = Err.Number
End Function

Private Sub SetProperties(ByVal Ext As IMarkupExtension, ByVal Properties As cSortedDictionary)
    Dim Index As Long
    Dim Key
    Dim Value

    For Index = 0 To Properties.Count - 1
        Key = Properties.KeyByIndex(Index)
        API.CopyVariable Properties.ItemByIndex(Index), Value
        If IsObject(Value) Then
            CallByName API.CUnk(Ext), Key, VbSet, Value
        Else
            CallByName API.CUnk(Ext), Key, VbLet, Value
        End If
    Next
End Sub
