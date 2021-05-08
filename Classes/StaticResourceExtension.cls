VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "StaticResourceExtension"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMarkupExtension

Public ResourceKey As String
Public Target As IUIElement

Private Sub Class_Initialize()
    ResourceKey = ""
End Sub

Private Function IMarkupExtension_ProvideValue() As Variant
    Call API.CopyVariable(ProvideValue, IMarkupExtension_ProvideValue)
End Function

Public Function ProvideValue() As Variant
    API.CopyVariable Target.Base.TryFindResource(ResourceKey), ProvideValue
End Function

'==============================================================================
Public Function InitializeFromMarkup(ByVal Target As IDependencyObject, _
                                            ByVal MarkupProperties As String) As IMarkupExtension

    Dim ArgList
    Dim PropList
    
    Dim Arg As Variant
    Dim vProp As Variant
    Dim Prop As KeyValuePair
    Dim Properties As cSortedDictionary
            
    If ParseMarkupPropertiesString(MarkupProperties, ArgList, PropList) <> 0 Then Exit Function
    If IsEmpty(ArgList) And IsEmpty(PropList) Then Exit Function
    
    Set Me.Target = Target
    
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
    
    Set InitializeFromMarkup = Me
End Function


Private Function ParseMarkupPropertiesString(ByVal MarkupProperties As String, _
                                            ByRef ArgList As Variant, _
                                            ByRef PropList As Variant) As Long

Try:
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
            'It is an Argument
            
            If IsEmpty(ArgList) Then
                ReDim ArgList(0 To 0) As Variant
            Else
                ReDim Preserve ArgList(0 To UBound(ArgList) + 1) As Variant
            End If
                        
            API.CopyVariable Ext.GetMarkupValue(Body(0), Me.Target), ArgList(UBound(ArgList))
            
        ElseIf UBound(Body) = 1 Then
            'It is a Property
            
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
    Dim Prop
    Dim Index As Long
    Dim Key
    Dim Value
    
    For Index = 0 To Properties.Count - 1
        
        Key = Properties.KeyByIndex(Index)
        API.CopyVariable Properties.ItemByIndex(Index), Value
        
        'With New MarkupExtensions
        '    API.CopyVariable .GetMarkupValue(Value), Value
        'End With
        
        If IsObject(Value) Then
            CallByName API.CUnk(Ext), Key, VbSet, Value
        Else
            CallByName API.CUnk(Ext), Key, VbLet, Value
        End If
    Next
End Sub

