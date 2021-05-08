VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "MarkupExtensions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Function GetExtensionValue(ByVal MarkupText As String, _
                        Optional ByVal Target As IDependencyObject = Nothing, _
                        Optional ByVal TargetProperty As String = "", _
                        Optional ByRef IsValidExtension As Boolean) As Variant
    
    On Error Resume Next
    
    Dim Position As Long

    MarkupText = Left$(MarkupText, Len(MarkupText) - 1)
    MarkupText = Right$(MarkupText, Len(MarkupText) - 1)
    MarkupText = Trim$(MarkupText)
    
    Position = InStr(1, MarkupText, " ")
    
    Dim Classname As String
    If Position > 0 Then Classname = Mid$(MarkupText, 1, Position - 1)
        
    Dim MarkupProperties As String
    MarkupProperties = Right$(MarkupText, Len(MarkupText) - Len(Classname))
    MarkupProperties = Trim$(MarkupProperties)
    
    Dim ExtObject As Object
    Set ExtObject = CreateExtObject(Classname, Target, TargetProperty, MarkupProperties)
    
    'MsgBox "ExtType: " & Classname & ":" & TypeName(ExtObject)
    
    If ExtObject Is Nothing Then Exit Function
    If Not TypeOf ExtObject Is IMarkupExtension Then Exit Function
        
    IsValidExtension = True
    'MsgBox "Still OK"
    Dim Ext As IMarkupExtension
    Set Ext = ExtObject
    
    'MsgBox "Still2 OK"
        
    Call API.CopyVariable(Ext.ProvideValue, GetExtensionValue)
    'MsgBox "Still3 OK"

    'MsgBox "Classname = '" & "': Value=" & TypeName(GetExtensionValue)
End Function

Private Function CreateExtObject(ByVal Classname As String, _
                                ByVal Target As IDependencyObject, _
                                ByVal TargetProperty As String, _
                                ByVal MarkupProperties As String) As Object
                                
    Select Case Classname
    
        Case "Binding":
            
            Set CreateExtObject = BindingsManager.CreateBindingFromMarkup(Target, TargetProperty, MarkupProperties)
        
        Case "StaticResource"
            Dim StaticExt As StaticResourceExtension
            
            Set StaticExt = New StaticResourceExtension
            Set CreateExtObject = StaticExt.InitializeFromMarkup(Target, MarkupProperties)
        
        Case Else:
        
            Set CreateExtObject = modConstructors.NewCustomObject(Classname)
            
            If CreateExtObject Is Nothing Then Exit Function
            If Not TypeOf CreateExtObject Is IMarkupExtension Then Exit Function
            
            SetProperties CreateExtObject, MarkupProperties
    
    End Select
End Function

Private Sub SetProperties(ByVal Ext As IMarkupExtension, ByVal MarkupProperties As String)
    Dim Props
    Dim Prop
    Dim Key As String
    Dim Value
    Dim Literals
    
    ParseLiterals MarkupProperties, Literals
    
    Props = Split(MarkupProperties, ",")
    
    For Each Prop In Props
        
        RestoreLiterals Prop, Literals
        
        Key = Trim$(Split(Trim$(Prop), "=")(0))
        Value = Trim$(Split(Trim$(Prop), "=")(1))
        
        API.CopyVariable GetMarkupValue(Value), Value
        
        If IsObject(Value) Then
            CallByName API.CUnk(Ext), Key, VbSet, Value
        Else
            CallByName API.CUnk(Ext), Key, VbLet, Value
        End If
    Next
End Sub

Public Sub RestoreLiterals(ByRef Text, ByRef Literals As Variant)
    Dim Index As Long
    
    If IsEmpty(Literals) Then Exit Sub
    
    For Index = 0 To UBound(Literals)
        Text = Replace$(Text, "{" & Index & "}", Literals(Index))
    Next
End Sub

Public Sub ParseLiterals(ByRef MarkupText As String, ByRef Literals As Variant)
    Dim Index As Long
    
    ParseLiteralsProc MarkupText, Literals
    
    If IsEmpty(Literals) Then Exit Sub
    
    For Index = 0 To UBound(Literals)
        MarkupText = Replace$(MarkupText, Chr(1) & Index & Chr(1), "{" & Index & "}")
    Next
End Sub

Private Sub ParseLiteralsProc(ByRef MarkupText As String, ByRef Literals As Variant)
    Dim StartPos As Long
    Dim EndPos As Long
    Dim Literal As String
    
    StartPos = InStr(1, MarkupText, "{")
    
    If StartPos = 0 Then Exit Sub
    
    If IsEmpty(Literals) Then
        ReDim Literals(0 To 0)
    Else
        ReDim Preserve Literals(0 To UBound(Literals) + 1)
    End If
    
    EndPos = InStrRev(MarkupText, "}") + 1
    
    Literal = Mid$(MarkupText, StartPos, EndPos - StartPos)
        
    MarkupText = Replace$(MarkupText, Literal, Chr(1) & UBound(Literals) & Chr(1))
        
    'Literal = Left$(Literal, Len(Literal) - 1)
    'Literal = Right$(Literal, Len(Literal) - 1)

    Literals(UBound(Literals)) = Literal
    
    'ParseLiteralsProc MarkupText, Literals
End Sub

Public Function GetMarkupValue(ByVal MarkupText As String, _
                                Optional ByVal Target As IDependencyObject = Nothing, _
                                Optional ByVal TargetProperty As String = "") As Variant
    
    Dim IsValidExtension As Boolean
    
    MarkupText = Trim$(MarkupText)
    
    ' No {} Brackets defined. Return the exact string
    If Not (Left$(MarkupText, 1) = "{" And Right$(MarkupText, 1) = "}") Then
        GetMarkupValue = MarkupText
        Exit Function
    End If
    
    ' Try to get the Extension (if any) Value
    API.CopyVariable GetExtensionValue(MarkupText, Target, TargetProperty, IsValidExtension), GetMarkupValue
        
    If Not IsEmpty(GetMarkupValue) Then Exit Function
    If IsValidExtension Then Exit Function
    
    ' {} Barckets were defined but no extension was matched
    ' Remove the brackets and return the string
    MarkupText = Left$(MarkupText, Len(MarkupText) - 1)
    MarkupText = Right$(MarkupText, Len(MarkupText) - 1)

    GetMarkupValue = MarkupText
End Function
