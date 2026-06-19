VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLStyleReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CSEH: ErrMsgBox
Public Sub LoadStyle(ByVal ResourceDictionary As ObservableDictionary, Node As cElement)
    If Node.BaseName <> "Style" Then Exit Sub
    
    Dim NewStyle As Style
    Set NewStyle = CreateStyle(Node)
    
    Dim Style As Style
    
    If ResourceDictionary.ContainsKey(NewStyle.Key) Then
        Set Style = ResourceDictionary(NewStyle.Key)
        
        Dim e
        Dim Setter As KeyValuePair
        
        ' Overwrite the old style's setters
        Style.Clear
        For Each e In NewStyle
            Setter = e
            Style.SetSetter Setter.Key, Setter.Value
        Next
    Else
        ResourceDictionary.Add NewStyle.Key, NewStyle
    End If
    
End Sub

Private Function CreateStyle(Node As cElement) As VCF.Style
    On Error Resume Next
    
    Dim TargetType As String
    
    TargetType = Node.Attributes("TargetType").Text
    
    If Err Then
        'No TargetType defined -> Abort
        Err.Clear
        Exit Function
    End If
    
    Dim Key As String
    
    Dim Attr As cAttribute
    For Each Attr In Node.Attributes
        If Attr.Name = "x:Key" Then
            Key = Attr.Text
            Exit For
        End If
    Next
    
    Dim BasedOn As Style
    For Each Attr In Node.Attributes
        If Attr.Name = "BasedOn" Then
            With New MarkupExtensions
                API.CopyVariable .GetMarkupValue(Attr.Text), BasedOn
            End With
            If BasedOn Is Nothing Then Set BasedOn = CustomConstructor.CreateInstance(Attr.Text)
            Exit For
        End If
    Next
    
    Set CreateStyle = modConstructors.NewStyle(TargetType, Key, BasedOn)
    
    If Not Node.HasChildElements Then Exit Function
    
    Dim Setter As cElement
    For Each Setter In Node.ChildElements
        If Setter.BaseName = "Setter" Then CreateStyle.SetSetter Setter.Attributes("Property").Text, Setter.Attributes("Value").Text
    Next
    If Err Then Err.Clear
End Function

