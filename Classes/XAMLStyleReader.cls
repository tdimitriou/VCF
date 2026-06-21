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
Public Sub LoadStyle(ByVal Resources As ResourceDictionary, Node As cElement)
    If Node.BaseName <> "Style" Then Exit Sub
    
    Dim NewStyle As Style
    Set NewStyle = CreateStyle(Node)
    
    Dim Style As Style
    
    If Resources.ContainsKey(NewStyle.Key) Then
        Set Style = Resources(NewStyle.Key)
        
        Dim e
        Dim Setter As KeyValuePair
        Dim i As Long
        
        Style.Clear
        Style.ClearTriggers
        Set Style.Template = Nothing
        For Each e In NewStyle
            Setter = e
            Style.SetSetter Setter.Key, Setter.Value
        Next
        For i = 0 To NewStyle.TriggerCount - 1
            Style.AddTrigger NewStyle.TriggerAt(i).Clone
        Next
        If Not NewStyle.Template Is Nothing Then Set Style.Template = NewStyle.Template.Clone
    Else
        Resources.Add NewStyle.Key, NewStyle
    End If
    
End Sub

Private Function CreateStyle(Node As cElement) As VCF.Style
    On Error Resume Next
    
    Dim TargetType As String
    
    TargetType = Node.Attributes("TargetType").Text
    
    If Err Then
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
    
    Dim Child As cElement
    For Each Child In Node.ChildElements
        Select Case Child.BaseName
            Case "Setter"
                If IsTemplateSetter(Child) Then
                    Set CreateStyle.Template = LoadControlTemplateNode(FindControlTemplateNode(Child))
                Else
                    CreateStyle.SetSetter GetSetterPropertyName(Child), GetAttributeText(Child, "Value")
                End If
            Case "Style.Triggers"
                LoadTriggers CreateStyle, Child
            Case "Trigger"
                CreateStyle.AddTrigger CreatePropertyTrigger(Child)
        End Select
    Next
    If Err Then Err.Clear
End Function

Private Sub LoadTriggers(ByVal Style As Style, Node As cElement)
    Dim Child As cElement

    For Each Child In Node.ChildElements
        If Child.BaseName = "Trigger" Then Style.AddTrigger CreatePropertyTrigger(Child)
    Next
End Sub

Private Function CreatePropertyTrigger(Node As cElement) As PropertyTrigger
    Dim Trig As PropertyTrigger
    Dim Setter As cElement

    Set Trig = modConstructors.NewPropertyTrigger(GetTriggerPropertyName(Node), GetAttributeText(Node, "Value"))

    For Each Setter In Node.ChildElements
        If Setter.BaseName = "Setter" Then
            Trig.SetSetter GetSetterPropertyName(Setter), GetAttributeText(Setter, "Value")
        End If
    Next

    Set CreatePropertyTrigger = Trig
End Function

Private Function GetSetterPropertyName(ByVal Node As cElement) As String
    GetSetterPropertyName = GetAttributeText(Node, "Property")
    If Len(GetSetterPropertyName) = 0 Then GetSetterPropertyName = GetAttributeText(Node, "Name")
End Function

Private Function GetTriggerPropertyName(ByVal Node As cElement) As String
    GetTriggerPropertyName = GetAttributeText(Node, "Property")
    If Len(GetTriggerPropertyName) = 0 Then GetTriggerPropertyName = GetAttributeText(Node, "Name")
End Function

Private Function GetAttributeText(ByVal Node As cElement, ByVal Name As String) As String
    On Error Resume Next
    GetAttributeText = Node.Attributes(Name).Text
    If Err Then Err.Clear
End Function

Private Function IsTemplateSetter(ByVal SetterNode As cElement) As Boolean
    IsTemplateSetter = (LCase$(GetSetterPropertyName(SetterNode)) = "template")
End Function

Private Function FindControlTemplateNode(ByVal SetterNode As cElement) As cElement
    Dim Child As cElement
    Dim Inner As cElement

    For Each Child In SetterNode.ChildElements
        If Child.BaseName = "ControlTemplate" Then
            Set FindControlTemplateNode = Child
            Exit Function
        End If
        If Child.BaseName = "Setter.Value" Then
            For Each Inner In Child.ChildElements
                If Inner.BaseName = "ControlTemplate" Then
                    Set FindControlTemplateNode = Inner
                    Exit Function
                End If
            Next
        End If
    Next
End Function

Private Function LoadControlTemplateNode(ByVal Node As cElement) As ControlTemplate
    Dim Reader As XAMLReader
    Dim Obj As Object

    If Node Is Nothing Then Exit Function

    Set Reader = New XAMLReader
    Set Obj = Reader.LoadElement(Node)
    If TypeOf Obj Is ControlTemplate Then Set LoadControlTemplateNode = Obj
End Function
