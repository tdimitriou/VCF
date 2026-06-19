Attribute VB_Name = "modStyleWriter"
'Public Sub EnumerateProperties(Obj As Object)
'    On Error Resume Next
'
'    Static Objs As cSortedDictionary
'    If Objs Is Nothing Then Set Objs = New_c.SortedDictionary
'
'    If Objs.Exists(Typename(Obj)) Then Exit Sub
'
'    Objs.Add Typename(Obj)
'
'    Dim Props As cSortedDictionary
'    Set Props = New_c.SortedDictionary
'
'    Dim Setter As String
'    Dim Prefix As String
'
'    If TypeOf Obj Is IDependencyObject Then
'        Dim Dep As IDependencyObject
'        Set Dep = Obj
'
'        Dim DepP As DependencyProperty
'        For Each DepP In Dep.DependencyProperties.RegisteredProperties
'
'            'Prefix = "DependencyProperty."
'            'Setter = ""
'            'Setter = GetXMLNode(Prefix & DepP.Name, DepP.GetValue)
'            'If Len(Setter) Then SB.AppendNL Setter
'
'            If Not Props.Exists(DepP.Name) Then Props.Add DepP.Name, DepP
'        Next
'    End If
'
'    Dim Properties As cProperties
'    Set Properties = New_c.Properties
'
'
'    Properties.BindTo Obj, True
'
'    Dim Prop As cProperty
'    For Each Prop In Properties
'
'        'Prefix = "Property."
'        'Setter = ""
'        'Setter = GetXMLNode(Prefix & Prop.Name, Prop.Value)
'        'If Len(Setter) Then SB.AppendNL Setter
'
'        If Not Props.Exists(Prop.Name) Then Props.Add Prop.Name, Prop
'    Next
'
'    Set Properties = New_c.Properties
'    Properties.BindTo Obj.Widget, True
'
'    For Each Prop In Properties
'
'        'Prefix = "WidgetProperty."
'        'Setter = ""
'        'Setter = GetXMLNode(Prefix & Prop.Name & "." & "Widget", Prop.Value)
'        'If Len(Setter) Then SB.AppendNL Setter
'
'        If Not Props.Exists(Prop.Name) Then Props.Add Prop.Name, Prop
'
'    Next
'
'    If Err Then Err.Clear
'
'    ProcessProps Typename(Obj), Props
'End Sub
'
'Private Sub ProcessProps(ObjType As String, Props As cSortedDictionary)
'    On Error Resume Next
'
'    Dim Prop
'    Dim SB As cStringBuilder
'    Set SB = New_c.StringBuilder
'
'    Dim Name As String
'    Dim Value As String
'    Dim P1 As DependencyProperty
'    Dim P2 As cProperty
'
'    Dim Node As String
'
'    For Each Prop In Props
'        If Err Then Err.Clear
'
'        If TypeOf Prop Is DependencyProperty Then
'            Set P1 = Prop
'            Name = P1.Name
'            Value = P1.GetValue
'        Else
'            Set P2 = Prop
'            Name = P2.Name
'            Value = P2.Value
'        End If
'
'        If Err Then
'            Err.Clear
'            Node = ""
'        Else
'            Node = GetXMLNode(Name, Value)
'        End If
'
'        If Len(Node) Then SB.AppendNL Node
'    Next
'
'    OutputProperties ObjType, SB.ToString
'End Sub
'
'Private Function GetXMLNode(ByVal Name As String, ByVal Value As String) As String
'    On Error Resume Next
'    If Len(Value) = 0 Then Exit Function
'    Value = Replace$(Value, "True", 1)
'    Value = Replace$(Value, "False", 0)
'
'    GetXMLNode = vbTab & "<Setter Name=""" & Name & """ Value=""" & Value & """/>"
'End Function
'
'Private Sub OutputProperties(ObjType As String, Properties As String)
'    If Not New_c.FSO.FolderExists(App.Path & "\Styles") Then New_c.FSO.CreateDirectory (App.Path & "\Styles")
'    With New_c.StringBuilder
'        .AppendNL "<Style TargetType=""" & ObjType & """>"
'
'        .Append Properties
'
'        .AppendNL "</Style>"
'
'        New_c.FSO.WriteTextContent App.Path & "\Styles\" & ObjType & ".xml", .ToString
'    End With
'End Sub
'
