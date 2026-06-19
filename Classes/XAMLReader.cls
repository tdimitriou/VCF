VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLReader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Load(XML As String) As Object
    Dim Dom As cSimpleDOM
    Set Dom = New_c.SimpleDOM
    Dom.XML = XML

    If Not Dom.WellFormed Then
        If modConstructors.StrictXamlLoad Then
            RaiseXamlLoad "Malformed XAML document."
        End If
        Exit Function
    End If

    Dim CustomClassname As String
    CustomClassname = GetCustomClass(Dom.Root)

    If Len(CustomClassname) > 0 Then
        Set Load = NewCustomObject(CustomClassname)

        If Not Load Is Nothing Then Exit Function

        If modConstructors.StrictXamlLoad Then
            RaiseXamlLoad "x:Class '" & CustomClassname & "' could not be created.", Dom.Root.BaseName
        End If
    End If

    Set Load = NewObject(Dom.Root)

    If Load Is Nothing And modConstructors.StrictXamlLoad Then
        RaiseXamlLoad "Failed to create root element '" & ElementTagName(Dom.Root) & "'.", ElementTagName(Dom.Root)
    End If
End Function

Public Sub LoadSuperclassData(Superclass As Object, XML As String)
    Dim IsValidSuperclass As Boolean
    Dim Attr As cAttribute
        
    With New_c.SimpleDOM
        .XML = XML
        
        If Not .WellFormed Then Exit Sub
        
        'Allow only specified Object Types to be inherited
        If Not (.Root.BaseName = "Application" Or _
                .Root.BaseName = "Window" Or _
                .Root.BaseName = "UserControl") Then Exit Sub
            
        For Each Attr In .Root.Attributes
            ' Make sure that the Object Passed is if Valid Type
            If Attr.Prefix = "x" And Attr.BaseName = "Class" Then
                If Attr.Text = TypeName(Superclass) Then
                    IsValidSuperclass = True
                    Exit For
                End If
            End If
        Next
        
        If Not IsValidSuperclass Then
            Debug.Print "Invalid Type: " & TypeName(Superclass)
            Exit Sub
        End If
        
        SetObjData Superclass, .Root
    End With
        
End Sub

'CSEH: ErrMsgBox
Public Sub LoadApp(Superclass As Object, XML As String)
        '<EhHeader>
        On Error GoTo LoadApp_Err
        '</EhHeader>
        Dim IsValidSuperclass As Boolean
        Dim Attr As cAttribute
    
100     With New_c.SimpleDOM
102         .XML = XML
        
104         If Not .WellFormed Then Exit Sub
        
            'Allow only specified Object Types to be inherited
106         If Not (.Root.BaseName = "Application") Then Exit Sub
            
108         For Each Attr In .Root.Attributes
                ' Make sure that the Object Passed is if Valid Type
110             If Attr.Prefix = "x" And Attr.BaseName = "Class" Then
112                 If Attr.Text = TypeName(Superclass) Then
114                     IsValidSuperclass = True
                        Exit For
                    End If
                End If
            Next
        
116         If Not IsValidSuperclass Then Exit Sub
        
118         SetAppProperties Superclass, .Root.Attributes
        
            Dim BaseApp As VCF.Application
120         Set BaseApp = Superclass.Base
            
            If .Root.HasChildElements Then LoadAppData Superclass, .Root
            
122         BaseApp.OnInitialized
        End With
    
        '<EhFooter>
        Exit Sub

LoadApp_Err:
        MsgBox Err.Description & vbCrLf & _
               "in VCF.XAMLReader.LoadApp " & _
               "at line " & Erl, _
               vbExclamation + vbOKOnly, "Application Error"
        Resume Next
        '</EhFooter>
End Sub

Private Sub LoadAppData(ByVal Object As IApplication, Node As cElement)
    If Not Node.HasChildElements Then Exit Sub
    
    Dim Child As cElement
        
    For Each Child In Node.ChildElements
    
        Select Case Child.BaseName
    
            Case "Application.Resources"
                LoadResources Object, Child
            Case Else
                '
    
        End Select
    Next
End Sub

Private Sub LoadResources(ByVal Object As IApplication, Node As cElement)
    If Not Node.HasChildElements Then Exit Sub
    
    Dim Child As cElement
            
    For Each Child In Node.ChildElements
        Select Case Child.BaseName
        
            Case "Style"
            
                Dim Reader As XAMLStyleReader
                Set Reader = New XAMLStyleReader
                Reader.LoadStyle Object.Resources, Child
                
            Case Else
            
                Dim Key As String
                Key = GetResourceKey(Child)
                Dim Data As Variant
                
                If Len(Key) > 0 Then
                    API.CopyVariable GetResource(Child), Data
                    
                    If IsEmpty(Data) Then GoTo SKIP_ITEM
                    If IsNull(Data) Then GoTo SKIP_ITEM
                    If IsNothing(Data) Then GoTo SKIP_ITEM
                    
                    Object.Resources.Add Key, Data
                    
SKIP_ITEM:
                    '
                End If
        End Select
    Next

End Sub

Private Function GetResource(Node As cElement) As Variant
    On Error Resume Next
    
    Select Case LCase$(Node.BaseName)
    
        Case "string":
            GetResource = ""
            GetResource = Node.Attributes("Value").Text
            
        Case Else:
            Set GetResource = modConstructors.NewObject(Node.BaseName)
            If Not GetResource Is Nothing Then
                SetProperties GetResource, Node.Attributes
                SetChildren GetResource, Node
            End If
    End Select
    
    If Err Then Err.Clear
End Function

Private Function GetResourceKey(Node As cElement, Optional xKey As String) As String
    On Error Resume Next
    GetResourceKey = Node.Attributes("Key").Text
    GetResourceKey = Node.Attributes("x:Key").Text
    xKey = Node.Attributes("x:Key").Text
    If Err Then Err.Clear
End Function

Private Function SetChildren(ByVal Parent As Object, Node As cElement)
    If Not Node.HasChildElements Then Exit Function
    
    Dim ChildNode As cElement
    For Each ChildNode In Node.ChildElements
        SetChild Parent, ChildNode
    Next
End Function

Private Function SetChild(Parent As Object, Child As cElement)
Try:
    On Error GoTo Catch
    
    Dim Key As String
    Dim xKey As String
    Key = GetResourceKey(Child, xKey)
    
    Dim Data As Variant
    
    If Len(Key) > 0 Then
        API.CopyVariable GetResource(Child), Data
        
        If IsEmpty(Data) Then Exit Function
        If IsNull(Data) Then Exit Function
        If IsNothing(Data) Then Exit Function
        
        'Object.Resources.Add Key, Data
        
        If Len(xKey) Then
            CallByName Parent, "", VbSet, xKey, Data
        End If
    End If
    
    If Not IsObject(Data) Then Exit Function
    If Child.HasChildElements Then SetChildren Data, Child
    
    Exit Function
    
Catch:
    
    MsgBox Err.Number & Err.Description

End Function

Private Function SetObjData(Object As Object, Node As cElement) As Object
    On Error Resume Next
    
    Err.Clear
        
    SetObjectProperties Object, Node.Attributes
        
    If Node.HasChildElements = False Then Exit Function
    
    Dim Child As cElement
    
    If TypeOf Object Is VCF.IWindow Then
        Dim ObjWin As IWindow
        Set ObjWin = Object

        For Each Child In Node.ChildElements
            ObjWin.Base.Children.Add NewObject(Child)
        Next
        
        If Err Then Err.Clear
        Exit Function
    End If
    
    For Each Child In Node.ChildElements
        Object.Children.Add NewObject(Child)
    Next
    
    If Err Then Err.Clear
    
End Function

Private Function GetCustomClass(Node As cElement) As String
    Dim Attr As cAttribute
    For Each Attr In Node.Attributes
        If Attr.Name = "x:Class" Then
            GetCustomClass = Attr.Text
            Exit Function
        End If
    Next
End Function

Private Function NewCustomObject(Classname As String) As Object

    If CustomConstructor Is Nothing Then Exit Function
    
    Set NewCustomObject = CustomConstructor.CreateInstance(Classname)
End Function

Private Function NewObject(Node As cElement) As Object
    On Error GoTo NewObjectErr

    Dim Namespace As String, Class As String

    Namespace = Node.Prefix
    If Len(Namespace) = 0 Then Namespace = "VCF"

    Class = Node.BaseName

    Set NewObject = CreateInstance(Namespace, Class)

    If NewObject Is Nothing Then
        If modConstructors.StrictXamlLoad Then
            RaiseXamlLoad "Unknown type '" & ElementTagName(Node) & "'.", ElementTagName(Node)
        End If
        Exit Function
    End If
    
    SetObjectProperties NewObject, Node.Attributes
        
    If Node.HasChildElements = False Then Exit Function
    
    Dim Child As cElement
    Dim Prop As Variant
    
    For Each Child In Node.ChildElements
        If InStr(1, Child.BaseName, ".") > 0 Then
            If Split(Child.BaseName, ".")(0) = TypeName(NewObject) Then
                
                Err.Clear
                
                Call API.CopyVariable(CallByName(NewObject, Split(Child.BaseName, ".")(1), VbGet), Prop)
                                                
                If Err Then GoTo SKIP_PROP
                    
                If IsObject(Prop) And Child.HasChildElements Then

                    If Not Prop Is Nothing Then
                        If TypeOf Prop Is ObservableDictionary Or _
                            TypeOf Prop Is ObservableCollection Or _
                            TypeOf Prop Is List Then

                            AddPropertyList Prop, Child.ChildElements
                        
                        Else
                            CallByName NewObject, Split(Child.BaseName, ".")(1), VbSet, NewObject(Child.ChildElements(0))
                        End If
                    Else
                        CallByName NewObject, Split(Child.BaseName, ".")(1), VbSet, NewObject(Child.ChildElements(0))
                    End If

                End If
            
            End If
            
SKIP_PROP:
            If Err Then Err.Clear
        
        Else
            NewObject.Children.Add NewObject(Child)
            If Err Then Err.Clear
        End If
    Next
    Exit Function

NewObjectErr:
    If modConstructors.StrictXamlLoad Then
        RaiseXamlLoad Err.Description, ElementTagName(Node), vbNullString, 0, 0, Err.Number
    End If
    Err.Clear
End Function

Private Function ElementTagName(ByVal Node As cElement) As String
    If Len(Node.Prefix) > 0 Then
        ElementTagName = Node.Prefix & ":" & Node.BaseName
    Else
        ElementTagName = Node.BaseName
    End If
End Function

Private Sub AddPropertyList(ByVal Property As Object, ByVal Elements As cElements)
    Dim Node As cElement
    Dim NewObj As Object
    
    For Each Node In Elements
        Set NewObj = NewObject(Node)
        
        If TypeOf Property Is ObservableDictionary Then
            AddDictionaryItem Property, NewObj
        End If
    Next
    
End Sub

Private Sub AddDictionaryItem(Dict As ObservableDictionary, Item As Object)
    Dim Key

    Key = Null
    
    Key = GetProp(Item, "Key")
    If IsNull(Key) Then Key = GetProp(Item, "Name")
        
    If Not IsNull(Key) Then
        AddDictionaryEntry Dict, Key, Item
    Else
        Key = GetProp(Item, "DataType")
        
        If IsNull(Key) Then Exit Sub
        If Len(Key) = 0 Then Exit Sub
        Key = TypeName(Item) & "_" & Key
        AddDictionaryEntry Dict, Key, Item
    End If
End Sub

'CSEH: ErrMsgBox
Private Function GetProp(Object As Object, PropertyName As String) As Variant
Try:
    On Error GoTo Catch
    
    GetProp = Null
    
    Call API.CopyVariable(CallByName(Object, PropertyName, VbGet), GetProp)

    If VarType(GetProp) = vbString Then
        If Len(GetProp) = 0 Then GetProp = Null
    End If
    
    Exit Function
    
Catch:
    GetProp = Null
End Function

Private Sub AddDictionaryEntry(Dict As ObservableDictionary, Key, Value)
    Dict.Add Key, Value
End Sub

Private Sub SetAppProperties(Obj As Object, Attributes As cAttributes)
    Dim Attr As cAttribute
    
    For Each Attr In Attributes
        SetProperty Obj, Attr.Name, Attr.Text
    Next
End Sub

'CSEH: ErrMsgBox
Private Sub SetObjectProperties(Obj As Object, Attributes As cAttributes)
    
    Dim Dep As IDependencyObject
    
    If TypeOf Obj Is IDependencyObject Then
        Set Dep = Obj
        SetDependencyProperties Dep, Attributes
    Else
        SetProperties Obj, Attributes
    End If
End Sub

Private Sub SetDependencyProperties(ByVal Dep As IDependencyObject, Attributes As cAttributes)
    Dim Attr As cAttribute
    Dim Value As Variant
    Dim Obj As Object
    Dim PropName As String
    
    Set Obj = API.CObj(Dep)
    
    For Each Attr In Attributes
        PropName = MapDesignPropertyAlias(Dep, Attr.Name)

        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(Attr.Text, Dep, PropName), Value
        End With
        
        If InStr(1, PropName, ".") > 0 Then
            SetAttachedProperty Obj, PropName, Value
        Else
            If Dep.DependencyProperties.Exists(PropName) Then
                SetDependencyProperty Dep, PropName, Value
            Else
                SetProperty Obj, Attr.Name, Value
            End If
        End If
    Next
End Sub

Private Sub SetProperties(ByVal Obj As Object, Attributes As cAttributes)
    Dim Attr As cAttribute
    
    For Each Attr In Attributes
        If InStr(1, Attr.Name, ".") > 0 Then
            ' It is an Attached Property
            SetAttachedProperty Obj, Attr.Name, Attr.Text
        Else
            SetProperty Obj, Attr.Name, Attr.Text
        End If
    Next
End Sub

Private Sub SetAttachedProperty(ByVal Obj As IUIElement, ByVal PropertyName As String, Value)
    Dim Dict As ObservableDictionary
    Dim PropDict As String
    Dim PropKey As String
    Dim PropData
            
    PropData = Split(PropertyName, ".")
    PropDict = PropData(0)
    PropKey = PropData(1)
    
    If Obj.AttachedProperties.ContainsKey(PropDict) Then
        Set Dict = Obj.AttachedProperties(PropDict)
    Else
        Set Dict = New ObservableDictionary
        Obj.AttachedProperties.Add PropDict, Dict
    End If
    
    If Not Dict.ContainsKey(PropKey) Then Dict.Add PropKey, Value

End Sub

Private Sub SetDependencyProperty(ByVal Obj As IDependencyObject, ByVal PropertyName As String, Value)
    Dim Prop As DependencyProperty
    
    Set Prop = Obj.DependencyProperties.GetProperty(PropertyName)
    
    With New XAMLDependencyPropertyManager
    
        Obj.DependencyProperties.SetValue PropertyName, .GetPropertyValueFromString(Prop, Value)
        
        'On Error Resume Next
        'Debug.Print "SetDependencyProperty Value: "; Value,
        'Debug.Print "Typename:"; TypeName(.GetPropertyValueFromString(Prop, Value)),
        'If TypeName(.GetPropertyValueFromString(Prop, Value)) = "Style" Then Debug.Print .GetPropertyValueFromString(Prop, Value).Key,
        'Debug.Print
    
    End With

End Sub

Private Sub SetProperty(ByVal Obj As Object, ByVal Name As String, ByVal Value As String, Optional PropertyValue)
    On Error Resume Next
    
    Err.Clear
    
    If LCase$(Name) = "imagekey" Then
        If Not TypeOf Obj Is cWidgetBase Then
            With New XAMLImagePropertyManager
                Value = .LoadImage(Value)
            End With
        End If
    End If
    
    If IsMissing(PropertyValue) Then
        ' Get The Final Value of the property
        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(Value), PropertyValue
        End With
    End If
    
    If IsObject(PropertyValue) Then
        CallByName Obj, Name, VbSet, PropertyValue
    Else
        CallByName Obj, Name, VbLet, PropertyValue
    End If
    
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    
    ' Pass the final property value, (PropertyValue) directly to the inner object
    ' to avoid the extra call to the MarkupExtensions.GetMarkupValue
    
    If TypeOf Obj Is IWindow Then
        Dim ObjWin As IWindow
        Set ObjWin = Obj
        
        SetProperty ObjWin.Base, Name, Value, PropertyValue

        If Err.Number = 0 Then Exit Sub
        Err.Clear
        SetProperty ObjWin.Base.Form, Name, Value, PropertyValue

        If Err.Number = 0 Then Exit Sub
        Err.Clear
        SetProperty ObjWin.Base.Form.WidgetRoot, Name, Value, PropertyValue
    Else
        SetProperty Obj.Widget, Name, Value, PropertyValue
    End If
End Sub
