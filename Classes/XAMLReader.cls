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
    With New_c.SimpleDOM
        .XML = XML
        
        If Not .WellFormed Then Exit Function
        
        Dim CustomClassname As String
        CustomClassname = GetCustomClass(.Root)
        
        If Len(CustomClassname) > 0 Then
        
            ' If a custom class is provided through the 'x:Class' Attribute the
            ' loader tries to create a 'NewCustomObject' of the specified
            ' 'CustomClassname', provided that a CustomConstructor exists.
            Set Load = NewCustomObject(CustomClassname)
            
            ' Return the newly created object and exit if the 'NewCustomObject'
            ' function succeeded. Note that we just create the object defined in
            ' the Root element of the XAML document based on the 'x:Class' Attribute
            ' (normally a 'UserControl' or a custom 'Window' class) and ignore the
            ' rest of the document. The Object should handle the rest during the
            ' InitializeComponent Event and load its XAML Contents using the
            ' 'LoadSuperclassData' method provided by the current class (XAMLReader)
            
            ' Even if the function failed (returned Nothing) we give it another
            ' try using the 'NewObject' function.
            ' (The second try should always fail in strict terms because of the
            ' existence of the x:Class Attribute. However the loader will be allowed
            ' to succeed if the TagName of the Root element defines a public class which
            ' can be created with the 'VB.CreateObject(Namespace.Classname)' function.
            ' In that case the 'x:Class' Attribute will be completely ignored)
            If Not Load Is Nothing Then Exit Function
        End If
        
        Set Load = NewObject(.Root)
    End With
    
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
    '
End Function

Private Function GetResourceKey(Node As cElement) As String
    On Error Resume Next
    GetResourceKey = Node.Attributes("Key").Text
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

        Exit Function
    End If
    
    For Each Child In Node.ChildElements
        Object.Children.Add NewObject(Child)
    Next
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
    On Error Resume Next
    
    Err.Clear
    
    Dim Namespace As String, Class As String
    
    Namespace = Node.Prefix
    If Len(Namespace) = 0 Then Namespace = "VCF"
    
    Class = Node.BaseName

    Set NewObject = CreateInstance(Namespace, Class) 'CreateObject(Namespace & "." & Class)
    
    If NewObject Is Nothing Then Exit Function
    
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
            If Err Then
                Err.Clear
            End If
        Else
            NewObject.Children.Add NewObject(Child)
        End If
    Next
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
'    With New StyleManager
'        .ApplyStyle Obj
'    End With
    
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
    
    Set Obj = API.CObj(Dep)
    
    For Each Attr In Attributes
        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(Attr.Text, Dep, Attr.Name), Value
        End With
    
        If InStr(1, Attr.Name, ".") > 0 Then
            ' It is an Attached Property
            SetAttachedProperty Obj, Attr.Name, Value
        Else
            If Dep.DependencyProperties.Exists(Attr.Name) Then
                SetDependencyProperty Dep, Attr.Name, Value
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
    
    End With

End Sub

Private Sub SetProperty(ByVal Obj As Object, ByVal Name As String, ByVal Value As String)
    On Error Resume Next
    
    Err.Clear
    
    If LCase$(Name) = "imagekey" Then
        If Not TypeOf Obj Is cWidgetBase Then
            With New XAMLImagePropertyManager
                Value = .LoadImage(Value)
            End With
        End If
    End If
    
    CallByName Obj, Name, VbLet, Value
    
    If Err.Number = 0 Then Exit Sub
    Err.Clear
    
    If TypeOf Obj Is IWindow Then
        Dim ObjWin As IWindow
        Set ObjWin = Obj
        
        SetProperty ObjWin.Base, Name, Value

        If Err.Number = 0 Then Exit Sub
        Err.Clear
        SetProperty ObjWin.Base.Form, Name, Value

        If Err.Number = 0 Then Exit Sub
        Err.Clear
        SetProperty ObjWin.Base.Form.WidgetRoot, Name, Value
    Else
        SetProperty Obj.Widget, Name, Value
    End If
End Sub


