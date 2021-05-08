VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NestedProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents ChildNotifier    As PropertyChangedEvent
Attribute ChildNotifier.VB_VarHelpID = -1

Private m_ParentPtr                 As Long
Private m_SourcePtr                 As Long

Private m_Path                      As String
Private m_Name                      As String
Private m_Child                     As NestedProperty
Attribute m_Child.VB_VarHelpID = -1

Private PathList                    As Variant

Public Event ValueChanged(ByVal Sender As Object, e As Variant)

Public Property Get Path() As String
    Path = m_Path
End Property

Private Property Let Path(ByVal Value As String)
    If Value = m_Path Then Exit Property
    
    m_Path = Value
        
    PathList = Empty
    
    If Len(Value) = 0 Then Exit Property
    
    PathList = Split(m_Path, ".")
End Property

Public Property Get Source() As Object
    Set Source = API.ObjFromPtr(m_SourcePtr)
End Property

Public Property Set Source(ByVal Value As Object)
    
    m_SourcePtr = ObjPtr(Value)
    
    If Child Is Nothing Then Exit Property
    
    SetupChildSource
End Property

Public Property Get Name() As String
    Name = m_Name
End Property

Private Property Let Name(ByVal Value As String)
    m_Name = Value
End Property

Public Property Get Child() As NestedProperty
    Set Child = m_Child
End Property

Private Property Set Child(ByVal Value As NestedProperty)
    Set m_Child = Value
End Property

Public Property Get Parent() As Object
    Set Parent = API.ObjFromPtr(m_ParentPtr)
End Property

Private Property Set Parent(ByVal Value As Object)
    m_ParentPtr = ObjPtr(Value)
End Property

'CSEH: ErrMsgBox
Public Sub OnPropertyChanged(ByVal PropertyName As String)
    ' The Parent (either the Binding Object itself or the Parent
    ' NestedProperty) of the current object has notified that the
    ' PropertyChangedEvent has been fired.
    
    If PropertyName = Name Then
        If Not Child Is Nothing Then SetupChildSource
        OnValueChanged Me, PropertyName
    Else
        ' Pass to the child (if any)
        If Not m_Child Is Nothing Then m_Child.OnPropertyChanged PropertyName
    End If
End Sub

Private Sub ChildNotifier_PropertyChanged(ByVal Sender As Object, e As Variant)
    ' The object holds the PropertyChangedEvent for its Child Object.
    ' It must notify the Child if the event has been fired for it (the child).
    
    ' The Child Null Check is not necessary because the 'ChildNotifier' is set
    ' only when the child is set and the Type of the Property is INotifyPropertyChanged
    
    'Debug.Print "Child Is Nothing: " & (Child Is Nothing) & ", PropertyName = " & e
    If Child Is Nothing Then Exit Sub
    If e <> Child.Name Then Exit Sub

    Child.OnPropertyChanged e
End Sub

'CSEH: ErrMsgBox
Public Function GetValue() As Variant
    If Me.Source Is Nothing Then Exit Function
    
    ' Aka : (Set) GetValue = Me.Source.Name
    Call API.CopyVariable(CallByName(Me.Source, Name, VbGet), GetValue)
    
    ' If the Object has no child we just return the result we  got so far.
    If Child Is Nothing Then Exit Function
    
    ' If the property has a child property we return the result of the Child's
    ' GetValue. This will occur recursively until the Child (Nested Property)
    ' is the final Object in the hierarchy (does not have a child).
    
    Call API.CopyVariable(Child.GetValue, GetValue)
End Function

'CSEH: ErrMsgBox
Public Sub SetValue(Value)
    If Me.Source Is Nothing Then Exit Sub
    
    ' Allow the function to be called for the top level property only
    If TypeOf Me.Source Is NestedProperty Then
        Err.Raise 5 ' Invalid procedure call or argument
        Exit Sub
    End If
    
    SetVal Me.Source, Value
End Sub

'CSEH: ErrMsgBox
Friend Sub SetVal(ByVal Target As Object, Value)
    If Me.Source Is Nothing Then Exit Sub
    
    ' Allow the function to be called for the child properties only
    If TypeOf Me.Source Is NestedProperty Then
        Err.Raise 5 ' Invalid procedure call or argument
        Exit Sub
    End If
    
    If Me.Child Is Nothing Then
        If IsObject(Value) Then
            CallByName Target, Name, VbSet, Value
        Else
            CallByName Target, Name, VbLet, Value
        End If
    Else
        Set Target = CallByName(Target, Name, VbGet)
        Me.Child.SetVal Target, Value
    End If
End Sub

Public Sub Initialize(ByVal Parent As Object, ByVal Source As Object, ByVal Path As String)
    Init Parent, Source, Path
End Sub

'CSEH: ErrMsgBox
' Just used to be able to pass the Arguments to the corresponding Properties with same names
Private Sub Init(ByVal vParent As Object, ByVal vSource As Object, ByVal vPath As String)
    Path = vPath
    If Len(Path) > 0 Then Name = PathList(0)
    
    Set Parent = vParent
    Set Source = vSource
    
    If Len(Path) = 0 Then Exit Sub
    If UBound(PathList) = 0 Then Exit Sub
    
    CreateChild
    
    SetupChildSource
    
'    Debug.Print "Name: " & Name & _
'                ", Child: " & TypeName(Child) & _
'                ", Path: " & Path & _
'                ", ChildNotifier: " & TypeName(ChildNotifier) & "(" & (ChildNotifier Is Nothing) & ")" & _
'                ", Source: " & TypeName(Source)

End Sub

'CSEH: ErrMsgBox
Private Sub CreateChild()
    Dim InnerPath As String
    Dim InnerPathList
    Dim Index As Long
        
    ' Remove the previous child
    Set m_Child = Nothing
    
    If Len(Path) = 0 Then Exit Sub
    ' No child
    If UBound(PathList) = 0 Then Exit Sub
    
    
    ReDim InnerPathList(0 To UBound(PathList) - 1)
    For Index = 0 To UBound(InnerPathList)
        InnerPathList(Index) = PathList(Index + 1)
    Next
    
    InnerPath = Strings.Join(InnerPathList, ".")
        
    Set m_Child = New NestedProperty
    
    m_Child.Initialize Me, GetNestedObject(Name), InnerPath

End Sub

Friend Sub OnValueChanged(Sender As Object, e As Variant)
    Dim ParentProp  As NestedProperty
    
    If TypeOf Parent Is NestedProperty Then
        Set ParentProp = Parent
        Call ParentProp.OnValueChanged(Me, e)
    Else
        RaiseEvent ValueChanged(Me, e)
    End If
    
End Sub

Private Function GetNestedObject(ByVal PropertyName As String) As Object
    On Error Resume Next
    
    If Source Is Nothing Then Exit Function
    
    Set GetNestedObject = CallByName(Source, PropertyName, VbGet)
    
    If Err.Number <> 0 Then Err.Clear
End Function

Private Sub SetupChildSource()
    On Error Resume Next
    
    Dim Obj As INotifyPropertyChanged
    
    Set ChildNotifier = Nothing
    
    If Child Is Nothing Then Exit Sub
    
    Set Child.Source = Nothing
    
    If Source Is Nothing Then Exit Sub
    
    If Len(Path) = 0 Then Exit Sub
    If UBound(PathList) = 0 Then Exit Sub
    
    Set Child.Source = GetNestedObject(Name)
        
    If Child.Source Is Nothing Then Exit Sub
    
    If TypeOf Child.Source Is INotifyPropertyChanged Then
        Set Obj = Child.Source
        Set ChildNotifier = Obj.PropertyChangedEvent
    End If
End Sub

