Attribute VB_Name = "modConstructors"
Option Explicit

Public CustomConstructor As IObjectConstructor

Public Function NewCollectionChangedEventArgs(ByVal Action As CollectionChangedAction, _
                                                ByVal NewItems As List, _
                                                ByVal NewStartingIndex As Long, _
                                                ByVal OldItems As List, _
                                                ByVal OldStartingIndex As Long) As CollectionChangedEventArgs

    Set NewCollectionChangedEventArgs = New CollectionChangedEventArgs
    
    Call NewCollectionChangedEventArgs.Initialize(Action, NewItems, NewStartingIndex, OldItems, OldStartingIndex)
End Function

Public Function NewCustomObject(ByVal Classname As String) As Object
    On Error Resume Next
    
    If modConstructors.CustomConstructor Is Nothing Then Exit Function
    
    Set NewCustomObject = CustomConstructor.CreateInstance(Classname)
End Function

Public Function NewList(ParamArray Values() As Variant) As List
    Set NewList = New List
    
    Dim v As Variant
    For Each v In Values
        NewList.Add v
    Next
End Function

Public Function NewUIElementCollection(Parent As Object) As UIElementCollection
    On Error GoTo Catch
    
    Dim Obj As UIElementCollection
    
Try:
    Set Obj = New UIElementCollection
    Call Obj.Initialize(Parent)
    
    Set NewUIElementCollection = Obj
        
    Exit Function

Catch:
    Set NewUIElementCollection = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

Public Function NewDependencyProperties(Target As Object) As DependencyProperties
    Dim Obj As DependencyProperties
    Set Obj = New DependencyProperties
    Set NewDependencyProperties = Obj.DependencyProperties(Target)
End Function

Public Function NewThickness(ParamArray Params()) As Thickness
    Set NewThickness = New Thickness
    
    If IsMissing(Params) Then Exit Function
    If IsEmpty(Params) Then Exit Function
    
    With NewThickness
        If UBound(Params) = 0 Then
            .Left = Params(0)
            .Top = Params(0)
            .Right = Params(0)
            .Bottom = Params(0)
        ElseIf UBound(Params) = 1 Then
            .Left = Params(0)
            .Top = Params(1)
            .Right = Params(0)
            .Bottom = Params(1)
        ElseIf UBound(Params) = 3 Then
            .Left = Params(0)
            .Top = Params(1)
            .Right = Params(2)
            .Bottom = Params(3)
        End If
    End With

End Function

Public Function NewBinding(ByVal Source As Object, _
                            ByVal SourcePropertyName As String, _
                            ByVal Target As IDependencyObject, _
                            ByVal TargetProperty As DependencyProperty, _
                            Optional ByVal Converter As IValueConverter, _
                            Optional ByVal StringFormat As String) As Binding
Try:
    On Error GoTo Catch
    
    Set NewBinding = New Binding
    NewBinding.Initialize Source, SourcePropertyName, Target, TargetProperty, Converter, StringFormat
    
    Exit Function
    
Catch:
    Set NewBinding = Nothing
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
End Function

Public Function NewDependencyPropertyMetadata(Optional ByVal AffectsMeasure As Boolean, _
                                                Optional ByVal AffectsRender As Boolean, _
                                                Optional ByVal IsInheritable As Boolean, _
                                                Optional ByVal BindingMode As VCF.BindingMode = VCF.BindingMode.OneWay) _
                                                As DependencyPropertyMetadata
    
    Set NewDependencyPropertyMetadata = New DependencyPropertyMetadata
    
    With NewDependencyPropertyMetadata
        .AffectsMeasure = AffectsMeasure
        .AffectsRender = AffectsRender
        .IsInheritable = IsInheritable
        .BindingMode = BindingMode
    End With
End Function

Public Function NewFunction(ByVal Object As Object, ByVal Method As String, Optional ByVal CallType As VbCallType = VbMethod, Optional Parameter) As VCF.Function
    Set NewFunction = New VCF.Function
    
    With NewFunction
        Set .Object = Object
        .Method = Method
        .CallType = CallType
        If Not IsMissing(Parameter) Then
            If IsObject(Parameter) Then
                Set .Parameter = Parameter
            Else
                .Parameter = Parameter
            End If
        End If
    End With
End Function

Public Function CreateInstance(ByVal Namespace As String, ByVal Class As String) As Object
    On Error Resume Next
    
    Dim Obj
    Dim ProgId As String
    
    If Not CustomConstructor Is Nothing Then
        ProgId = Class
        If LCase$(Namespace) = "res" Then ProgId = Namespace & "." & Class
        Set Obj = modConstructors.NewCustomObject(ProgId)
    End If
    
    If InStr(1, Class, ".") > 0 Then
        ProgId = Class
    Else
        ProgId = Namespace & "." & Class
    End If
    
    If Obj Is Nothing Then Set Obj = CreateObject(ProgId)

    Set CreateInstance = Obj
End Function

Public Function NewUIElementBase(ByVal Superclass As IUIElement, Optional ByRef Baseclass As VCF.UIElementBase) As VCF.UIElementBase
Try:
    On Error GoTo Catch
            
    Dim Obj As VCF.UIElementBase
    
    Set Obj = New VCF.UIElementBase
        
    If Not IsMissing(Baseclass) Then Set Baseclass = Obj
    
    Obj.Initialize Superclass
    
    Set NewUIElementBase = Obj
    
    Exit Function

Catch:
    Set NewUIElementBase = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

