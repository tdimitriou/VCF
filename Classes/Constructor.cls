VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Constructor"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type KeyValuePair
    Key     As Variant
    Value   As Variant
End Type

Public Sub SetCustomConstructor(ByVal Cctor As IObjectConstructor)
    Set modConstructors.CustomConstructor = Cctor
End Sub

Public Function GetCustomConstructor() As IObjectConstructor
    Set GetCustomConstructor = modConstructors.CustomConstructor
End Function

Public Function NewCollectionChangedEventArgs(ByVal Action As CollectionChangedAction, _
                                                ByVal NewItems As List, _
                                                ByVal NewStartingIndex As Long, _
                                                ByVal OldItems As List, _
                                                ByVal OldStartingIndex As Long) As CollectionChangedEventArgs

    Set NewCollectionChangedEventArgs = New CollectionChangedEventArgs
    
    Call NewCollectionChangedEventArgs.Initialize(Action, NewItems, NewStartingIndex, OldItems, OldStartingIndex)
End Function

Public Function NewDependencyProperties(Target As Object) As DependencyProperties
    Set NewDependencyProperties = modConstructors.NewDependencyProperties(Target)
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

Public Function NewWindow(ByVal Superclass As IWindow, _
                            Optional ByRef Baseclass As VCF.Window, _
                            Optional ParentWindow As Object = Nothing, _
                            Optional CanGetFocus As Boolean = True) As VCF.Window

Try:
    
    On Error GoTo Catch
            
    Dim Obj As VCF.Window
    
    Set Obj = New VCF.Window
        
    If Not IsMissing(Baseclass) Then Set Baseclass = Obj
    
    Obj.Initialize Superclass, ParentWindow, CanGetFocus
    
    Set NewWindow = Obj
    
    Exit Function

Catch:
    
    Set NewWindow = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

Public Function NewUIElementBase(ByVal Superclass As IUIElement, Optional ByRef Baseclass As VCF.UIElementBase) As VCF.UIElementBase
    API.CopyVariable modConstructors.NewUIElementBase(Superclass, Baseclass), NewUIElementBase
End Function

Public Function NewUserControl(ByVal Superclass As IUserControl, Optional ByRef Baseclass As VCF.UserControl) As VCF.UserControl
Try:
    On Error GoTo Catch
            
    Dim Obj As VCF.UserControl
    
    Set Obj = New VCF.UserControl
        
    If Not IsMissing(Baseclass) Then Set Baseclass = Obj
    
    Obj.Initialize Superclass
    
    Set NewUserControl = Obj
    
    Exit Function

Catch:
    Set NewUserControl = Nothing
    Err.Raise Err.Number, , Err.Description
End Function

Public Function NewBinding(ByVal Source As Object, _
                            ByVal SourcePropertyName As String, _
                            ByVal Target As IDependencyObject, _
                            ByVal TargetProperty As DependencyProperty, _
                            Optional ByVal Converter As IValueConverter, _
                            Optional ByVal StringFormat As String) As Binding
    
    Set NewBinding = modConstructors.NewBinding(Source, _
                                                SourcePropertyName, _
                                                Target, _
                                                TargetProperty, _
                                                Converter, _
                                                StringFormat)

End Function

Public Function CreateInstance(ByVal Namespace As String, ByVal Class As String) As Object
    Set CreateInstance = modConstructors.CreateInstance(Namespace, Class)
End Function

Public Function NewFunction(ByVal Object As Object, ByVal Method As String, Optional ByVal CallType As VbCallType = VbMethod, Optional Parameter) As VCF.Function
    Set NewFunction = modConstructors.NewFunction(Object, Method, CallType, Parameter)
End Function


