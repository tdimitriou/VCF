VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XAMLDependencyPropertyManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'CSEH: ErrMsgBox
    Public Function GetPropertyValueFromString(ByVal Prop As DependencyProperty, Value) As Variant
Try:
        On Error GoTo Catch
        
        With Prop
            If .PropertyType = vbUserDefinedType Then
                GetPropertyValueFromString = GetUDFProp(.PropertyTypeName, Value)
            ElseIf .PropertyType = vbObject Or .PropertyType = vbDataObject Then
                If IsObject(Value) And (Len(.PropertyTypeName) = 0 Or TypeName(Value) = .PropertyTypeName) Then
                    Call API.CopyVariable(Value, GetPropertyValueFromString)
                Else
                    'If Prop.Name = "DataContext" Then Debug.Print TypeName(Value)
                    Call API.CopyVariable(GetObjectProperty(.PropertyTypeName, Value), GetPropertyValueFromString)
                    'If Prop.Name = "DataContext" Then Debug.Print TypeName(GetPropertyValueFromString)
                End If
            Else
                GetPropertyValueFromString = Value
            End If
        End With
        
        Exit Function
        
Catch:
        ' Return the initial value if error
        Call API.CopyVariable(Value, GetPropertyValueFromString)
    End Function

    Private Function GetUDFProp(PropertyTypeName As String, ByVal Value As String) As Variant
        Select Case PropertyTypeName
        
            Case "CornerRadius": GetUDFProp = modUDFConstructors.NewCornerRadius(Value)
                
        End Select
    End Function
    
    Private Function GetObjectProperty(PropertyTypeName As String, ByVal Value As String) As Object
        On Error Resume Next
        
        Set GetObjectProperty = modConstructors.NewCustomObject(Value)
        
        If Not GetObjectProperty Is Nothing Then Exit Function
        
        Select Case PropertyTypeName
        
            Case "Thickness"
                
                With New XAMLThicknessConstructor
                    Set GetObjectProperty = .NewThickness(Value)
                End With
            
            Case Else
                Set GetObjectProperty = CreateObject("VCF." & Value)
        End Select
    
    End Function

