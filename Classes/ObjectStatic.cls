VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ObjectStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function Equals(Value1, Value2) As Boolean

Try:
    
    On Error GoTo Catch
    
    If VarType(Value1) = vbUserDefinedType Or VarType(Value2) = vbUserDefinedType Then
        Equals = False
        Exit Function
    End If
    
    If IsArray(Value1) Or IsArray(Value2) Then
        Equals = False
        Exit Function
    End If
    
    If IsObject(Value1) And Not IsObject(Value2) Then
        Equals = False
        Exit Function
    End If
    
    If IsObject(Value2) And Not IsObject(Value1) Then
        Equals = False
        Exit Function
    End If
    
    If IsObject(Value1) And IsObject(Value2) Then
        Equals = (Value1 Is Value2)
        Exit Function
    End If
    
    If IsEmpty(Value1) Then
        Equals = (IsEmpty(Value2))
        Exit Function
    End If
    
    If IsEmpty(Value2) Then
        Equals = (IsEmpty(Value1))
        Exit Function
    End If
    
    If IsNull(Value1) Then
        Equals = (IsNull(Value2))
        Exit Function
    End If
    
    If IsNull(Value2) Then
        Equals = (IsNull(Value1))
        Exit Function
    End If
    
    Dim v1, v2
    API.CopyVariable Value1, v1
    API.CopyVariable Value2, v2
    
    If VarType(v1) <> vbString And VarType(v2) = vbString Then v2 = Conversion.Cast(v2, VarType(v1), v2)
    If VarType(v2) <> vbString And VarType(v1) = vbString Then v1 = Conversion.Cast(v1, VarType(v2), v1)
    
    Equals = (v1 = v2)
    
    Exit Function
    
Catch:
    
    ' Just in case ...
    Equals = False
    
End Function

Public Function Greater(Value1, Value2) As Boolean

Try:
    
    On Error GoTo Catch
        
    If IsObject(Value1) Or IsObject(Value2) Then
        Greater = False
        Exit Function
    End If
    
    Dim v1, v2
    API.CopyVariable Value1, v1
    API.CopyVariable Value2, v2
    
    If VarType(v1) <> vbString And VarType(v2) = vbString Then v2 = Conversion.Cast(v2, VarType(v1), v2)
    If VarType(v2) <> vbString And VarType(v1) = vbString Then v1 = Conversion.Cast(v1, VarType(v2), v1)
    
    Greater = (v1 > v2)
    
    Exit Function
    
Catch:
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
End Function

Public Function Less(Value1, Value2) As Boolean

Try:
    
    On Error GoTo Catch
        
    If IsObject(Value1) Or IsObject(Value2) Then
        Less = False
        Exit Function
    End If
    
    Dim v1, v2
    API.CopyVariable Value1, v1
    API.CopyVariable Value2, v2
    
    If VarType(v1) <> vbString And VarType(v2) = vbString Then v2 = Conversion.Cast(v2, VarType(v1), v2)
    If VarType(v2) <> vbString And VarType(v1) = vbString Then v1 = Conversion.Cast(v1, VarType(v2), v1)
    
    Less = (v1 < v2)
    
    Exit Function
    
Catch:
    
    Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext

End Function

Public Function EqualOrGreater(Value1, Value2) As Boolean
    
Try:

    On Error GoTo Catch
    
    EqualOrGreater = (Equals(Value1, Value2) Or Greater(Value1, Value2))

    Exit Function
    
Catch:
    
    EqualOrGreater = False

End Function

Public Function EqualOrLess(Value1, Value2) As Boolean
    
Try:

    On Error GoTo Catch
    
    EqualOrLess = (Equals(Value1, Value2) Or Less(Value1, Value2))

    Exit Function
    
Catch:
    
    EqualOrLess = False

End Function

Public Function Between(Value, _
                            Value1, _
                            Value2, _
                            Optional ByVal Included As Boolean = True) As Boolean

Try:
    On Error GoTo Catch
        
    If Included Then
        Between = (EqualOrGreater(Value, Value1) And EqualOrLess(Value, Value2))
    Else
        Between = (Greater(Value, Value1) And Less(Value, Value2))
    End If
    
    Exit Function
    
Catch:
    
    Err.Clear
    Between = False

End Function
