VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Conversion"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Function CType(Value, ByVal TargetType As VbVarType)

Try:

    On Error GoTo Catch
        
    Select Case TargetType
    
        Case VbVarType.vbBoolean:       CType = CBool(Value)
        Case VbVarType.vbByte:          CType = CByte(Value)
        Case VbVarType.vbCurrency:      CType = CCur(Value)
        Case VbVarType.vbDataObject:    Set CType = Value
        Case VbVarType.vbDate:          CType = CDate(Value)
        Case VbVarType.vbDecimal:       CType = CDec(Value)
        Case VbVarType.vbDouble:        CType = CDbl(Value)
        Case VbVarType.vbEmpty:         CType = Empty
        Case VbVarType.vbInteger:       CType = CInt(Value)
        Case VbVarType.vbLong:          CType = CLng(Value)
        Case VbVarType.vbNull:          CType = Null
        Case VbVarType.vbObject:        Set CType = Value
        Case VbVarType.vbSingle:        CType = CSng(Value)
        Case VbVarType.vbString:        CType = CStr(Value)
        Case VbVarType.vbVariant:       API.CopyVariable Value, CType
        
        Case Else:                      API.CopyVariable Value, CType
            
    End Select
        
    Exit Function
    
Catch:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Function

Public Function TryCast(Value, ByVal TargetType As VbVarType)

Try:

    On Error GoTo Catch
    
    
    API.CopyVariable CType(Value, TargetType), TryCast
    
    Exit Function
    
Catch:
    
    If TargetType = vbObject Or TargetType = vbDataObject Then
        
        Set TryCast = Nothing
        
    Else
    
        TryCast = Null
        
    End If

End Function

Public Function DirectCast(Value, ByVal TargetType As VbVarType)
    
    API.CopyVariable CType(Value, TargetType), DirectCast
        
End Function

Public Function Cast(Value, ByVal TargetType As VbVarType, Optional FailValue)

Try:
        
    On Error GoTo Catch
    
    API.CopyVariable CType(Value, TargetType), Cast
        
    Exit Function
    
Catch:

    If IsMissing(FailValue) Then
        
        Err.Raise Err.Number, Err.Source, Err.Description, Err.HelpFile, Err.HelpContext
    
    Else
    
        API.CopyVariable FailValue, Cast
        
    End If
    
End Function

Public Function CObj(ByVal Object As Object) As Object
    
    API.CopyVariable API.CObj(Object), CObj

End Function

Public Function CUnk(ByVal Object As IUnknown) As IUnknown

    API.CopyVariable API.CUnk(Object), CUnk

End Function

