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

Public Function Equals(Var1, Var2) As Boolean

Try:
    
    On Error GoTo Catch
    
    If VarType(Var1) = vbUserDefinedType Or VarType(Var2) = vbUserDefinedType Then
        Equals = False
        Exit Function
    End If
    
    If IsArray(Var1) Or IsArray(Var2) Then
        Equals = False
        Exit Function
    End If
    
    If IsObject(Var1) And Not IsObject(Var2) Then
        Equals = False
        Exit Function
    End If
    
    If IsObject(Var2) And Not IsObject(Var1) Then
        Equals = False
        Exit Function
    End If
    
    If IsObject(Var1) And IsObject(Var2) Then
        Equals = (Var1 Is Var2)
        Exit Function
    End If
    
    If IsEmpty(Var1) Then
        Equals = (IsEmpty(Var2))
        Exit Function
    End If
    
    If IsEmpty(Var2) Then
        Equals = (IsEmpty(Var1))
        Exit Function
    End If
    
    If IsNull(Var1) Then
        Equals = (IsNull(Var2))
        Exit Function
    End If
    
    If IsNull(Var2) Then
        Equals = (IsNull(Var1))
        Exit Function
    End If
    
    Equals = (Var1 = Var2)
    
    Exit Function
    
Catch:
    
    ' Just in case ...
    Equals = False
    
End Function

