VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Variable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Value

'#Region Public Properties
    Public Property Get IsEmpty() As Boolean
        IsEmpty = VBA.IsEmpty(m_Value)
    End Property
    
    Public Property Get IsNothing() As Boolean
        IsNothing = VCF.IsNothing(m_Value)
    End Property
    
    Public Property Get IsNull() As Boolean
        IsNull = VBA.IsNull(m_Value)
    End Property
    
    Public Property Get IsNumeric() As Boolean
        IsNumeric = VBA.IsNumeric(m_Value)
    End Property
    
    Public Property Get Typename() As String
        Typename = VBA.Typename(m_Value)
    End Property
    
    Public Property Get Value()
Attribute Value.VB_UserMemId = 0
        API.CopyVariable m_Value, Value
    End Property
    
    Public Property Let Value(NewValue)
        API.CopyVariable NewValue, m_Value
    End Property
    
    Public Property Get VarType() As VbVarType
        VarType = VBA.VarType(m_Value)
    End Property
'#End Region

'#Region Comparison Functions
    Public Function Between(Value1, _
                            Value2, _
                            Optional ByVal Included As Boolean = True) As Boolean
    
        Between = Object.Between(m_Value, Value1, Value2, Included)
    
    End Function

    Public Function Equals(Value) As Boolean
        
        Equals = Object.Equals(m_Value, Value)
        
    End Function

    Public Function EqualOrLess(Value) As Boolean
        
        EqualOrLess = Object.EqualOrLess(m_Value, Value)
        
    End Function

    Public Function EqualOrGreater(Value) As Boolean
        
        EqualOrGreater = Object.EqualOrGreater(m_Value, Value)
        
    End Function
'#End Region

'#Region Casting Functions
    Public Function CType(TargetType As VbVarType)
        API.CopyVariable Conversion.CType(m_Value, TargetType), CType
    End Function
    
    Public Function DirectCast(TargetType As VbVarType)
        API.CopyVariable Conversion.DirectCast(m_Value, TargetType), DirectCast
    End Function
    
    Public Function TryCast(TargetType As VbVarType)
        API.CopyVariable Conversion.TryCast(m_Value, TargetType), TryCast
    End Function
'#End Region
