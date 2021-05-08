VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Function"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ObjectPtr             As Long
Private m_Parameter             As Variant
Private m_ParameterPtr          As Long
Private m_IsObjectParameter     As Boolean
Public Method                   As String
Public CallType                 As VbCallType

Public Function GetHashCode() As Long
    GetHashCode = ObjPtr(Me)
End Function

Public Property Get Parameter() As Variant
    If m_IsObjectParameter Then
        If m_ParameterPtr > 0 Then
            Set Parameter = API.ObjFromPtr(m_ParameterPtr)
        Else
            Set Parameter = Nothing
        End If
    Else
        Parameter = m_Parameter
    End If
End Property

Public Property Let Parameter(ByVal Value As Variant)
    m_ParameterPtr = 0
    m_Parameter = Value
    m_IsObjectParameter = False
End Property

Public Property Set Parameter(ByVal Value As Variant)
    m_ParameterPtr = ObjPtr(Value)
    m_Parameter = Empty
    m_IsObjectParameter = True
End Property

Public Property Get Object() As Object
    Set Object = API.ObjFromPtr(m_ObjectPtr)
End Property

Public Property Set Object(ByVal Value As Object)
    m_ObjectPtr = ObjPtr(Value)
End Property

Public Sub Class_Initialize()
    CallType = VbMethod
    m_Parameter = Empty
End Sub

Public Function Execute(Optional Parameter)
    If IsMissing(Parameter) Or IsEmpty(Parameter) Or IsNull(Parameter) Then
        If IsEmpty(Me.Parameter) Then
            Execute = CallByName(Object, Method, CallType)
        Else
            Execute = CallByName(Object, Method, CallType, Me.Parameter)
        End If
    Else
        Execute = CallByName(Object, Method, CallType, Parameter)
    End If
End Function
