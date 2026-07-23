VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependencyPropertyMetadata"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public IsInheritable        As Boolean
Public AffectsRender        As Boolean
Public AffectsMeasure       As Boolean
Public BindingMode          As VCF.BindingMode

Private m_DefaultValue      As Variant
Private m_HasDefaultValue   As Boolean

Public Property Get HasDefaultValue() As Boolean
    HasDefaultValue = m_HasDefaultValue
End Property

Public Property Get DefaultValue() As Variant
    Call API.CopyVariable(m_DefaultValue, DefaultValue)
End Property

Public Property Let DefaultValue(ByVal Value As Variant)
    Call API.CopyVariable(Value, m_DefaultValue)
    m_HasDefaultValue = True
End Property

Public Property Set DefaultValue(ByVal Value As Object)
    Set m_DefaultValue = Value
    m_HasDefaultValue = True
End Property

Private Sub Class_Initialize()
    Me.BindingMode = OneWay
    m_HasDefaultValue = False
End Sub
