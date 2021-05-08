VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "SelfBinding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements IMarkupExtension

Private m_Object As Object

Public Function InitializeSelf(ByVal Self As Object) As SelfBinding
    Set m_Object = Self
    Set InitializeSelf = Me
End Function

Private Function IMarkupExtension_ProvideValue() As Variant
    Set IMarkupExtension_ProvideValue = m_Object
End Function
