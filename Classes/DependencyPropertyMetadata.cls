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

Private Sub Class_Initialize()
    Me.BindingMode = OneWay
End Sub
