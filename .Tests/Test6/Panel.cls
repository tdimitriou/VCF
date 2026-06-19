VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Panel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1

Public Property Get Widget() As cWidgetBase
    Set Widget = W
End Property

Public Property Get Widgets() As cWidgets
    Set Widgets = W.Widgets
End Property

Private Sub Class_Initialize()
    Set W = Cairo.WidgetBase
End Sub

