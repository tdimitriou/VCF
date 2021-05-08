VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type CornerRadius
    TopLeft As Double
    TopRight As Double
    BottomLeft As Double
    BottomRight As Double
End Type

Public Enum ClickMode
    ClickModeRelease = 0
    ClickModePress = 1
    ClickModeHover = 2
End Enum

Public Enum Visibility
    VisibilityVisible = 0
    VisibilityHidden = 1
    VisibilityCollapsed = 2
End Enum

Public Property Get Widget() As cWidgetBase: End Property
Public Property Get Widgets() As cWidgets: End Property
Public Property Get Children() As UIElementCollection: End Property
