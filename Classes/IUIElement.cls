VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IUIElement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Set DataContext(ByVal Value As Object): End Property
Public Property Get DataContext() As Object: End Property
Public Property Get Base() As UIElementBase: End Property
Public Property Get AttachedProperties() As ObservableDictionary: End Property

Public Property Get Parent() As IControl: End Property
Public Property Set Parent(Value As IControl): End Property
Public Sub Move(Left As Single, Optional Top As Single = 0, Optional Width As Single, Optional Height As Single): End Sub
