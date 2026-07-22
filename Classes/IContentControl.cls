VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior = 0  'vbNone
  MTSTransactionMode = 0  'NotAnMTSObject
END
Attribute VB_Name = "IContentControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' WPF ContentControl contract (VB6: implement + compose ContentHost; no class inheritance).

Public Property Get Content() As Variant: End Property
Public Property Let Content(ByVal Value As Variant): End Property
Public Property Set Content(ByVal Value As Object): End Property

Public Property Get ContentPresenter() As ContentPresenter: End Property

Public Sub SyncContentPresenter(): End Sub
