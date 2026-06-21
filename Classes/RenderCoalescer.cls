VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "RenderCoalescer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Sub BeginRenderUpdate()
    modRenderCoalescer.BeginRenderUpdate
End Sub

Public Sub EndRenderUpdate()
    modRenderCoalescer.EndRenderUpdate
End Sub

Public Sub RequestWidgetRefresh(ByVal W As cWidgetBase)
    modRenderCoalescer.RequestWidgetRefresh W
End Sub

Public Function PendingCount() As Long
    PendingCount = modRenderCoalescer.PendingRefreshCount
End Function

Public Function LastFlushCount() As Long
    LastFlushCount = modRenderCoalescer.LastFlushRefreshCount
End Function
