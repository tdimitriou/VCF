VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0CollectionSink"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_Collection As VCF.ObservableCollection
Attribute m_Collection.VB_VarHelpID = -1

Public NotifyCount As Long
Public LastAction As VCF.CollectionChangedAction

Public Sub Attach(ByVal Coll As VCF.ObservableCollection)
    NotifyCount = 0
    Set m_Collection = Coll
End Sub

Public Sub Detach()
    Set m_Collection = Nothing
End Sub

Private Sub m_Collection_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
    NotifyCount = NotifyCount + 1
    LastAction = e.Action
End Sub
