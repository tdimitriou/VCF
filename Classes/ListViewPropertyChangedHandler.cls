VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListViewPropertyChangedHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_Notifier As VCF.PropertyChangedEvent
Attribute m_Notifier.VB_VarHelpID = -1
Private m_ListViewPtr As Long

Private Property Get ListView() As ListView
    Set ListView = API.ObjFromPtr(m_ListViewPtr)
End Property

Friend Sub Init(Notifier As VCF.PropertyChangedEvent, LV As VCF.ListView)
    Set m_Notifier = Notifier
        
    m_ListViewPtr = ObjPtr(LV)
End Sub

Private Sub m_Notifier_PropertyChanged(ByVal Sender As Object, e As Variant)
    On Error Resume Next
        
    ListView.PropertyChangedCallback Sender, e
End Sub


