VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0SelectionHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_LV As ListView
Attribute m_LV.VB_VarHelpID = -1
Private m_ListIndexChangedCount As Long
Private m_SelectionChangedCount As Long

Public Sub Setup(ByVal Coll As ObservableCollection)
    m_ListIndexChangedCount = 0
    m_SelectionChangedCount = 0
    Set m_LV = New ListView
    Set m_LV.ItemsSource = Coll
End Sub

Public Property Get ListView() As ListView
    Set ListView = m_LV
End Property

Public Property Get ListIndexChangedCount() As Long
    ListIndexChangedCount = m_ListIndexChangedCount
End Property

Public Property Get SelectionChangedCount() As Long
    SelectionChangedCount = m_SelectionChangedCount
End Property

Public Sub ResetCounts()
    m_ListIndexChangedCount = 0
    m_SelectionChangedCount = 0
End Sub

Private Sub m_LV_ListIndexChanged()
    m_ListIndexChangedCount = m_ListIndexChangedCount + 1
End Sub

Private Sub m_LV_SelectionChanged()
    m_SelectionChangedCount = m_SelectionChangedCount + 1
End Sub
