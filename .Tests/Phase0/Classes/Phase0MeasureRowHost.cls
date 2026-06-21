VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Phase0MeasureRowHost"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents m_LV As ListView
Attribute m_LV.VB_VarHelpID = -1
Private m_ParentHeight As Single
Private m_ChildHeight As Single

Public Sub Setup(ByVal ParentHeight As Single, ByVal ChildHeight As Single, ByVal RowCount As Long)
    m_ParentHeight = ParentHeight
    m_ChildHeight = ChildHeight
    Set m_LV = New ListView
    m_LV.Base.FixedRowHeight = False
    m_LV.Base.ListCount = RowCount
End Sub

Public Function MeasuredHeight(ByVal Index As Long) As Long
    MeasuredHeight = m_LV.Base.MeasuredRowHeight(Index)
End Function

Private Sub m_LV_MeasureRow(ByVal Index As Long, ByVal ItemWidth As Single, RowHeight As Single)
    If Index = 0 Then
        RowHeight = m_ParentHeight
    Else
        RowHeight = m_ChildHeight
    End If
End Sub
