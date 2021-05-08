VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CollectionViewSource"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Views As cSortedDictionary

Public Function GetDefaultView(ByVal Collection As ObservableCollection) As ListCollectionView
    If Views Is Nothing Then Set Views = New_c.SortedDictionary
    
    Dim CollectionPtr As Long
    
    CollectionPtr = Collection.GetHashCode
    
    If Views.Exists(CollectionPtr) Then
        Set GetDefaultView = Views(CollectionPtr)
        Exit Function
    End If
    
    Set GetDefaultView = New ListCollectionView
    GetDefaultView.Initialize Collection
    Views.Add CollectionPtr, GetDefaultView
End Function

Friend Sub DestroyDefaultView(ByVal Ptr As Long)
    If Views Is Nothing Then Exit Sub
    If Not Views.Exists(Ptr) Then Exit Sub
    Views.Remove Ptr
End Sub

Friend Function GetView(ByVal CollectionPtr As Long) As ListCollectionView
    If Views Is Nothing Then Exit Function
    
    If Views.Exists(CollectionPtr) Then Set GetView = Views(CollectionPtr)
End Function
