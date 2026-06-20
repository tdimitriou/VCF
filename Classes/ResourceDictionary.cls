VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ResourceDictionary"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)

Private WithEvents m_LocalResources As ObservableDictionary
Private m_MergedDictionaries As ObservableCollection
Private m_Source As String
Private m_SourceLoaded As Boolean

Private Sub Class_Initialize()
    Set m_LocalResources = New ObservableDictionary
    Set m_MergedDictionaries = New ObservableCollection
End Sub

Public Property Get Source() As String
    Source = m_Source
End Property

Public Property Let Source(ByVal Value As String)
    m_Source = Trim$(Value)
    m_SourceLoaded = False
End Property

Public Property Get MergedDictionaries() As ObservableCollection
    Set MergedDictionaries = m_MergedDictionaries
End Property

Public Property Get LocalResources() As ObservableDictionary
    Set LocalResources = m_LocalResources
End Property

Public Property Get Count() As Long
    Count = m_LocalResources.Count
End Property

Public Sub Add(Key, Value)
    m_LocalResources.Add Key, Value
End Sub

Public Function ContainsKey(Key) As Boolean
    Dim Dummy As Variant
    ContainsKey = TryGetResource(Key, Dummy)
End Function

Public Function Contains(Value) As Boolean
    Contains = m_LocalResources.Contains(Value)
End Function

Public Function IndexOfKey(Key) As Long
    IndexOfKey = m_LocalResources.IndexOfKey(Key)
End Function

Public Sub RemoveKey(ByVal Key)
    m_LocalResources.RemoveKey Key
End Sub

Public Sub Clear()
    m_LocalResources.Clear
End Sub

Public Property Get Item(Key) As Variant
Attribute Item.VB_UserMemId = 0
    Dim Value As Variant
    If Not TryGetResource(Key, Value) Then
        Err.Raise 9
    End If
    If IsObject(Value) Then
        Set Item = Value
    Else
        Item = Value
    End If
End Property

Public Property Let Item(Key, Value)
    m_LocalResources.Item(Key) = Value
End Property

Public Property Set Item(Key, Value)
    Set m_LocalResources.Item(Key) = Value
End Property

Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
    Set GetEnumerator = m_LocalResources.GetEnumerator
End Function

Public Function TryGetResource(ByVal Key As String, ByRef Value As Variant) As Boolean
    EnsureSourceLoaded

    If m_LocalResources.ContainsKey(Key) Then
        API.CopyVariable m_LocalResources(Key), Value
        TryGetResource = True
        Exit Function
    End If

    Dim Merged As Variant
    Dim Dict As ResourceDictionary

    For Each Merged In m_MergedDictionaries
        If TypeOf Merged Is ResourceDictionary Then
            Set Dict = Merged
            If Dict.TryGetResource(Key, Value) Then
                TryGetResource = True
                Exit Function
            End If
        End If
    Next
End Function

Public Function FindResource(ByVal Key As String) As Variant
    If Not TryGetResource(Key, FindResource) Then
        Err.Raise 1761, , "Resource '" & Key & "' not found."
    End If
End Function

Public Sub Merge(ByVal Dictionary As ResourceDictionary)
    If Dictionary Is Nothing Then Exit Sub
    m_MergedDictionaries.Add Dictionary
End Sub

Friend Sub EnsureSourceLoaded()
    Dim Loaded As ResourceDictionary
    Dim Resolver As XamlResourceResolver

    If m_SourceLoaded Then Exit Sub
    If Len(m_Source) = 0 Then
        m_SourceLoaded = True
        Exit Sub
    End If

    Set Resolver = New XamlResourceResolver
    Set Loaded = Resolver.LoadFromSource(m_Source)
    If Not Loaded Is Nothing Then
        Merge Loaded
    End If
    m_SourceLoaded = True
End Sub

Private Sub m_LocalResources_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
    RaiseEvent CollectionChanged(Sender, e)
End Sub
