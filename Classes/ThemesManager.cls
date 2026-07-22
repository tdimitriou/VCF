VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ThemesManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements ObservableDictionary

Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
Event ThemeCkanged()

Private WithEvents Base As ObservableDictionary
Attribute Base.VB_VarHelpID = -1
Private m_ActiveThemeName As String
' Host Application.Resources (or test host) that receives the active theme merge.
Private m_HostResources As ResourceDictionary
' Currently merged theme ResourceDictionary instance (for unmerge on swap).
Private m_PublishedTheme As ResourceDictionary

'#Region ObservableDictionary Implementation & Inheritance

    '========== IMPLEMENTATION ==========
    
    Private Sub ObservableDictionary_Add(Key As Variant, Value As Variant)
        Add Key, Value
    End Sub
    
    Private Sub ObservableDictionary_Clear()
        Base.Clear
    End Sub
    
    Private Function ObservableDictionary_Contains(Value As Variant) As Boolean
        ObservableDictionary_Contains = Contains(Value)
    End Function
    
    Private Function ObservableDictionary_ContainsKey(Key As Variant) As Boolean
        ObservableDictionary_ContainsKey = ContainsKey(Key)
    End Function
    
    Private Property Get ObservableDictionary_Count() As Long
        ObservableDictionary_Count = Count
    End Property
    
    Private Function ObservableDictionary_GetEnumerator() As stdole.IUnknown
        Set ObservableDictionary_GetEnumerator = GetEnumerator
    End Function
    
    Private Function ObservableDictionary_IndexOf(Value As Variant) As Long
        ObservableDictionary_IndexOf = IndexOf(Value)
    End Function
    
    Private Function ObservableDictionary_IndexOfKey(Key As Variant) As Long
        ObservableDictionary_IndexOfKey = IndexOfKey(Key)
    End Function
    
    Private Property Let ObservableDictionary_Item(Key As Variant, RHS As Variant)
        Item(Key) = RHS
    End Property

    Private Property Set ObservableDictionary_Item(Key As Variant, RHS As Variant)
        Set Item(Key) = RHS
    End Property
    
    Private Property Get ObservableDictionary_Item(Key As Variant) As Variant
        Set ObservableDictionary_Item = Item(Key)
    End Property
    
    Private Property Let ObservableDictionary_ItemAt(ByVal Index As Long, RHS As Variant)
        Set ItemAt(Index) = RHS
    End Property

    Private Property Set ObservableDictionary_ItemAt(ByVal Index As Long, RHS As Variant)
        Set ItemAt(Index) = RHS
    End Property
    
    Private Property Get ObservableDictionary_ItemAt(ByVal Index As Long) As Variant
        Set ObservableDictionary_ItemAt = ItemAt(Index)
    End Property
    
    Private Function ObservableDictionary_KeyOfIndex(ByVal Index As Long) As Variant
        ObservableDictionary_KeyOfIndex = KeyOfIndex(Index)
    End Function
    
    Private Sub ObservableDictionary_Remove(Value As Variant)
        Remove Value
    End Sub
    
    Private Sub ObservableDictionary_RemoveAt(ByVal Index As Long)
        RemoveAt Index
    End Sub
    
    Private Sub ObservableDictionary_RemoveKey(ByVal Key As Variant)
        RemoveKey Key
    End Sub

    '========== EVENTS IMPLEMENTATION ==========
    
    Private Sub Base_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
        RaiseEvent CollectionChanged(Sender, e)
    End Sub

    '========== INHERITANCE ==========

    Public Sub Add(Key As Variant, Value As Variant)
        Base.Add Key, Value
    End Sub
    
    Public Sub Clear()
        Base.Clear
    End Sub
    
    Public Function Contains(Value As Variant) As Boolean
        Contains = Base.Contains(Value)
    End Function
    
    Public Function ContainsKey(Key As Variant) As Boolean
        ContainsKey = Base.ContainsKey(Key)
    End Function
    
    Public Property Get Count() As Long
        Count = Base.Count
    End Property
    
    Public Function GetEnumerator() As stdole.IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
        Set GetEnumerator = Base.GetEnumerator
    End Function
    
    Public Function IndexOf(Value As Variant) As Long
        IndexOf = Base.IndexOf(Value)
    End Function
    
    Public Function IndexOfKey(Key As Variant) As Long
        IndexOfKey = Base.IndexOfKey(Key)
    End Function
    
    Public Property Set Item(Key As Variant, Value As Variant)
        Set Base.Item(Key) = Value
    End Property
    
    Public Property Get Item(Key As Variant) As Variant
Attribute Item.VB_UserMemId = 0
        Set Item = Base.Item(Key)
    End Property
    
    Public Property Set ItemAt(ByVal Index As Long, Value As Variant)
        Set Base.ItemAt(Index) = Value
    End Property
    
    Public Property Get ItemAt(ByVal Index As Long) As Variant
        Set ItemAt = Base.ItemAt(Index)
    End Property
    
    Public Function KeyOfIndex(ByVal Index As Long) As Variant
        API.CopyVariable Base.KeyOfIndex(Index), KeyOfIndex
    End Function
    
    Public Sub Remove(Value As Variant)
        Base.Remove Value
    End Sub
    
    Public Sub RemoveAt(ByVal Index As Long)
        Base.RemoveAt Index
    End Sub
    
    Public Sub RemoveKey(ByVal Key As Variant)
        Base.RemoveKey Key
    End Sub
'#End Region

    Public Property Get ActiveThemeName() As String
        ActiveThemeName = m_ActiveThemeName
    End Property

    Public Property Let ActiveThemeName(ByVal Value As String)
        Dim Host As ResourceDictionary

        If m_ActiveThemeName = Value Then Exit Property

        m_ActiveThemeName = Value

        Set Host = ResolveHostResources()
        If Not Host Is Nothing Then PublishActiveThemeTo Host

        RaiseEvent ThemeCkanged
    End Property

    Public Property Get ActiveTheme() As ObservableDictionary
        Dim ThemeObj As Object

        If Base.ContainsKey(m_ActiveThemeName) Then
            Set ThemeObj = Base(m_ActiveThemeName)
        ElseIf Base.Count > 0 Then
            Set ThemeObj = Base.ItemAt(0)
        End If

        ' POS / legacy: theme bags are ObservableDictionary.
        ' ResourceDictionary themes resolve via merged Application.Resources.
        If ThemeObj Is Nothing Then Exit Property
        If TypeOf ThemeObj Is ObservableDictionary Then
            Set ActiveTheme = ThemeObj
        End If
    End Property

    ' Wire ThemesManager to Application.Resources (or a Phase0 host dictionary).
    Public Sub AttachToResources(ByVal Resources As ResourceDictionary)
        Set m_HostResources = Resources
        If Not m_HostResources Is Nothing Then PublishActiveThemeTo m_HostResources
    End Sub

    ' WPF-aligned: merge active theme ResourceDictionary into host MergedDictionaries.
    Public Sub PublishActiveThemeTo(ByVal Resources As ResourceDictionary)
        Dim ThemeObj As Object
        Dim ThemeRD As ResourceDictionary
        Dim ThemeOD As ObservableDictionary
        Dim i As Long
        Dim K As Variant
        Dim V As Variant

        On Error GoTo Handler

        If Resources Is Nothing Then Exit Sub

        ' Unmerge previous published theme bag.
        If Not m_PublishedTheme Is Nothing Then
            If Resources.MergedDictionaries.Contains(m_PublishedTheme) Then
                Resources.MergedDictionaries.Remove m_PublishedTheme
            End If
            Set m_PublishedTheme = Nothing
        End If

        If Len(m_ActiveThemeName) = 0 Then Exit Sub
        If Not Base.ContainsKey(m_ActiveThemeName) Then Exit Sub

        Set ThemeObj = Base(m_ActiveThemeName)
        If ThemeObj Is Nothing Then Exit Sub

        If TypeOf ThemeObj Is ResourceDictionary Then
            Set ThemeRD = ThemeObj
        ElseIf TypeOf ThemeObj Is ObservableDictionary Then
            Set ThemeOD = ThemeObj
            Set ThemeRD = New ResourceDictionary
            For i = 0 To ThemeOD.Count - 1
                API.CopyVariable ThemeOD.KeyOfIndex(i), K
                API.CopyVariable ThemeOD.ItemAt(i), V
                If IsObject(V) Then
                    ThemeRD.Add K, V
                Else
                    ThemeRD.Add K, V
                End If
            Next
        Else
            Exit Sub
        End If

        Resources.Merge ThemeRD
        Set m_PublishedTheme = ThemeRD
        Exit Sub

Handler:
        On Error Resume Next
        Set m_PublishedTheme = Nothing
        Err.Clear
    End Sub

    Private Function ResolveHostResources() As ResourceDictionary
        On Error Resume Next
        If Not m_HostResources Is Nothing Then
            Set ResolveHostResources = m_HostResources
            Exit Function
        End If
        If Not Application.Current Is Nothing Then
            Set ResolveHostResources = Application.Current.Base.Resources
        End If
        Err.Clear
    End Function

    Private Sub Class_Initialize()
        Set Base = New ObservableDictionary
    End Sub

