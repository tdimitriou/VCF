VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "XamlResourceResolver"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_BasePath As String

Public Property Get BasePath() As String
    BasePath = m_BasePath
End Property

Public Property Let BasePath(ByVal Value As String)
    m_BasePath = Value
End Property

Public Function LoadFromSource(ByVal Source As String) As ResourceDictionary
    Dim Path As String
    Dim Xml As String
    Dim Dom As cSimpleDOM
    Dim Root As cElement

    Set LoadFromSource = Nothing
    If Len(Trim$(Source)) = 0 Then Exit Function

    Path = ResolveSourcePath(Source)
    If Len(Path) = 0 Then Exit Function
    If Not New_c.FSO.FileExists(Path) Then
        If modConstructors.StrictXamlLoad Then
            RaiseXamlLoad "ResourceDictionary Source not found: '" & Source & "' (" & Path & ")", "ResourceDictionary", "Source"
        End If
        Exit Function
    End If

    Xml = New_c.FSO.ReadByteContent(Path)
    Set Dom = New_c.SimpleDOM
    Dom.XML = Xml

    If Not Dom.WellFormed Then
        If modConstructors.StrictXamlLoad Then
            RaiseXamlLoad "Malformed resource dictionary: '" & Source & "'", "ResourceDictionary", "Source"
        End If
        Exit Function
    End If

    Set Root = Dom.Root
    Set LoadFromSource = LoadDictionaryNode(Root)
End Function

Public Function LoadDictionaryNode(ByVal Node As cElement) As ResourceDictionary
    Dim Dict As ResourceDictionary
    Dim Child As cElement
    Dim Source As String
    Dim Merged As ResourceDictionary
    Dim Key As String
    Dim Data As Variant

    Set Dict = New ResourceDictionary

    On Error Resume Next
    Source = Node.Attributes("Source").Text
    If Len(Source) > 0 Then
        Dict.Source = Source
        Dict.EnsureSourceLoaded
    End If
    On Error GoTo 0

    If Not Node.HasChildElements Then
        Set LoadDictionaryNode = Dict
        Exit Function
    End If

    For Each Child In Node.ChildElements
        Select Case Child.BaseName
            Case "ResourceDictionary.MergedDictionaries"
                LoadMergedDictionaries Dict, Child
            Case "Style"
                Dim Reader As XAMLStyleReader
                Set Reader = New XAMLStyleReader
                Reader.LoadStyle Dict, Child
            Case "ResourceDictionary"
                Set Merged = LoadDictionaryNode(Child)
                If Not Merged Is Nothing Then Dict.Merge Merged
            Case Else
                Key = GetResourceKey(Child)
                If Len(Key) > 0 Then
                    API.CopyVariable InstantiateResource(Child), Data
                    If Not IsEmpty(Data) And Not IsNull(Data) Then
                        If IsObject(Data) Then
                            If Not Data Is Nothing Then Dict.Add Key, Data
                        Else
                            Dict.Add Key, Data
                        End If
                    End If
                End If
        End Select
    Next

    Set LoadDictionaryNode = Dict
End Function

Private Sub LoadMergedDictionaries(ByVal Target As ResourceDictionary, ByVal Node As cElement)
    Dim Child As cElement
    Dim Source As String
    Dim Merged As ResourceDictionary

    If Not Node.HasChildElements Then Exit Sub

    For Each Child In Node.ChildElements
        Select Case Child.BaseName
            Case "ResourceDictionary"
                Source = vbNullString
                On Error Resume Next
                Source = Child.Attributes("Source").Text
                On Error GoTo 0
                If Len(Source) > 0 Then
                    Set Merged = LoadFromSource(Source)
                Else
                    Set Merged = LoadDictionaryNode(Child)
                End If
                If Not Merged Is Nothing Then Target.Merge Merged
        End Select
    Next
End Sub

Private Function InstantiateResource(ByVal Node As cElement) As Variant
    Dim Reader As XAMLReader
    Dim Obj As Object

    Select Case LCase$(Node.BaseName)
        Case "string"
            InstantiateResource = vbNullString
            On Error Resume Next
            InstantiateResource = Node.Attributes("Value").Text
            On Error GoTo 0
        Case Else
            Set Reader = New XAMLReader
            Set Obj = Reader.LoadElement(Node)
            If Not Obj Is Nothing Then
                Set InstantiateResource = Obj
            End If
    End Select
End Function

Private Function GetResourceKey(ByVal Node As cElement) As String
    On Error Resume Next
    GetResourceKey = Node.Attributes("Key").Text
    GetResourceKey = Node.Attributes("x:Key").Text
    If Err Then Err.Clear
End Function

Private Function ResolveSourcePath(ByVal Source As String) As String
    Dim Path As String

    Path = Trim$(Source)
    If Len(Path) = 0 Then Exit Function

    If InStr(1, Path, ":\") > 0 Or Left$(Path, 2) = "\\" Then
        ResolveSourcePath = Path
        Exit Function
    End If

    If Len(m_BasePath) > 0 Then
        ResolveSourcePath = m_BasePath & "\" & Path
    Else
        ResolveSourcePath = Path
    End If
End Function
