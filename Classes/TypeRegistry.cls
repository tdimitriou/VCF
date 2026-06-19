VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TypeRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_ProgIds As ObservableDictionary

Private Sub Class_Initialize()
    Set m_ProgIds = New ObservableDictionary
End Sub

' Register an app or framework type by logical name (e.g. "ShellWindow", "VCF.Button").
' ProgId is passed to CreateObject (e.g. "MyApp.ShellWindow", "VCF.Button").
Public Sub Register(ByVal TypeName As String, ByVal ProgId As String)
    If m_ProgIds.ContainsKey(TypeName) Then
        m_ProgIds(TypeName) = ProgId
    Else
        m_ProgIds.Add TypeName, ProgId
    End If
End Sub

Public Function IsRegistered(ByVal TypeName As String) As Boolean
    IsRegistered = m_ProgIds.ContainsKey(TypeName)
End Function

Public Function ResolveProgId(ByVal TypeName As String) As String
    If Not m_ProgIds.ContainsKey(TypeName) Then Exit Function
    ResolveProgId = m_ProgIds(TypeName)
End Function

Public Function TryCreate(ByVal TypeName As String) As Object
    Dim ProgId As String
    ProgId = ResolveProgId(TypeName)
    If Len(ProgId) = 0 Then Exit Function

    On Error GoTo Fail
    Set TryCreate = CreateObject(ProgId)
    Exit Function

Fail:
    If modConstructors.StrictXamlLoad Then
        RaiseXamlLoad "TypeRegistry failed to create '" & TypeName & "' (" & ProgId & "): " & Err.Description, _
            TypeName, vbNullString, 0, 0, Err.Number
    End If
End Function

' Namespace.Class resolution used by XAMLReader.CreateInstance.
Public Function TryCreateInstance(ByVal Namespace As String, ByVal ClassName As String) As Object
    Dim FullName As String
    Dim ProgId As String

    FullName = ClassName
    If InStr(1, FullName, ".") = 0 Then
        If Len(Namespace) > 0 And LCase$(Namespace) <> "vcf" Then
            FullName = Namespace & "." & ClassName
        End If
    End If

    Set TryCreateInstance = TryCreate(FullName)
    If Not TryCreateInstance Is Nothing Then Exit Function

    Set TryCreateInstance = TryCreate(ClassName)
    If Not TryCreateInstance Is Nothing Then Exit Function

    ProgId = ResolveProgId(FullName)
    If Len(ProgId) = 0 Then ProgId = ResolveProgId(ClassName)
    If Len(ProgId) = 0 Then Exit Function

    On Error GoTo Fail
    Set TryCreateInstance = CreateObject(ProgId)
    Exit Function

Fail:
    If modConstructors.StrictXamlLoad Then
        RaiseXamlLoad "TypeRegistry failed to create '" & FullName & "': " & Err.Description, _
            FullName, vbNullString, 0, 0, Err.Number
    End If
End Function

Public Sub Clear()
    m_ProgIds.Clear
End Sub
