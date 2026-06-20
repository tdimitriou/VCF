VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Application"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private WithEvents m_Resources As ResourceDictionary
Attribute m_Resources.VB_VarHelpID = -1
Private m_StartupURI As String
Private m_Windows As cSortedDictionary
Private m_ThemesManager As ThemesManager

Public Event Startup()

Friend Property Get ThemesManager()
    Set ThemesManager = m_ThemesManager
End Property

Public Property Get Windows() As cSortedDictionary
    Set Windows = m_Windows
End Property

Public Property Get StartupURI() As String
    StartupURI = m_StartupURI
End Property

Public Property Let StartupURI(ByVal Value As String)
    m_StartupURI = Value
End Property

Public Property Get Resources() As ResourceDictionary
    Set Resources = m_Resources
End Property

Public Sub Run(Optional ByVal Window As Object)
    Dim StartupObject
    
    Set StartupObject = Window
    
    OnStartup
    
    If StartupObject Is Nothing Then Set StartupObject = GetStartupObject
    
    If StartupObject Is Nothing Then Exit Sub
    
    StartupObject.Show
    
    Cairo.WidgetForms.EnterMessageLoop
End Sub

Private Sub Class_Initialize()
    Set m_Windows = New_c.SortedDictionary(TextCompare)
    Set m_Resources = New ResourceDictionary
End Sub

Friend Sub Initialize(Superclass As IApplication)
    Superclass.InitializeComponent
End Sub

Friend Sub OnInitialized()
    Dim StartupObject As Object
    Set StartupObject = GetStartupObject
    
    If StartupObject Is Nothing Then Exit Sub
    ' Check Valid Type : IWindow
    
    Run StartupObject
End Sub

Friend Sub OnStartup()
    RaiseEvent Startup
End Sub

Private Function GetStartupObject() As Object
    Dim URI As String
    URI = m_StartupURI
    
    If LCase$(Right$(URI, 4)) = ".xml" Then URI = Left$(URI, Len(URI) - 4)
    
    Set GetStartupObject = CreateInstance("", URI)
End Function

Public Function FindResource(ByVal Key As String)
    If m_Resources.TryGetResource(Key, FindResource) Then Exit Function
    Err.Raise 1761
End Function

Public Function TryFindResource(ByVal Key As String)
    m_Resources.TryGetResource Key, TryFindResource
End Function

Private Sub m_Resources_CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
    On Error Resume Next
    
    Dim o
    If Not e.OldItems Is Nothing Then
        For Each o In e.OldItems
            If IsObject(o) Then
                If o Is m_ThemesManager Then
                    Set m_ThemesManager = Nothing
                    Exit For
                End If
            End If
        Next
    End If
    
    If Not e.NewItems Is Nothing Then
        For Each o In e.NewItems
            If IsObject(o) Then
                If TypeOf o Is ThemesManager Then
                    Set m_ThemesManager = o
                    Exit For
                End If
            End If
        Next
    End If
    
    If Err Then Err.Clear
    
End Sub
