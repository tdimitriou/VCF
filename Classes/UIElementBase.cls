VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UIElementBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Resources As ObservableDictionary
Private m_SuperclassPtr As Long
Private m_AttachedProperties As ObservableDictionary

Public Property Get AttachedProperties() As ObservableDictionary
    Set AttachedProperties = m_AttachedProperties
End Property

Public Property Get Resources() As ObservableDictionary
    Set Resources = m_Resources
End Property

'CSEH: ErrMsgBox
Public Function FindResource(ByVal Key As String)
    Dim Parent As IUIElement
    
    If m_Resources.ContainsKey(Key) Then
        API.CopyVariable m_Resources(Key), FindResource
    Else
        Set Parent = Superclass.Parent
        
        If Parent Is Nothing Then
            If Application.Current Is Nothing Then
                Err.Raise 1761
            Else
                API.CopyVariable Application.Current.Base.FindResource(Key), FindResource
            End If
        Else
            API.CopyVariable Parent.Base.FindResource(Key), FindResource
        End If
    End If
End Function

'CSEH: ErrMsgBox
Public Function TryFindResource(ByVal Key As String)
    On Error Resume Next
    
    API.CopyVariable FindResource(Key), TryFindResource
    
    If Err.Number = 0 Then Exit Function
    
    TryFindResource = Empty
End Function

Private Sub Class_Initialize()
    Set m_Resources = New ObservableDictionary
    Set m_AttachedProperties = New ObservableDictionary
End Sub

Friend Sub Initialize(ByVal Superclass As IUIElement)
    m_SuperclassPtr = ObjPtr(Superclass)
End Sub

Private Property Get Superclass() As IUIElement
    Dim Obj As Object
    Set Obj = API.ObjFromPtr(m_SuperclassPtr)
    Set Superclass = Obj
End Property

