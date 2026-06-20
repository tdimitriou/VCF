VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "BindingExpression"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Binding As Binding
Private m_TargetPtr As Long
Private m_IsAttached As Boolean
Private m_IsDetaching As Boolean

Public Property Get Binding() As Binding
    Set Binding = m_Binding
End Property

Public Property Get Target() As IDependencyObject
    Dim Obj As Object
    Set Obj = API.ObjFromPtr(m_TargetPtr)
    If Obj Is Nothing Then Exit Property
    If Not TypeOf Obj Is IDependencyObject Then Exit Property
    Set Target = Obj
End Property

Public Function Attach( _
    ByVal Target As IDependencyObject, _
    ByVal TargetPropertyName As String, _
    ByVal Source As Object, _
    ByVal Path As String, _
    Optional ByVal Mode As BindingMode = BindingMode.Default) As BindingExpression

    Dim TargetProp As DependencyProperty

    If Target Is Nothing Then Exit Function
    If Not Target.DependencyProperties.Exists(TargetPropertyName) Then Exit Function

    Set TargetProp = Target.DependencyProperties.GetProperty(TargetPropertyName)
    Set m_Binding = New Binding
    m_Binding.Mode = Mode
    m_TargetPtr = ObjPtr(Target)
    m_Binding.Initialize Source, Path, Target, TargetProp

    API.CObj(Target).Bindings.Add Me
    m_IsAttached = True
    Set Attach = Me
End Function

Public Sub Detach()
    Dim TargetObj As Object

    If m_Binding Is Nothing Then Exit Sub
    If m_IsDetaching Then Exit Sub
    If Not m_IsAttached Then Exit Sub

    m_IsDetaching = True

    m_Binding.DetachBinding

    If m_TargetPtr <> 0 Then
        On Error Resume Next
        Set TargetObj = API.ObjFromPtr(m_TargetPtr)
        If Not TargetObj Is Nothing Then TargetObj.Bindings.Remove Me
        On Error GoTo 0
    End If

    m_IsAttached = False
    m_TargetPtr = 0
    Set m_Binding = Nothing
    m_IsDetaching = False
End Sub

Public Sub UpdateTarget()
    If Not m_Binding Is Nothing Then m_Binding.RefreshTarget
End Sub

Public Sub UpdateSource()
    If m_Binding Is Nothing Then Exit Sub
    If m_Binding.Mode = OneWay Then Exit Sub
    ' TwoWay / OneWayToSource handled by existing Binding listener graph
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    If m_IsAttached And Not m_IsDetaching Then Detach
End Sub
