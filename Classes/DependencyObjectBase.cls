VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependencyObjectBase"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_DependencyProperties As DependencyProperties
Private m_ParentPtr As Long
Private m_ChildrenPtr As Long

Public Property Get DependencyProperties() As DependencyProperties
    Set DependencyProperties = m_DependencyProperties
End Property

Public Property Get Parent() As IDependencyObject
    Set Parent = API.ObjFromPtr(m_ParentPtr)
End Property

Public Property Set Parent(ByVal Value As IDependencyObject)
    m_ParentPtr = ObjPtr(Value)
End Property

Public Property Get Children() As Object
    Set Children = API.ObjFromPtr(m_ChildrenPtr)
End Property

Public Property Set Children(ByVal Value As Object)
    m_ChildrenPtr = Value
End Property

Private Sub Class_Initialize()
    Set m_DependencyProperties = NewDependencyProperties(Me)
End Sub
