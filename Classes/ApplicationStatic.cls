VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ApplicationStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_CurrentPtr As Long

Public Property Get Current() As VCF.IApplication
    Dim Obj As Object
    Set Obj = API.ObjFromPtr(m_CurrentPtr)
    
    Set Current = Obj
End Property

Public Property Set Current(ByVal Value As IApplication)
    m_CurrentPtr = ObjPtr(Value)
End Property

Public Sub Create(ByVal Superclass As IApplication)
Try:
    On Error GoTo Catch
            
    Dim Obj As VCF.Application
    
    If Not Me.Current Is Nothing Then
        Err.Raise vbObjectError, App.Title, "Application already created."
        Exit Sub
    End If
    
    Set Obj = New VCF.Application
    Set Me.Current = Superclass
    
    Superclass.SetBase Obj
    
    Obj.Initialize Superclass
    
    Exit Sub

Catch:
    Err.Raise Err.Number, , Err.Description
End Sub

