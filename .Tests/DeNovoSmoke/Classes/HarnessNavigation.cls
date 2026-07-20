VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "HarnessNavigation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetActiveView(ByVal Host As Window, ByVal View As Object)
    On Error GoTo Fail
    
    If View Is Nothing Then Exit Sub
    
    Dim i As Long
    Dim Child As Object
    Dim InTree As Boolean
    
    InTree = False
    For i = 0 To Host.Children.Count - 1
        Set Child = Host.Children(i)
        If Child Is View Then
            InTree = True
        ElseIf TypeOf Child Is IUserControl Then
            Child.Visibility = VisibilityCollapsed
        End If
    Next i
    
    If Not InTree Then Host.Children.Add View
    
    If TypeOf View Is IUserControl Then View.Visibility = VisibilityVisible
    
    Host.RelayoutChildren
    Host.RebuildNamedItemsList
    
    If TypeOf View Is HarnessScreen Then View.OnActivated
    
    Host.Form.WidgetRoot.Refresh
    
    Exit Sub
    
Fail:
    
    Err.Raise Err.Number, Err.Source, Err.Description
    
End Sub
