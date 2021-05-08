VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "NamingManager"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'#Region Public Members
    Public Function GetNamedChildren(ByVal Parent As IControl) As ObservableDictionary
        Dim Obj As Object
        
        Set GetNamedChildren = New ObservableDictionary
        
        
        For Each Obj In Parent.Children
            AddNamedItem Obj, GetNamedChildren
        Next
    End Function
'#End Region

'#Region Private Members
    Private Sub AddNamedItem(Obj As Object, Dict As ObservableDictionary)
    
    On Error Resume Next
        
    Dim Name As String
    Dim Children As Object
                
    Err.Clear
    
    Name = Obj.Name
    Set Children = Obj.Children
    
    Err.Clear
        
Try:
    On Error GoTo Catch
     
    If Len(Name) Then Dict.Add Name, Obj
          
    If Children Is Nothing Then Exit Sub
     
    For Each Obj In Children
        AddNamedItem Obj, Dict
    Next
        
    Exit Sub

Catch:
        
        Err.Raise Err.Number, TypeName(Me), Err.Description, Err.HelpFile, Err.HelpContext
    
    End Sub
'#End Region

