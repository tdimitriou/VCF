VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ArrayWrapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_SourceArray
Private m_IsInitialized As Boolean

Public Sub Initialize(ByRef SourceArray)

    If m_IsInitialized Then
    
        Err.Raise vbObjectError + 1, Typename(Me), "Object already initialized"
        Exit Sub
        
    End If
            
    m_IsInitialized = Not (API.pArrPtr(SourceArray) = 0)
    If m_IsInitialized Then m_SourceArray = SourceArray
    
End Sub

Public Function InitializeParamArray(ParamArray Params()) As ArrayWrapper
    
    If m_IsInitialized Then
    
        Err.Raise vbObjectError + 1, Typename(Me), "Object already initialized"
        Exit Function
        
    End If
    
    Dim Parameters
    Parameters = Params
    
    Initialize Parameters
    
    Set InitializeParamArray = Me
End Function

Public Property Get SourceArray()
    SourceArray = m_SourceArray
End Property

Public Property Get IsInitialized() As Boolean

    IsInitialized = m_IsInitialized
    
End Property

Function Find(Value, Optional ByRef Index As Long) As Boolean

Try:

    On Error GoTo Catch
    
    Dim Result As Boolean
    
    If Not m_IsInitialized Then GoTo Catch
    
    Dim Idx As Long
    For Idx = LBound(m_SourceArray) To UBound(m_SourceArray)
    
        If Variable(m_SourceArray(Idx)).Equals(Value) Then
        
            Result = True
            Exit For
            
        End If
        
    Next
        
    If Not Result Then GoTo Catch
    
Finally:
        
    Index = Idx
    Find = Result
    
    Exit Function
    
Catch:
    
    Idx = -1
    Result = False
    
    Resume Finally
    
End Function

Public Property Get LowerBound() As Long
    
    If Not m_IsInitialized Then
    
        Err.Raise vbObjectError + 2, Typename(Me), "Object not initialized"
        Exit Sub
        
    End If

    LowerBound = LBound(m_SourceArray)

End Property

Public Property Get UpperBound() As Long
    
    If Not m_IsInitialized Then
    
        Err.Raise vbObjectError + 2, Typename(Me), "Object not initialized"
        Exit Sub
        
    End If
    
    UpperBound = UBound(m_SourceArray)

End Property

Public Property Get Count() As Long
    
    If Not m_IsInitialized Then
    
        Err.Raise vbObjectError + 2, Typename(Me), "Object not initialized"
        Exit Sub
        
    End If
    
    Count = UpperBound - LowerBound + 1

End Property

Public Sub EraseArray()

    Erase m_SourceArray
    m_IsInitialized = Not (API.pArrPtr(m_SourceArray) = 0)
    
End Sub

Public Function Join(Optional Delimiter) As String

    If Not m_IsInitialized Then
    
        Err.Raise vbObjectError + 2, Typename(Me), "Object not initialized"
        Exit Function
        
    End If

    Join = Strings.Join(m_SourceArray, Delimiter)
    
End Function

