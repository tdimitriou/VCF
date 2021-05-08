VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ListCollectionView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Event CollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
Public Event CurrentChanged(Sender As Object, ByVal NewPosition As Long, ByVal OldPosition As Long)

Private m_SourcePtr             As Long
Attribute m_SourcePtr.VB_VarHelpID = -1
Private m_CurrentPosition       As Long

'#Region Constructors
    Public Sub Initialize(ByVal List As ObservableCollection)
        Static bIsInitialized As Boolean
        
        ' Allow initialization to occur only once
        If bIsInitialized Then
            Err.Raise vbObjectError, TypeName(Me), "Already Initialized"
            Exit Sub
        End If
        
        m_SourcePtr = List.GetHashCode
        m_CurrentPosition = -1
        
        bIsInitialized = True
    End Sub
'#End Region

'#Region Public Methods Inherited
    Public Sub Add(Value)
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Source.Add Value
    End Sub
    
    Public Sub Insert(Index, Value)
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Source.Insert Index, Value
    End Sub
    
    Public Sub Remove(Value)
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Source.Remove Value
    End Sub
    
    Public Sub RemoveAt(ByVal Index As Long)
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Source.RemoveAt Index
    End Sub
    
    Public Sub Clear()
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Source.Clear
    End Sub
    
    Public Function Contains(Value) As Boolean
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Contains = Source.Contains(Value)
    End Function
    
    Public Function IndexOf(Value) As Long
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        IndexOf = Source.IndexOf(Value)
    End Function
    
    Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Set GetEnumerator = Source.GetEnumerator
    End Function
'#End Region

'#Region Public Properties Inherited
    Public Property Get Count() As Long
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Count = Source.Count
    End Property
    
    Public Property Get Item(ByVal Index As Long) As Variant
Attribute Item.VB_UserMemId = 0
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        Call API.CopyVariable(Source(Index), Item)
    End Property
    
    Public Property Set Item(ByVal Index As Long, Value As Variant)
        If Source Is Nothing Then Err.Raise vbObjectError, TypeName(Me), "Not Initialized properly"
        
        If IsObject(Value) Then
            Set Source(Index) = Value
        Else
            Source(Index) = Value
        End If
    End Property
'#End Region
    
'#Region Public Properties
    Public Property Get Source() As ObservableCollection
        Set Source = API.ObjFromPtr(m_SourcePtr)
    End Property
    
    Public Property Get CurrentPosition() As Long
        CurrentPosition = m_CurrentPosition
    End Property
    
    Public Property Get CurrentItem()
        If m_CurrentPosition < 0 Or m_CurrentPosition >= Source.Count Then Exit Property
        API.CopyVariable Source(m_CurrentPosition), CurrentItem
    End Property
'#End Region
        
'#Region Public Methods
    Public Sub MoveCurrentTo(Value)
        If Source.Contains(Value) Then SetCurrentPosition Source.IndexOf(Value)
    End Sub
    
    Public Sub MoveCurrentToFirst()
        Dim OldPosition As Long
        OldPosition = m_CurrentPosition
                
        If Source.Count = 0 Then
            SetCurrentPosition -1
        Else
            SetCurrentPosition 0
        End If
    End Sub
    
    Public Sub MoveCurrentToLast()
        If Source.Count = 0 Then
            SetCurrentPosition -1
        Else
            SetCurrentPosition Source.Count - 1
        End If
    End Sub
    
    Public Sub MoveCurrentToNext()
        If Source.Count = 0 Then
            SetCurrentPosition -1
        Else
            If m_CurrentPosition = Source.Count - 1 Then Exit Sub
            SetCurrentPosition m_CurrentPosition + 1
        End If
    End Sub
    
    Public Sub MoveCurrentToPosition(ByVal Position As Long)
        If Source.Count = 0 Then
            SetCurrentPosition -1
            Exit Sub
        End If
        
        If Position < 0 Or Position > Source.Count - 1 Then
            ' Raise "Subscript out of range" error
            Err.Raise 9
            Exit Sub
        End If
        
        SetCurrentPosition Position
    End Sub

    Public Sub MoveCurrentToPrevious()
        If Source.Count = 0 Then
            SetCurrentPosition -1
        Else
            If m_CurrentPosition = 0 Then Exit Sub
            SetCurrentPosition m_CurrentPosition - 1
        End If
    End Sub
'#End Region
        
'#Region ObservableCollection Events
    Friend Sub OnSourceCollectionChanged(Sender As Object, e As CollectionChangedEventArgs)
        RaiseEvent CollectionChanged(Me, e)
    End Sub
'#End Region

'#Region Helpers
    Private Sub SetCurrentPosition(ByVal NewPosition As Long)
        If NewPosition = m_CurrentPosition Then Exit Sub
        
        Dim OldPosition As Long
        OldPosition = m_CurrentPosition
        
        m_CurrentPosition = NewPosition
        
        RaiseEvent CurrentChanged(Me, m_CurrentPosition, OldPosition)
    End Sub
'#End Region

