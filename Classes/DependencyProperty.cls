VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependencyProperty"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements INotifyPropertyChanged

Private m_Name                      As String
Private m_PropertyType              As VbVarType
Private m_PropertyTypeName          As String
Private m_ProgId                    As String
Private m_UnsetValue                As Variant
Private m_Value                     As Variant
Private m_ParentPtr                 As Long
Private m_CurrentValue              As Variant
Private m_Metadata                  As DependencyPropertyMetadata
Private m_Listeners                 As List
Private m_Callbacks                 As List
Private m_PropertyChangedEvent As PropertyChangedEvent

'#Region Public Properties
    Public Property Get Name() As String
        Name = m_Name
    End Property
    
    Public Property Get PropertyType() As VbVarType
        PropertyType = m_PropertyType
    End Property
    
    Public Property Get PropertyTypeName() As String
        PropertyTypeName = m_PropertyTypeName
    End Property
    
    Public Property Get ProgId() As String
        ProgId = m_ProgId
    End Property
    
    Public Property Get UnsetValue() As Variant
        Call API.CopyVariable(m_UnsetValue, UnsetValue)
    End Property
    
    Public Property Get Metadata() As DependencyPropertyMetadata
        Set Metadata = m_Metadata
    End Property
'#End Region

'#Region Friend Members
    Friend Sub AddListener(ByVal Value As IDependencyPropertyCallbackListener)
        Dim ValuePtr As Long
        
        ValuePtr = ObjPtr(Value)
        
        If m_Listeners.Contains(ValuePtr) Then Exit Sub
        
        m_Listeners.Add ValuePtr
    End Sub
    
    Friend Sub RemoveListener(ByVal ValuePtr As Long)
        If Not m_Listeners.Contains(ValuePtr) Then Exit Sub
        
        m_Listeners.Remove ValuePtr
    End Sub
    
    Friend Sub AddCallback(ByVal Value As IDependencyPropertyCallbackListener)
        Dim ValuePtr As Long
        
        ValuePtr = ObjPtr(Value)
        
        If m_Callbacks.Contains(ValuePtr) Then Exit Sub
        
        m_Callbacks.Add ValuePtr
    End Sub
    
    Friend Sub RemoveCallback(ByVal ValuePtr As Long)
        If Not m_Callbacks.Contains(ValuePtr) Then Exit Sub
        
        m_Callbacks.Remove ValuePtr
    End Sub
    
    Friend Property Get Parent() As DependencyProperties
        Set Parent = API.ObjFromPtr(m_ParentPtr)
    End Property
    
'CSEH: ErrMsgBox
    Friend Function GetValue() As Variant
        Dim Listener As IDependencyPropertyCallbackListener
        Dim ListenerPtr
        Dim ProvidedValue
        
        For Each ListenerPtr In m_Listeners
            Set Listener = API.ObjFromPtr(ListenerPtr)
            ProvidedValue = Empty
            If Not Listener Is Nothing Then
                If Listener.OnValueRequested(ProvidedValue) = True Then
                    API.CopyVariable ProvidedValue, m_Value
                End If
            End If
        Next
        
        If Object.Equals(m_Value, m_UnsetValue) Then
            Call API.CopyVariable(m_CurrentValue, GetValue)
        Else
            Call API.CopyVariable(m_Value, GetValue)
        End If
    End Function
    
    Friend Sub ClearValue()
        Call API.CopyVariable(m_UnsetValue, m_Value)
        Call API.CopyVariable(m_UnsetValue, m_CurrentValue)
    End Sub
    
    Friend Sub SetCurrentValue(Value)
        Dim OldEffectiveValue
        
        If Object.Equals(Value, m_CurrentValue) Then Exit Sub
        
        Call API.CopyVariable(GetValue, OldEffectiveValue)
        Call API.CopyVariable(Value, m_CurrentValue)
        
        ' No need to notify because the effective value has not changed
        If Object.Equals(GetValue, OldEffectiveValue) Then Exit Sub
        
        OnDependencyPropertyChanged OldEffectiveValue
    End Sub
    
'CSEH: ErrMsgBox
    Friend Sub SetValue(Value)
        Dim OldValue
        
        If Object.Equals(Value, m_Value) Then Exit Sub
        
        Call API.CopyVariable(m_Value, OldValue)
        Call API.CopyVariable(Value, m_Value)
                
        OnDependencyPropertyChanged OldValue
    End Sub
    
    Friend Sub Register(ByVal ParentPtr As Long, _
                        ByVal PropertyName As String, _
                        ByVal PropertyType As VbVarType, _
                        ByVal PropertyTypeName As String, _
                        ProgId As String, _
                        Optional UnsetValue, _
                        Optional Metadata As DependencyPropertyMetadata)
                        
        m_ParentPtr = ParentPtr
        m_Name = PropertyName
        m_PropertyType = PropertyType
        
        If m_PropertyType = vbObject Then
            Set m_UnsetValue = Nothing
            Set m_CurrentValue = Nothing
            Set m_Value = Nothing
        Else
            m_UnsetValue = Null
            m_CurrentValue = Null
            m_Value = Null
        End If
        
        m_PropertyTypeName = PropertyTypeName
        m_ProgId = ProgId
        
        If Not IsMissing(UnsetValue) Then
            Call API.CopyVariable(UnsetValue, m_UnsetValue)
            Call API.CopyVariable(UnsetValue, m_CurrentValue)
        End If
            
        If IsMissing(Metadata) Or (Metadata Is Nothing) Then
            Set m_Metadata = New DependencyPropertyMetadata
        Else
            Set m_Metadata = Metadata
        End If
    End Sub
'#End Region

'#Region Private Members
    Private Sub OnDependencyPropertyChanged(OldEffectiveValue)
        Call Parent.OnDependencyPropertyChanged(Me, OldEffectiveValue)
        
        m_PropertyChangedEvent.OnPropertyChanged m_Name
        
        Dim Callback As IDependencyPropertyCallbackListener
        Dim CallbackPtr
        
        For Each CallbackPtr In m_Callbacks
            Set Callback = API.ObjFromPtr(CallbackPtr)
            
            If Not Callback Is Nothing Then
                Callback.OnValueChanged Me, GetValue
            End If
        Next
        
    End Sub
'#End Region

Private Sub Class_Initialize()
    Set m_PropertyChangedEvent = New PropertyChangedEvent
    m_PropertyChangedEvent.Register Me
    
    Set m_Listeners = New List
    Set m_Callbacks = New List
End Sub

Private Sub Class_Terminate()
    m_PropertyChangedEvent.Unregister
End Sub

Private Function INotifyPropertyChanged_PropertyChangedEvent() As PropertyChangedEvent
    Set INotifyPropertyChanged_PropertyChangedEvent = PropertyChangedEvent
End Function

Public Function PropertyChangedEvent() As PropertyChangedEvent
    Set PropertyChangedEvent = m_PropertyChangedEvent
End Function

