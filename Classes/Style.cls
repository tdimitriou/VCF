VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Style"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IEnumerable

Event StyleChanged()

Private m_ChildStyles As cSortedDictionary

Private m_Key As String
Private m_TargetType As String

Private m_Setters As cCollection
Private m_Triggers As cCollection
Private IsInitialized As Boolean

Private m_BasedOn As VCF.Style
Attribute m_BasedOn.VB_VarHelpID = -1

Private WithEvents ThemesManager As ThemesManager
Attribute ThemesManager.VB_VarHelpID = -1

'#Region Constructor
    Public Function Initialize(ByVal TargetType As String, Optional ByVal Key As String, Optional BasedOn As Style)
        If IsInitialized Then
            Err.Raise vbObjectError, TypeName(Me), "Object already initialized"
            Exit Function
        End If
        
        If Len(TargetType) = 0 Then
            Err.Raise vbObjectError, TypeName(Me), "No Target Type defined"
            Exit Function
        End If
        
        If Not BasedOn Is Nothing And Len(Key) = 0 Then
            Err.Raise vbObjectError, TypeName(Me), "Cannot set BasedOn on Bas Styles(if not Key is exclusively defined)"
            Exit Function
        End If

        If Not BasedOn Is Nothing Then
            If BasedOn.TargetType <> TargetType Then
                Err.Raise vbObjectError, TypeName(Me), "Cannot inherint the Style of a different TargetType"
                Exit Function
            End If
        End If

        If Len(Key) And BasedOn Is Nothing Then Set BasedOn = GetBaseStyle(TargetType)
        
        m_TargetType = TargetType
        If Len(Key) = 0 Then Key = m_TargetType & ".BaseStyle"
        m_Key = Key
        
        IsInitialized = True
        
        If Not BasedOn Is Nothing Then
            Dim e
            Dim Setter As KeyValuePair
            For Each e In BasedOn
                Setter = e
                SetSetter Setter.Key, Empty 'Setter.Value
            Next
        End If
        
        Set m_BasedOn = BasedOn
    End Function
'#End Region

'#Region Friend Properties
'    Friend Property Get Setters() As cCollection
'        Set Setters = m_Setters
'    End Property
    
    Friend Sub AddChild(Child As VCF.Style)
        Dim Ptr As Long
        Ptr = ObjPtr(Child)
        If Not m_ChildStyles.Exists(Ptr) Then m_ChildStyles.Add Ptr, Ptr
    End Sub

    Friend Sub RemoveChild(ByVal Ptr As Long)
        If m_ChildStyles.Exists(Ptr) Then m_ChildStyles.Remove Ptr
    End Sub
    
    Friend Sub BasedOnStyleChanged()
        OnStyleChanged
    End Sub
'#End Region

'#Region Public Properties
    Public Property Get TargetType() As String
        VerifyInitialization
        TargetType = m_TargetType
    End Property
    
    Public Property Get Key() As String
        VerifyInitialization
        Key = m_Key
    End Property

    Public Property Get ItemAt(ByVal Index As Long) As KeyValuePair
        VerifyInitialization
        Dim e As KeyValuePair
        e = m_Setters.ItemByIndex(Index)
        
        With ItemAt
            .Key = e.Key
        
            If Not IsEmpty(e.Value) Then
                API.CopyVariable e.Value, .Value
            Else
                If Not m_BasedOn Is Nothing Then API.CopyVariable m_BasedOn.ItemAt(Index).Value, ItemAt.Value
            End If
        
        End With
    End Property
    
    Public Property Get KeyAt(ByVal Index As Long) As String
        VerifyInitialization
        KeyAt = m_Setters.KeyByIndex(Index)
    End Property
    
    Public Property Get BasedOn() As VCF.Style
        Set BasedOn = m_BasedOn
    End Property
'#End Region

'#Region Public Functions
    Public Function GetSetter(ByVal Property As String) As KeyValuePair
Attribute GetSetter.VB_UserMemId = 0
        VerifyInitialization
        
        On Error Resume Next
        
        If Not m_Setters.Exists(Property) Then
            Err.Raise vbObjectError, TypeName(Me), "Item not found"
            Exit Function
        End If
        
        'API.CopyVariable ItemAt(m_Setters.IndexByKey(Property)), GetSetter
        'Exit Function
        
        Dim e As KeyValuePair
        e = m_Setters(Property)

        GetSetter.Key = e.Key

        If Not IsEmpty(e.Value) Then
            API.CopyVariable e.Value, GetSetter.Value
        Else
            If Not m_BasedOn Is Nothing Then API.CopyVariable m_BasedOn.GetSetter(Property).Value, GetSetter.Value
        End If
    End Function
    
    Public Function SetSetter(ByVal Property As String, ByVal Value)
                
        VerifyInitialization
        
        Dim e As KeyValuePair
        
        If m_Setters.Exists(Property) Then
            ' Ignore if the Setter exists and has the same value
            e = m_Setters(Property)
            If Object.Equals(e.Value, Value) Then Exit Function
        End If
        
        e.Key = Property
        API.CopyVariable Value, e.Value
        
        m_Setters.Prop(Property) = e
        
        OnStyleChanged
    End Function
    
    Public Sub RemoveSetter(ByVal Property As String)
        VerifyInitialization
        m_Setters.Remove Property
        OnStyleChanged
    End Sub

    Public Sub AddTrigger(ByVal Trig As PropertyTrigger)
        VerifyInitialization
        If Trig Is Nothing Then Exit Sub
        m_Triggers.Add Trig
        OnStyleChanged
    End Sub

    Public Sub ClearTriggers()
        VerifyInitialization
        m_Triggers.RemoveAll
        OnStyleChanged
    End Sub

    Public Property Get TriggerCount() As Long
        VerifyInitialization
        TriggerCount = m_Triggers.Count
    End Property

    Public Property Get TriggerAt(ByVal Index As Long) As PropertyTrigger
        VerifyInitialization
        Set TriggerAt = m_Triggers.ItemByIndex(Index)
    End Property
    
    Public Sub Clear()
        VerifyInitialization
        m_Setters.RemoveAll
        OnStyleChanged
    End Sub
    
    Public Property Get Count() As Long
        VerifyInitialization
        Count = m_Setters.Count
    End Property
    
    Public Property Get MarkupValue(ByVal Key As String) As Variant
        With New MarkupExtensions
            API.CopyVariable .GetMarkupValue(GetSetter(Key).Value), MarkupValue
        End With
End Property
    
'#End Region

'#Region Enumerator
    Public Function GetEnumerator() As IUnknown
Attribute GetEnumerator.VB_UserMemId = -4
        VerifyInitialization
        Set GetEnumerator = New_c.EnumerateOn(Me, m_Setters.Count, 0)
    End Function
'#End Region

'#Region Class Events
    Private Sub Class_Initialize()
        Set m_ChildStyles = New_c.SortedDictionary
        Set m_Setters = New_c.Collection(False)
        Set m_Triggers = New_c.Collection(False)
        
        If Not Application.Current Is Nothing Then
            Set ThemesManager = Application.Current.Base.ThemesManager
        End If
    End Sub
'#End Region

Private Sub VerifyInitialization()
    If Not IsInitialized Then Err.Raise vbObjectError, TypeName(Me), "Object not initialized properly"
End Sub

Private Function GetBaseStyle(ByVal TargetType As String) As Style
    Dim Style

    If Application.Current Is Nothing Then Exit Function

    API.CopyVariable Application.Current.TryFindResource(TargetType & ".BaseStyle"), Style
    
    If Not IsObject(Style) Then Exit Function
    If TypeOf Style Is Style Then Set GetBaseStyle = Style
End Function

Private Sub OnStyleChanged()
    RaiseEvent StyleChanged
    
    Dim Ptr
    For Each Ptr In m_ChildStyles
        Dim Child As Style
        Set Child = API.ObjFromPtr(Ptr)
        If Not Child Is Nothing Then Child.BasedOnStyleChanged
    Next
End Sub

Private Function IEnumerable_NextItem(Idx As Long) As Variant
    API.CopyVariable ItemAt(Idx), IEnumerable_NextItem
End Function

Private Sub ThemesManager_ThemeCkanged()
    OnStyleChanged
End Sub
