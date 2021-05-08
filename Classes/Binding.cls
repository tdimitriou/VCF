VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Binding"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Implements IMarkupExtension
Implements IDependencyPropertyCallbackListener

Public Enum BindingMode
    TwoWay = 0
    OneWay = 1
    OneTime = 2
    OneWayToSource = 3
    Default = 4
End Enum

Private m_Mode                                  As BindingMode
Private m_SourcePtr                             As Long             'Holds the pointer of the 'Source' (Object)
Private m_TargetPtr                             As Long             'Holds the pointer of the 'Target' (DependencyObject)
Private m_Path                                  As String           'Holds the value of the Path to the Source Property(String)
Private m_TargetPropertyPtr                     As Long             'Holds the pointer of the 'TargetProperty' (DependencyProperty)
Private m_Converter                             As IValueConverter  'Holds the value of the Converter (IValueConverter)
Private m_StringFormat                          As String           'Holds the value of the StringFormat (String)
Private m_HashCode                              As Long
Private m_SrcDepObjPtr                          As Long

Private WithEvents m_SourceProperty             As NestedProperty
Attribute m_SourceProperty.VB_VarHelpID = -1

Private WithEvents SourcePropertyChangedEvent   As PropertyChangedEvent
Attribute SourcePropertyChangedEvent.VB_VarHelpID = -1
Private WithEvents TargetPropertyChangedEvent   As PropertyChangedEvent
Attribute TargetPropertyChangedEvent.VB_VarHelpID = -1

' Use this to control the target update if Mode = OneTime (2)
Private IsTargetValueSet As Boolean

Friend Sub Initialize(ByVal Source As Object, _
                        ByVal Path As String, _
                        ByVal Target As IDependencyObject, _
                        ByVal TargetProperty As DependencyProperty, _
                        Optional ByVal Converter As IValueConverter, _
                        Optional ByVal StringFormat As String)
    
Try:
    On Error GoTo Catch
        
    Set Me.Source = Source
    Me.Path = Path
    Set Me.Target = Target
    Set Me.TargetProperty = TargetProperty
    Set Me.Converter = Converter
    Me.StringFormat = StringFormat
    
    GoTo SetValue
    
    Exit Sub
    
Catch:
    Err.Raise Err.Number, Err.Source & "." & TypeName(Me), Err.Description, Err.HelpFile, Err.HelpContext
    
    Exit Sub
    
SetValue:
    
    On Error Resume Next
    
    SetTargetPropertyValue
    
End Sub

Friend Property Get SrcDepObj() As DependencyProperty
    Set SrcDepObj = API.ObjFromPtr(m_SrcDepObjPtr)
End Property

Friend Property Set SrcDepObj(ByVal Value As DependencyProperty)
    ' Remove the Callback of the previous Source DependencyProperty if it has been set before
    If Not SrcDepObj Is Nothing Then SrcDepObj.RemoveCallback m_HashCode
    
    m_SrcDepObjPtr = ObjPtr(Value)
    
    If Value Is Nothing Then Exit Property
    
    Value.AddCallback Me
    
    Dim Src As Object
    Set Src = Value.GetValue
    
    If Src Is Nothing Then Exit Property
    
    Set Me.Source = Src
End Property

Public Property Get Mode() As BindingMode
    Mode = m_Mode
End Property

Public Property Let Mode(ByVal Value As BindingMode)
    m_Mode = Value
End Property

Public Property Get Source() As Object
    Set Source = API.ObjFromPtr(m_SourcePtr)
End Property

Public Property Set Source(ByVal Value As Object)
    Dim OldSource As Object
    
    Set SourcePropertyChangedEvent = Nothing
    
    If Not Value Is Nothing Then
        If TypeOf Value Is DependencyProperty Then
            Set SrcDepObj = Value
            Exit Property
        End If
    Else
        Set SrcDepObj = Value
    End If
        
    Set OldSource = Me.Source
    
    m_SourcePtr = ObjPtr(Value)
    
    If Value Is Nothing Then
        'If Not OldSource Is Nothing Then If Not TargetProperty Is Nothing Then TargetProperty.ClearValue
        Exit Property
    End If
    
    SetSourceProperty
    
    Dim Obj As INotifyPropertyChanged
    If Not TypeOf Value Is INotifyPropertyChanged Then Exit Property
    
    Set Obj = Value
    Set SourcePropertyChangedEvent = Obj.PropertyChangedEvent
    
    SetSourceProperty
    
    If Not Target Is Nothing Then SetTargetPropertyValue
End Property

Public Property Get Target() As IDependencyObject
    Set Target = API.ObjFromPtr(m_TargetPtr)
End Property

Public Property Set Target(ByVal Value As IDependencyObject)
    m_TargetPtr = ObjPtr(Value)
    
    If Not Value Is Nothing Then SetTargetPropertyValue
End Property

Public Property Get TargetProperty() As DependencyProperty
    Set TargetProperty = API.ObjFromPtr(m_TargetPropertyPtr)
End Property

Public Property Set TargetProperty(ByVal Value As DependencyProperty)
    ' Remove the Listener of the previous Target Property if it has been set before
    If Not TargetProperty Is Nothing Then TargetProperty.RemoveListener m_HashCode
    
    Set TargetPropertyChangedEvent = Nothing
    
    m_TargetPropertyPtr = ObjPtr(Value)
    
    If Value Is Nothing Then Exit Property
    
    Set TargetPropertyChangedEvent = Value.PropertyChangedEvent
    
    Value.AddListener Me
End Property

Public Property Get Path() As String
    Path = m_Path
End Property

Public Property Let Path(ByVal Value As String)
    m_Path = Value
    
    SetSourceProperty
End Property

Public Property Get Converter() As IValueConverter
    Set Converter = m_Converter
End Property

Public Property Set Converter(ByVal Value As IValueConverter)
    Set m_Converter = Value
End Property

Public Property Get StringFormat() As String
    StringFormat = m_StringFormat
End Property

Public Property Let StringFormat(ByVal Value As String)
    m_StringFormat = Value
End Property

Private Sub Class_Initialize()
    m_HashCode = ObjPtr(Me)
    m_Mode = VCF.BindingMode.Default
End Sub

Private Sub Class_Terminate()
    On Error Resume Next
    
    Dim Obj As DependencyProperty
    
    ' Remove Listener from TargetProperty
    Set Obj = API.ObjFromPtr(m_TargetPropertyPtr)
    If Not Obj Is Nothing Then Obj.RemoveListener m_HashCode
    
    ' Remove Callback from SrcDepObj
    Set Obj = API.ObjFromPtr(m_SrcDepObjPtr)
    If Not Obj Is Nothing Then Obj.RemoveCallback m_HashCode
    
End Sub

Private Sub IDependencyPropertyCallbackListener_OnValueChanged(ByVal Sender As DependencyProperty, ByVal NewValue As Variant)
    ' This will ocuur when the data context will change
    
    If Not Sender Is SrcDepObj Then Exit Sub
    If Not IsObject(NewValue) Then Exit Sub
    
    ' Unset the variable so the target can be updated again in the case that Mode = OneTime (2)
    ' Based on 'OneTime' Binding mode definition in .Net documentation
    IsTargetValueSet = False
    
    Set Me.Source = NewValue
    If NewValue Is Nothing Then Exit Sub

    SetTargetPropertyValue
End Sub

Private Function IDependencyPropertyCallbackListener_OnValueRequested(ProvidedValue As Variant) As Boolean
    On Error Resume Next
        
    ' No need to do something because the PropertyChangedEvent will take care
    If Not SourcePropertyChangedEvent Is Nothing Then Exit Function
    
    IDependencyPropertyCallbackListener_OnValueRequested = True
        
    Call API.CopyVariable(ProvideValue, ProvidedValue)
End Function

Private Function IMarkupExtension_ProvideValue() As Variant
    
    Call API.CopyVariable(ProvideValue, IMarkupExtension_ProvideValue)

End Function

'CSEH: ErrMsgBox
Public Function ProvideValue() As Variant
Try:
    On Error GoTo Catch
    
    If m_SourceProperty Is Nothing Then Exit Function
    
    Call API.CopyVariable(m_SourceProperty.GetValue, ProvideValue)

    If Not m_Converter Is Nothing Then Call API.CopyVariable(m_Converter.Convert(ProvideValue), ProvideValue)

    If IsObject(ProvideValue) Then Exit Function
        
    If Len(m_StringFormat) > 0 Then ProvideValue = Format$(ProvideValue, m_StringFormat)
    
    Exit Function
    
Catch:
    Debug.Print m_SourceProperty.Name & ": " & Err.Description
End Function

Private Sub SetTargetPropertyValue()
    Dim myMode As BindingMode
    
    If TargetProperty Is Nothing Then Exit Sub
    
    myMode = GetEffectiveMode
    
    If myMode = OneWayToSource Then Exit Sub
    If (myMode = OneTime) And IsTargetValueSet Then Exit Sub
    
    ' This will prevent further updates of the target (until the data context is changed)
    ' in the case when the Binding Mode is set to OneTime (2)
    IsTargetValueSet = True
    
    Call TargetProperty.SetValue(ProvideValue)
End Sub

Private Sub SourcePropertyChangedEvent_PropertyChanged(ByVal Sender As Object, e As Variant)
    If m_SourceProperty Is Nothing Then Exit Sub
    m_SourceProperty.OnPropertyChanged e
End Sub

Private Sub m_SourceProperty_ValueChanged(ByVal Sender As Object, e As Variant)
    SetTargetPropertyValue
End Sub

Private Sub SetSourceProperty()
    Set m_SourceProperty = Nothing
    
    If Len(Path) = 0 Then Exit Sub
    
    Set m_SourceProperty = New NestedProperty
    m_SourceProperty.Initialize Me, Me.Source, Path
End Sub

Private Sub TargetPropertyChangedEvent_PropertyChanged(ByVal Sender As Object, e As Variant)
    Dim myMode As BindingMode
    
    myMode = GetEffectiveMode
        
    If Not ((myMode = TwoWay) Or (myMode = OneWayToSource)) Then Exit Sub
    
    m_SourceProperty.SetValue Me.TargetProperty.GetValue
End Sub

Private Function GetEffectiveMode() As VCF.BindingMode
    GetEffectiveMode = m_Mode
    
    If Me.Target Is Nothing Then Exit Function
    
    If GetEffectiveMode = BindingMode.Default Then GetEffectiveMode = Me.TargetProperty.Metadata.BindingMode
End Function
