VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DependencyPropertiesStatic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_PassCalls As Long
Private m_InheritCalls As Long

Public Sub ResetInheritanceCounters()
    m_PassCalls = 0
    m_InheritCalls = 0
End Sub

Public Property Get PassPropertyValueCalls() As Long
    PassPropertyValueCalls = m_PassCalls
End Property

Public Property Get InheritPropertyValuesCalls() As Long
    InheritPropertyValuesCalls = m_InheritCalls
End Property

' Phase 8b — push disabled. Call sites remain for compatibility.
Public Sub PassPropertyValue(ByVal Children As UIElementCollection, ByVal Source As DependencyProperty)
    If Source Is Nothing Then Exit Sub
    If Not Source.Metadata.IsInheritable Then Exit Sub
    ' Do not MarkInheritanceDirty here — every container DP handler calls Pass;
    ' dirty is set only when an inheritable value actually changes or attach is deferred.
End Sub

' On Parent attach: bindings were created before Parent existed, so refresh
' callbacks for unset inheritable DPs (no SetCurrentValue copy).
' During InheritanceBatch, defer to End → PropagateInheritableFrom (one tree wave).
Public Sub InheritPropertyValues(ByVal Target As IDependencyObject)
    Dim Prop As DependencyProperty
    Dim Inheritable As VCF.List
    Dim Item As Variant
    
    m_InheritCalls = m_InheritCalls + 1
    If Target Is Nothing Then Exit Sub
    If Target.Parent Is Nothing Then Exit Sub
    
    If modInheritanceBatch.IsInheritanceBatchActive Then
        modInheritanceBatch.MarkInheritanceDirty
        Exit Sub
    End If
    
    Set Inheritable = Target.DependencyProperties.InheritableProperties
    If Inheritable Is Nothing Then Exit Sub
    
    For Each Item In Inheritable
        Set Prop = Item
        If Prop Is Nothing Then GoTo NextInherit
        If Prop.HasExplicitValue Then GoTo NextInherit
        Prop.NotifyInheritedValueChanged
NextInherit:
    Next
End Sub

' After InheritanceBatch: one notify from root for inheritable DPs (lazy model).
Public Sub PropagateInheritableFrom(ByVal Root As IDependencyObject)
    Dim Prop As DependencyProperty
    Dim Inheritable As VCF.List
    Dim Item As Variant
    
    If Root Is Nothing Then Exit Sub
    
    Set Inheritable = Root.DependencyProperties.InheritableProperties
    If Inheritable Is Nothing Then Exit Sub
    
    For Each Item In Inheritable
        Set Prop = Item
        If Not Prop Is Nothing Then
            NotifyInheritableToDescendants Root, Prop
        End If
    Next
End Sub

' When an inheritable DP changes on Source, refresh bindings on descendants that
' still pull that value (no explicit local/current value).
Public Sub NotifyInheritableToDescendants(ByVal Source As IDependencyObject, ByVal ChangedProp As DependencyProperty)
    Dim Kids As Object
    Dim Children As UIElementCollection
    Dim Child As IDependencyObject
    Dim ChildProp As DependencyProperty
    
    If Source Is Nothing Then Exit Sub
    If ChangedProp Is Nothing Then Exit Sub
    If Not ChangedProp.Metadata.IsInheritable Then Exit Sub
    
    If modInheritanceBatch.IsInheritanceBatchActive Then
        modInheritanceBatch.MarkInheritanceDirty
        Exit Sub
    End If
    
    On Error Resume Next
    Set Kids = Source.Children
    Err.Clear
    On Error GoTo 0
    
    If Kids Is Nothing Then Exit Sub
    If Not TypeOf Kids Is UIElementCollection Then Exit Sub
    Set Children = Kids
    
    For Each Child In Children
        If Child Is Nothing Then GoTo NextChild
        If Not Child.DependencyProperties.Exists(ChangedProp.Name) Then GoTo NextChild
        
        Set ChildProp = Child.DependencyProperties.GetProperty(ChangedProp.Name)
        If ChildProp Is Nothing Then GoTo NextChild
        If Not ChildProp.Metadata.IsInheritable Then GoTo NextChild
        If ChildProp.PropertyType <> ChangedProp.PropertyType Then GoTo NextChild
        If ChildProp.PropertyTypeName <> ChangedProp.PropertyTypeName Then GoTo NextChild
        If ChildProp.ProgId <> ChangedProp.ProgId Then GoTo NextChild
        
        If ChildProp.HasExplicitValue Then GoTo NextChild
        
        ChildProp.NotifyInheritedValueChanged
        NotifyInheritableToDescendants Child, ChangedProp
NextChild:
    Next
End Sub
