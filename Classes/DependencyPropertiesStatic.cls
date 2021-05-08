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

Public Sub PassPropertyValue(ByVal Children As UIElementCollection, ByVal Source As DependencyProperty)
    Dim Child As IDependencyObject
    
    If Not Source.Metadata.IsInheritable Then Exit Sub
    
    If Children Is Nothing Then Exit Sub
    
    For Each Child In Children
        PassPropertyValueToChild Child, Source
    Next
End Sub

Private Sub PassPropertyValueToChild(ByVal Child As IDependencyObject, ByVal Source As DependencyProperty)
    Dim Prop As DependencyProperty
    
    If Not Child.DependencyProperties.Exists(Source.Name) Then Exit Sub
    Set Prop = Child.DependencyProperties.GetProperty(Source.Name)
    
    If Not Prop.Metadata.IsInheritable Then Exit Sub
    If Prop.PropertyType <> Source.PropertyType Then Exit Sub
    If Prop.PropertyTypeName <> Source.PropertyTypeName Then Exit Sub
    If Prop.ProgId <> Source.ProgId Then Exit Sub
    
    Prop.SetCurrentValue Source.GetValue

End Sub

Public Sub InheritPropertyValues(ByVal Target As IDependencyObject)
    Dim Prop As DependencyProperty
    
    For Each Prop In Target.DependencyProperties.RegisteredProperties
        InheritPropertyValue Prop, Target.Parent
    Next
End Sub

Private Sub InheritPropertyValue(ByVal Prop As DependencyProperty, ByVal Parent As IDependencyObject)
    Dim Source As DependencyProperty
    
    If Not Prop.Metadata.IsInheritable Then Exit Sub
    If Parent Is Nothing Then Exit Sub
    If Not Parent.DependencyProperties.Exists(Prop.Name) Then Exit Sub
    
    Set Source = Parent.DependencyProperties.GetProperty(Prop.Name)
    
    If Prop.PropertyType <> Source.PropertyType Then Exit Sub
    If Prop.PropertyTypeName <> Source.PropertyTypeName Then Exit Sub
    If Prop.ProgId <> Source.ProgId Then Exit Sub
    
    Prop.SetCurrentValue Source.GetValue
End Sub
