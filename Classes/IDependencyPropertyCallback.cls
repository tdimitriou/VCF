VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IDependencyPropertyCallbackListener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

' The Idea behind the implementation
' ==================================
'
'===============
' BACKGROUND   |
'===============
' If the Source for a Binding Implements the INotifyPropertyChanged Interface,
' then the Binding Object updates the Target DependencyProperty each time the
' PropertyChangedEvent is fired for the PropertyName in question.
'
' So far so good...
'
'===============
' THE PROBLEM  |
'===============
' However, if the Source does not implement the INotifyPropertyChanged Interface
' we do not have a mechanism to inform the Target that the Value of the Property
' in question has changed. The Target Property will not get the updated Value
' even when it will request it manually (usually when the UI requests refresh)
' through the DependencyProperty.GetValue Method, because it is not aware of
' the Binding, and nobody has notified that the value has changed.
'
'===============
' THE SOLUTION |
'===============
' We need a mechanism to interfere to the DependencyProperty.GetValue Method, in
' order to inform the DependencyProperty that its Value has changed before the
' (possibly) obsolete Value will be passed to the Target Property of the object
' (usually IUIElement) that requested it.
'
' 1. The DependencyProperty keeps a List of Listener Objects that implement the
' IDependencyPropertyCallbackListener interface and asks each of them to provide
' its Value, in order to evaluate wheather the Value has changed outside the
' DependencyProperty.SetValue Method (normally through binding), before it will
' provide its Value to the consumer Object Property.
'
' 2. The Binding Class, in it's turn, implements the Interface and adds a
' Callback Listener to the DependencyProperty as soon as its  TargetProperty
' will be set.
'
' 3. As a result the DependencyProperty will ask the Binding Object to provide
' its Value and then will evaluate if the Value has changed, before the Value
' will be handed to the consumers.
'
' Note: The Binding will not Provide a Value and will set the OnValueRequested
' Function to false if the Source Implements the INotifyPropertyChanged Interface.
' This is happening in order to avoid extra overhead because the Binding will
' call the DependencyProperty.SetValue Method as soon as the NotifyPropertyChanged
' Event will be fired.

Public Function OnValueRequested(ByRef ProvidedValue) As Boolean: End Function

' Used instead of Event declaration in Bindings to avoid cross reference
Public Sub OnValueChanged(ByVal Sender As DependencyProperty, ByVal NewValue): End Sub
