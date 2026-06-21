VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "ISelector"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get SelectedItem() As Object: End Property
Public Property Set SelectedItem(ByVal Value As Object): End Property
Public Property Get SelectedIndex() As Long: End Property
Public Property Let SelectedIndex(ByVal Value As Long): End Property
Public Property Get SelectedValue() As Variant: End Property
Public Property Let SelectedValue(ByVal Value As Variant): End Property
Public Property Get SelectedValuePath() As String: End Property
Public Property Let SelectedValuePath(ByVal Value As String): End Property
