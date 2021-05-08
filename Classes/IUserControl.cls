VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IUserControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Property Get DependencyProperties() As DependencyProperties: End Property

Public Property Set DataContext(ByVal Value As Object): End Property
Public Property Get DataContext() As Object: End Property
Public Property Let DesignHeight(ByVal Value As Double): End Property
Public Property Get DesignHeight() As Double: End Property
Public Property Let DesignLeft(ByVal Value As Double): End Property
Public Property Get DesignLeft() As Double: End Property
Public Property Let DesignTop(ByVal Value As Double): End Property
Public Property Get DesignTop() As Double: End Property
Public Property Let DesignWidth(ByVal Value As Double): End Property
Public Property Get DesignWidth() As Double: End Property
Public Property Let Visibility(ByVal Value As Visibility): End Property
Public Property Get Visibility() As Visibility: End Property
Public Property Set Parent(ByVal Value As IControl): End Property
Public Property Get Parent() As IControl: End Property
Public Property Get Children() As UIElementCollection: End Property
Public Property Get Widget() As cWidgetBase: End Property
Public Property Get Widgets() As cWidgets: End Property
Public Property Let Name(ByVal Value As String): End Property
Public Property Get Name() As String: End Property
Public Property Get Base() As VCF.UserControl: End Property

Public Sub Move(ByRef Left As Single, ByRef Top As Single, ByRef Width As Single, ByRef Height As Single): End Sub
Public Sub InitializeComponent(): End Sub
