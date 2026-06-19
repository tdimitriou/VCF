VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "DesignWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Implements VCF.IWindow

Private m_Win As VCF.Window

Private Sub Class_Initialize()
    On Error Resume Next
    ' Use Constructor to create Window
    Dim Constructor As VCF.Constructor
    Set Constructor = New VCF.Constructor
    Set m_Win = Constructor.NewWindow(Me)
    If Err Then Err.Clear
End Sub

Private Property Get IWindow_Base() As VCF.Window
    Set IWindow_Base = m_Win
End Property

Private Sub IWindow_InitializeComponent()
    ' Design-time initialization
    On Error Resume Next
    
    m_Win.DesignWidth = 800
    m_Win.DesignHeight = 600
    
    If Err Then Err.Clear
End Sub

Public Property Get Base() As VCF.Window
    Set Base = m_Win
End Property

