VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "OverlayWidget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents W As cWidgetBase
Attribute W.VB_VarHelpID = -1
Public Alpha As Single
Public Color As Long
Public CornerRadious As Double
Public Visible As Boolean

Public Property Get Widget() As cWidgetBase
    Set Widget = W
End Property

Public Property Get Widgets() As cWidgets
    Set Widgets = W.Widgets
End Property

Private Sub Class_Initialize()
    Alpha = 0.25
    Set W = Cairo.WidgetBase
    W.InnerWidget = True
    W.ImplementsHitTest = True
End Sub

Private Sub W_ContainerResize()
    W.Move 0, 0, W.Parent.Width, W.Parent.Height
End Sub

Private Sub W_HitTest(ByVal x As Single, ByVal y As Single, HitResultHit As Boolean)
    HitResultHit = False
End Sub

Private Sub W_Paint(CC As vbRichClient5.cCairoContext, ByVal xAbs As Single, ByVal yAbs As Single, ByVal dx_Aligned As Single, ByVal dy_Aligned As Single, UserObj As Object)
    If Not Visible Then Exit Sub
    If Not W.Parent.Visible Then Exit Sub
    
    CC.RoundedRect 0, 0, CDbl(dx_Aligned), CDbl(dy_Aligned), CornerRadious
    CC.SetSourceColor Color, Alpha
    CC.Fill
End Sub


