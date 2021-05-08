Attribute VB_Name = "modWindowsFormsHostHelper"
Option Explicit

Private Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function ShowWindowAPI Lib "user32" Alias "ShowWindow" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetFocus Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function MoveWindow Lib "user32" (ByVal hWnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongW" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetActiveWindow Lib "user32" (ByVal hWnd As Long) As Long

Private Const GWL_STYLE = (-16)
Private Const GWL_EXSTYLE = (-20)

Private Const WS_CAPTION = &HC00000
Private Const WS_OVERLAPPED = &H0&
Private Const WS_CHILD = &H40000000
Private Const WS_TABSTOP = &H10000

Private Const WS_EX_MDICHILD = &H40
Private Const WS_EX_CONTROLPARENT = &H10000
Private Const WS_EX_NOACTIVATE = &H8000000

Private Const SW_HIDE = 0
Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOW = 5
Private Const SW_RESTORE = 9

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20

Private Function GetHandle(ByVal Object As Object)
    On Error Resume Next
    
    GetHandle = Object.hWnd
End Function

Public Sub SetChild(ByVal Child As Object, ByVal Host As WindowsFormsHost)
    Dim ChildHandle As Long
    Dim HostHandle As Long
    Dim ParentHandle As Long
    
    ChildHandle = GetHandle(Child)
    If ChildHandle = 0 Then Exit Sub
    
    If Host.Widget.Root Is Nothing Then Exit Sub
    ParentHandle = Host.Widget.Root.hWnd
    
    SetWindowLong ChildHandle, GWL_STYLE, WS_CHILD Or WS_TABSTOP Or WS_OVERLAPPED
    SetWindowLong ChildHandle, GWL_EXSTYLE, WS_EX_NOACTIVATE Or WS_EX_CONTROLPARENT
    SetParent ChildHandle, ParentHandle
    ShowWindowAPI ChildHandle, SW_HIDE
    SyncCords ChildHandle, Host
    ShowWindowAPI ChildHandle, SW_SHOW
    
    'SetActiveWindow ParentHandle
    'SetFocus ChildHandle
End Sub

Public Sub ShowWindow(ByVal Window As Object, ByVal Show As Boolean)
    On Error Resume Next
    
    Dim Handle As Long
    Dim ShowCommand As Long
    
    Handle = Window.hWnd
    If Handle = 0 Then Exit Sub
    ShowCommand = IIf(Show, SW_SHOW, SW_HIDE)
    
    ShowWindowAPI Handle, ShowCommand
End Sub

Private Sub SyncCords(ByVal ChildHandle As Long, ByVal Host As WindowsFormsHost)
    Dim W As cWidgetBase
    
    Set W = Host.Widget
    Dim Left As Double, Top As Double, Width As Double, Height As Double
    
    Left = W.AbsLeftPxl
    Top = W.AbsTopPxl
    Width = W.Width
    Height = W.Height
    
    MoveWindow ChildHandle, Left, Top, Width, Height, 1
End Sub
