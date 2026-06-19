VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   15780
   LinkTopic       =   "Form1"
   ScaleHeight     =   8685
   ScaleWidth      =   15780
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents Root As cWidgetForm
Attribute Root.VB_VarHelpID = -1
Private Grid As UniformGrid
Private U As VCF.Border
'Dim G As UniformGrid

Private Sub Form_Load()
    ScaleMode = 3
    Set Root = Cairo.WidgetForms.CreateChild(Me.hWnd)
    Root.Move 0, 0, ScaleWidth, ScaleHeight
    Root.WidgetRoot.BackColor = VCF.Color.Multiply(vbBlue, 0.5)
    Set U = New VCF.Border
    
    Set Grid = New UniformGrid
    Dim Padding As Thickness
    Set Padding = New Thickness
    Padding.Bottom = 10
    Padding.Top = 10
    Padding.Right = 10
    Padding.Left = 10
    Grid.DependencyProperties.SetValue "Padding", Padding
    Grid.DependencyProperties.SetValue "ShowGridLines", True
    Root.Widgets.Add U, "_" & ObjPtr(U), 0, 0, Root.ScaleWidth, Root.ScaleHeight
    'Set G = New UniformGrid
    U.Children.Add Grid
    'G.Columns = 2
    'G.Rows = 2
    'G.Children.Add Grid
    'G.DependencyProperties.SetValue "ShowGridLines", True
    U.Widget.Alpha = 0
        
    LoadButtons
        
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    New_c.Timing True
    ScaleMode = 3
    Root.Move 0, 0, ScaleWidth, ScaleHeight
    
'    Dim Lst
'    Set Lst = New List
'    Dim v
'    For Each v In Grid.Children
'        Lst.Add v
'    Next
'    Grid.Children.Clear
    U.Widget.Move 0, 0, Root.ScaleWidth, Root.ScaleHeight
'    Grid.Children.AddRange Lst
    
    Root.WidgetRoot.Refresh
    
    'Debug.Print "Resize:"; New_c.Timing
End Sub

Public Sub LoadButtons()
    New_c.Timing True
    
    Static pp As Long
    pp = pp + 1
    
    Dim Idx As Long
    Static Buttons As VCF.List
    
    If Buttons Is Nothing Then Set Buttons = New VCF.List
    
    Static pq As Long
    'pq = IIf(pq = 0, 1, 0)
    
    Dim i As Long
    Dim j As Long
        
    Dim R As Long, C As Long, O As Long
    R = 7 - pq
    C = 5 + 2 * pq
        
    Grid.Widget.LockRefresh = True
    Grid.Rows = R
    Grid.Columns = C
            
    Dim Btn As VCF.Button
    For Idx = Buttons.Count To R * C - 1
        Set Btn = New VCF.Button
        Btn.CornerRadius = 8
        Btn.ClickMode = ClickModePress
        Btn.Children.Add New TextBlock
        Set Btn.Command = New ClickCommand
        With Btn.Children(0)
            .HorizontalAlignment = 2
            .VerticalAlignment = 2
            .ScaleFont = False
        End With
        Buttons.Add Btn
    Next
        
    'Debug.Print "Load Buttons Interm 0:"; New_c.Timing
    
    Dim Children As VCF.List
    Set Children = New VCF.List
    
    For i = 0 To R - 1
        For j = 0 To C - 1
            Idx = i * C + j
                                                                                                                                       
            Set Btn = Buttons(Idx)
            Btn.CommandParameter = Idx
            If Idx Mod 4 = 0 Then Btn.Widget.Visible = False
            Btn.Children(0).Text = "Button " & Idx + pp
            Children.Add Btn
            
        Next
    Next
    
    'Debug.Print "Load Buttons Interm 1:"; New_c.Timing
    
    Grid.Children.Clear
    'Debug.Print "Load Buttons Interm 2:"; New_c.Timing
    Grid.Children.AddRange Children
    Grid.Widget.LockRefresh = False
    
    'Debug.Print "Load Buttons:"; New_c.Timing
End Sub
