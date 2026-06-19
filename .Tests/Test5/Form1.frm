VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6225
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10695
   BeginProperty Font 
      Name            =   "Consolas"
      Size            =   8.25
      Charset         =   161
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   6225
   ScaleWidth      =   10695
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PreviewPane 
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   2655
      Left            =   1320
      ScaleHeight     =   2655
      ScaleWidth      =   3855
      TabIndex        =   0
      Top             =   720
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Click()
    PreviewPane.Cls
    
    PreviewPane.Print "CASTING TESTS"
    PreviewPane.Print "-------------"
    PreviewPane.Print
    
    PreviewPane.Print "Variable(5).Equals(""5"") : ", Variable(5).Equals("5")

    PreviewPane.Print "Variable(Date).Between(""2025-05-01"", ""20/05/2025"") : "; Variable(Date).Between("2025-05-01", "20/05/2025")
    
    PreviewPane.Print
    PreviewPane.Print
    
    PreviewPane.Print "OBJECT TESTS"
    PreviewPane.Print "------------"
    PreviewPane.Print
    
    Dim c As Collection
    Set c = New Collection
    
    PreviewPane.Print "Variable(c).Equals(c) : "; Variable(c).Equals(c)
    PreviewPane.Print "Variable(c).EqualOrLess(c) : "; Variable(c).EqualOrLess(c)
    
    PreviewPane.Print "Object.Equals(Variable(c).Value, c) : "; Object.Equals(Variable(c).Value, c)
    
    PreviewPane.Print "Variable(c) Is c : "; Variable(c) Is c
    PreviewPane.Print "Variable(c).Value Is c : "; Variable(c).Value Is c
    
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    PreviewPane.Move 240, 240, ScaleWidth - 240, ScaleHeight - 240
End Sub

Private Sub PreviewPane_Click()
    Form_Click
End Sub
