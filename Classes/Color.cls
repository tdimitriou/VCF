VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Color"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public Type RGBStruct
    R   As Long
    G   As Long
    B   As Long
End Type

Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Function FromSystemColor(ByVal SysColor As Long) As Long ' Returns the RGB of a long colour value (System colour aware)
    If ((SysColor And &HFFFFFF00) = &H80000000) Then FromSystemColor = GetSysColor(SysColor And &HFF) Else FromSystemColor = SysColor
End Function

Public Function Multiply(ByVal Color As Long, ByVal Multiplier As Single) As Long
    Dim RGBStruct As RGBStruct
    
    RGBStruct = ToRGB(Color)
    
    With RGBStruct
        Multiply = RGB(.R * Multiplier, .G * Multiplier, .B * Multiplier)
    End With
End Function

Public Function FromHtml(ByVal HtmlColor As String) As Long
    HtmlColor = UCase$(Replace$(Trim$(HtmlColor), "#", ""))
    If Len(HtmlColor) < 6 Then HtmlColor = HtmlColor & String(6 - Len(HtmlColor), "0")
    
    FromHtml = CLng("&H" & Mid(HtmlColor, 5, 2) & Mid(HtmlColor, 3, 2) & Mid(HtmlColor, 1, 2))
End Function

Public Function ToHtml(ByVal Color As Long) As String
    Dim HexColor As String
    
    HexColor = Hex$(FromSystemColor(Color))
    If Len(HexColor) < 6 Then HexColor = String(6 - Len(HexColor), "0") & HexColor
    
    ToHtml = "#" & Mid(HexColor, 5, 2) & Mid(HexColor, 3, 2) & Mid(HexColor, 1, 2)
End Function

Public Function ToRGB(ByVal Color As Long) As RGBStruct
    Color = FromSystemColor(Color)
    With ToRGB
        .R = &HFF& And Color
        .G = (&HFF00& And Color) \ 256
        .B = (&HFF0000 And Color) \ 65536
    End With
End Function

Public Function FromRGB(Struct As RGBStruct) As Long
    With Struct
        FromRGB = RGB(.R, .G, .B)
    End With
End Function

Public Function Luminance(Struct As RGBStruct) As Integer
    With Struct
        Luminance = 0.2126 * .R + 0.7152 * .G + 0.0722 * .B
    End With
End Function
