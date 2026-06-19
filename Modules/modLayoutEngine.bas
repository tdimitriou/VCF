Attribute VB_Name = "modLayoutEngine"
Option Explicit

Public Type LayoutRect
    Left As Single
    Top As Single
    Width As Single
    Height As Single
End Type

Public Function LayoutRectFromDesign( _
    ByVal DesignLeft As Double, _
    ByVal DesignTop As Double, _
    ByVal DesignWidth As Double, _
    ByVal DesignHeight As Double, _
    ByVal HostWidth As Single, _
    ByVal HostHeight As Single, _
    ByVal HostDesignWidth As Double, _
    ByVal HostDesignHeight As Double) As LayoutRect

    Dim xFactor As Double
    Dim yFactor As Double

    If HostDesignWidth <= 0 Then HostDesignWidth = 1
    If HostDesignHeight <= 0 Then HostDesignHeight = 1

    xFactor = HostWidth / HostDesignWidth
    yFactor = HostHeight / HostDesignHeight

    With LayoutRectFromDesign
        .Left = CSng(DesignLeft * xFactor)
        .Top = CSng(DesignTop * yFactor)
        .Width = CSng(DesignWidth * xFactor)
        .Height = CSng(DesignHeight * yFactor)
    End With
End Function

Public Function LayoutRectFromMargin( _
    ByVal Margin As Thickness, _
    ByVal Width As Double, _
    ByVal Height As Double) As LayoutRect

    With LayoutRectFromMargin
        .Left = CSng(Margin.Left)
        .Top = CSng(Margin.Top)
        If Width > 0 Then
            .Width = CSng(Width)
        End If
        If Height > 0 Then
            .Height = CSng(Height)
        End If
    End With
End Function

Public Sub ApplyLayoutRectToElement(ByVal Element As IUIElement, ByVal R As LayoutRect)
    Element.Move R.Left, R.Top, R.Width, R.Height
End Sub

Public Function IsLayoutCollapsed(ByVal Value As Visibility) As Boolean
    IsLayoutCollapsed = (Value = VisibilityCollapsed)
End Function

Public Function MapDesignPropertyAlias(ByVal Dep As IDependencyObject, ByVal Name As String) As String
    MapDesignPropertyAlias = Name

    Select Case LCase$(Name)
        Case "designwidth"
            If Dep.DependencyProperties.Exists("Width") Then MapDesignPropertyAlias = "Width"
        Case "designheight"
            If Dep.DependencyProperties.Exists("Height") Then MapDesignPropertyAlias = "Height"
    End Select
End Function
