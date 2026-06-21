Attribute VB_Name = "modSelectorEngine"
Option Explicit

Public Function ResolveSelectedValue(ByVal Item As Variant, ByVal SelectedValuePath As String) As Variant
    If IsEmpty(Item) Then Exit Function
    If IsNull(Item) Then Exit Function

    If Len(SelectedValuePath) = 0 Then
        ResolveSelectedValue = Item
        Exit Function
    End If

    On Error GoTo Fail

    If IsObject(Item) Then
        ResolveSelectedValue = CallByName(Item, SelectedValuePath, VbGet)
        Exit Function
    End If

Fail:
    ResolveSelectedValue = Item
End Function

Public Sub PublishSelectorState( _
    ByVal Props As DependencyProperties, _
    ByVal View As ListCollectionView, _
    ByVal SelectedValuePath As String, _
    ByRef Syncing As Boolean)

    Dim Idx As Long
    Dim Item As Variant
    Dim Val As Variant

    If Syncing Then Exit Sub
    Syncing = True

    If View Is Nothing Then
        Props.SetCurrentValue "SelectedIndex", -1
        Props.SetCurrentValue "SelectedItem", Nothing
        Props.SetCurrentValue "SelectedValue", Empty
        Syncing = False
        Exit Sub
    End If

    Idx = View.CurrentPosition
    Props.SetCurrentValue "SelectedIndex", Idx

    If Idx >= 0 And Idx < View.Count Then
        Call API.CopyVariable(View.CurrentItem, Item)
        If IsObject(Item) Then
            Props.SetCurrentValue "SelectedItem", Item
        Else
            Props.SetCurrentValue "SelectedItem", Nothing
        End If
        Val = ResolveSelectedValue(Item, SelectedValuePath)
        Props.SetCurrentValue "SelectedValue", Val
    Else
        Props.SetCurrentValue "SelectedItem", Nothing
        Props.SetCurrentValue "SelectedValue", Empty
    End If

    Syncing = False
End Sub

Public Sub ApplySelectorIndex( _
    ByVal View As ListCollectionView, _
    ByRef Syncing As Boolean, _
    ByVal Index As Long)

    If Syncing Then Exit Sub
    If View Is Nothing Then Exit Sub

    Syncing = True
    View.MoveCurrentToPosition Index
    Syncing = False
End Sub

Public Function FindItemIndex(ByVal Source As ObservableCollection, ByVal Value As Variant) As Long
    Dim i As Long
    Dim Item As Variant

    FindItemIndex = -1
    If Source Is Nothing Then Exit Function

    If IsObject(Value) Then
        FindItemIndex = Source.IndexOf(Value)
        Exit Function
    End If

    For i = 0 To Source.Count - 1
        Call API.CopyVariable(Source(i), Item)
        If Item = Value Then
            FindItemIndex = i
            Exit Function
        End If
    Next
End Function
