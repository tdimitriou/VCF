Attribute VB_Name = "modPhase0CleanBench"
Option Explicit

Private Const LOG_FILE As String = "Phase0Clean_bench.log"

' Framework-only teardown. No KeepAlive, no manual DetachBindingsTree for cleanup.
' Parent live controls under Shell.Base.Children, then:
'   Shell.Base.Unload
'   Set Shell = Nothing
'   Set AppHost = Nothing
'   VCF.ClearApplication

Private Sub EndSession(ByRef Shell As Phase0Shell, ByRef AppHost As Phase0App)
    On Error Resume Next
    Debug.Print "EndSession: before Unload"
    If Not Shell Is Nothing Then
        If Not Shell.Base Is Nothing Then Shell.Base.Unload
    End If
    Debug.Print "EndSession: after Unload"
    Set Shell = Nothing
    Debug.Print "EndSession: after Set Shell"
    Set AppHost = Nothing
    VCF.ClearApplication
    Debug.Print "EndSession: done"
    Err.Clear
End Sub

Private Sub BeginSession(ByRef Shell As Phase0Shell, ByRef AppHost As Phase0App)
    VCF.ClearApplication
    Set AppHost = New Phase0App
    Set Shell = New Phase0Shell
    If Shell Is Nothing Then Err.Raise vbObjectError, , "BeginSession: Phase0Shell is Nothing"
    If Shell.Base Is Nothing Then Err.Raise vbObjectError, , "BeginSession: Base is Nothing"
End Sub

Private Sub Park(ByVal Win As Window, ByVal Obj As Object)
    Dim El As IUIElement
    If Obj Is Nothing Or Win Is Nothing Then Exit Sub
    If Win.Children.Contains(Obj) Then Exit Sub
    If TypeOf Obj Is IUIElement Then
        Set El = Obj
        If Not El.Parent Is Nothing Then Exit Sub
    End If
    Win.Children.Add Obj
End Sub

Public Sub RunAll()
    Dim Failed As Long
    Failed = 0

    Debug.Print "=== Demac.VCF Phase0Clean (framework teardown) ==="
    ClearLog

    If Not Bench_GoldenXamlLoad() Then Failed = Failed + 1
    If Not Bench_BindingOneWay() Then Failed = Failed + 1
    If Not Bench_ItemsPanelUniformGrid() Then Failed = Failed + 1
    If Not Bench_WindowChrome() Then Failed = Failed + 1
    If Not Bench_ViewNavUnload() Then Failed = Failed + 1

    Debug.Print "=== Done: " & (5 - Failed) & " passed, " & Failed & " failed ==="
    If Failed > 0 Then
        MsgBox Failed & " Phase0Clean test(s) failed. See Immediate / " & LOG_FILE, vbExclamation, "Phase0Clean"
    Else
        MsgBox "All Phase0Clean seed tests passed." & vbCrLf & vbCrLf & _
               "IDE stays in run (idle). Use Task Manager if Run/End freezes" & vbCrLf & _
               "(compiled VCF teardown — known).", vbInformation, "Phase0Clean"
    End If
    Debug.Print "Phase0Clean: after MsgBox"
    ' Holds drained inside Window.Unload before Form.Unload (VCF_Unload.log).
End Sub

Public Function Bench_GoldenXamlLoad() As Boolean
    Dim Reader As XAMLReader
    Dim Xml As String
    Dim Root As Object

    On Error GoTo Fail
    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\GoldenPanel.xml")
    Set Root = Reader.Load(Xml)
    If Root Is Nothing Then Err.Raise vbObjectError, , "GoldenPanel load returned Nothing"
    Set Root = Nothing
    Set Reader = Nothing
    LogResult "P0-GOLDEN", 0, "OK"
    Debug.Print "PASS  P0-GOLDEN GoldenPanel.xml"
    Bench_GoldenXamlLoad = True
    Exit Function
Fail:
    LogResult "P0-GOLDEN", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P0-GOLDEN - " & Err.Description
    Bench_GoldenXamlLoad = False
End Function

Public Function Bench_BindingOneWay() As Boolean
    Dim AppHost As Phase0App
    Dim Shell As Phase0Shell
    Dim Tb As TextBlock
    Dim Vm As Phase0ViewModel
    Dim B As Binding

    On Error GoTo Fail
    BeginSession Shell, AppHost

    Set Vm = New Phase0ViewModel
    Vm.Title = "Hello"
    Set Tb = New TextBlock
    Set B = New Binding
    Set B.TargetProperty = Tb.DependencyProperties.GetProperty("Text")
    Set B.Source = Tb.DependencyProperties.GetProperty("DataContext")
    B.Path = "Title"
    B.Mode = OneWay
    Set B.Target = Tb
    Tb.Bindings.Add B
    Set Tb.DataContext = Vm
    If Tb.Text <> "Hello" Then Err.Raise vbObjectError, , "Expected Hello, got " & Tb.Text

    Park Shell.Base, Tb
    Set Tb = Nothing
    Set B = Nothing
    Set Vm = Nothing

    EndSession Shell, AppHost
    LogResult "P4-ONEWAY", 0, "OK"
    Debug.Print "PASS  P4-ONEWAY Binding OneWay + Unload"
    Bench_BindingOneWay = True
    Exit Function
Fail:
    LogResult "P4-ONEWAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-ONEWAY - " & Err.Description
    EndSession Shell, AppHost
    Bench_BindingOneWay = False
End Function

Public Function Bench_ItemsPanelUniformGrid() As Boolean
    Dim AppHost As Phase0App
    Dim Shell As Phase0Shell
    Dim IC As ItemsControl
    Dim PanelTmpl As ItemsPanelTemplate
    Dim UgProto As UniformGrid
    Dim UgHost As UniformGrid
    Dim Reader As XAMLReader
    Dim Root As ItemsControl
    Dim UgXaml As UniformGrid
    Dim ZeroPad As Thickness
    Dim Coll As ObservableCollection
    Dim Tmpl As DataTemplate
    Dim TbTmpl As TextBlock
    Dim BtnTmpl As Button
    Dim OkItem As Phase0DialogButtonItem
    Dim CancelItem As Phase0DialogButtonItem
    Dim Tb0 As TextBlock
    Dim Tb1 As TextBlock
    Dim Btn0 As Button
    Dim Btn1 As Button
    Dim B As Binding

    On Error GoTo Fail
    BeginSession Shell, AppHost
    Debug.Print "P7c-PANEL enter"

    Set UgProto = New UniformGrid
    Set ZeroPad = New Thickness
    ZeroPad.Left = 0: ZeroPad.Top = 0: ZeroPad.Right = 0: ZeroPad.Bottom = 0
    UgProto.Widget.LockRefresh = True
    UgProto.Rows = 1
    UgProto.Columns = 3
    Set UgProto.Padding = ZeroPad
    UgProto.Widget.LockRefresh = False
    Set PanelTmpl = New ItemsPanelTemplate
    PanelTmpl.Children.Add UgProto
    Set IC = New ItemsControl
    IC.Widget.Move 0, 0, 200, 40
    Set IC.ItemsPanel = PanelTmpl
    If Not TypeOf IC.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "Code ItemsHost expected UniformGrid"
    Set UgHost = IC.ItemsHost
    If UgHost.Rows <> 1 Or UgHost.Columns <> 3 Then Err.Raise vbObjectError, , "Code host expected 1x3"
    Debug.Print "P7c-PANEL C0 OK"
    Park Shell.Base, IC
    Set IC = Nothing
    Set UgHost = Nothing
    Set UgProto = Nothing
    Set PanelTmpl = Nothing

    Set Reader = New XAMLReader
    Set Root = Reader.Load( _
        "<ItemsControl Width=""200"" Height=""40"">" & _
        "<ItemsControl.ItemsPanel><ItemsPanelTemplate>" & _
        "<UniformGrid Rows=""1"" Columns=""3"" Padding=""0""/>" & _
        "</ItemsPanelTemplate></ItemsControl.ItemsPanel>" & _
        "</ItemsControl>")
    If Root Is Nothing Then Err.Raise vbObjectError, , "XAML ItemsControl returned Nothing"
    Root.Widget.Move 0, 0, 200, 40
    If Not TypeOf Root.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "XAML ItemsHost expected UniformGrid"
    Set UgXaml = Root.ItemsHost
    If UgXaml.Rows <> 1 Or UgXaml.Columns <> 3 Then Err.Raise vbObjectError, , "XAML host expected 1x3"
    Debug.Print "P7c-PANEL C OK"
    Park Shell.Base, Root
    Set UgXaml = Nothing
    Set Root = Nothing
    Set Reader = Nothing

    Set OkItem = New Phase0DialogButtonItem
    OkItem.Text = "OK"
    Set CancelItem = New Phase0DialogButtonItem
    CancelItem.Text = "Cancel"
    Set Coll = New ObservableCollection
    Coll.Add OkItem
    Coll.Add CancelItem

    Set TbTmpl = New TextBlock
    Set B = New Binding
    Set B.TargetProperty = TbTmpl.DependencyProperties.GetProperty("Text")
    Set B.Source = TbTmpl.DependencyProperties.GetProperty("DataContext")
    B.Path = "Text"
    B.Mode = OneWay
    Set B.Target = TbTmpl
    TbTmpl.Bindings.Add B
    Set Tmpl = New DataTemplate
    Tmpl.Children.Add TbTmpl

    Set UgProto = New UniformGrid
    Set ZeroPad = New Thickness
    ZeroPad.Left = 0: ZeroPad.Top = 0: ZeroPad.Right = 0: ZeroPad.Bottom = 0
    UgProto.Widget.LockRefresh = True
    UgProto.Rows = 1
    UgProto.Columns = 2
    Set UgProto.Padding = ZeroPad
    UgProto.Widget.LockRefresh = False
    Set PanelTmpl = New ItemsPanelTemplate
    PanelTmpl.Children.Add UgProto
    Set IC = New ItemsControl
    IC.Widget.Move 0, 0, 200, 40
    Set IC.ItemsPanel = PanelTmpl
    Set IC.ItemTemplate = Tmpl
    Set IC.ItemsSource = Coll
    If Not TypeOf IC.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "T ItemsHost expected UniformGrid"
    Set UgHost = IC.ItemsHost
    If UgHost.Children.Count <> 2 Then Err.Raise vbObjectError, , "T expected 2 children"
    Set Tb0 = UgHost.Children(0)
    Set Tb1 = UgHost.Children(1)
    If Tb0.Text <> "OK" Then Err.Raise vbObjectError, , "T Tb0 expected OK"
    If Tb1.Text <> "Cancel" Then Err.Raise vbObjectError, , "T Tb1 expected Cancel"
    Debug.Print "P7c-PANEL T OK"
    Park Shell.Base, IC
    Set Tb0 = Nothing
    Set Tb1 = Nothing
    Set UgHost = Nothing
    Set IC = Nothing
    Set Tmpl = Nothing
    Set TbTmpl = Nothing
    Set PanelTmpl = Nothing
    Set UgProto = Nothing
    Set B = Nothing

    Set BtnTmpl = New Button
    BtnTmpl.Width = 80
    BtnTmpl.Height = 28
    Set B = New Binding
    Set B.TargetProperty = BtnTmpl.DependencyProperties.GetProperty("Content")
    Set B.Source = BtnTmpl.DependencyProperties.GetProperty("DataContext")
    B.Path = "Text"
    B.Mode = OneWay
    Set B.Target = BtnTmpl
    BtnTmpl.Bindings.Add B
    Set Tmpl = New DataTemplate
    Tmpl.Children.Add BtnTmpl
    Set UgProto = New UniformGrid
    Set ZeroPad = New Thickness
    ZeroPad.Left = 0: ZeroPad.Top = 0: ZeroPad.Right = 0: ZeroPad.Bottom = 0
    UgProto.Widget.LockRefresh = True
    UgProto.Rows = 1
    UgProto.Columns = 2
    Set UgProto.Padding = ZeroPad
    UgProto.Widget.LockRefresh = False
    Set PanelTmpl = New ItemsPanelTemplate
    PanelTmpl.Children.Add UgProto
    Set IC = New ItemsControl
    IC.Widget.Move 0, 0, 200, 40
    Set IC.ItemsPanel = PanelTmpl
    Set IC.ItemTemplate = Tmpl
    Set IC.ItemsSource = Coll
    If Not TypeOf IC.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "B ItemsHost expected UniformGrid"
    Set UgHost = IC.ItemsHost
    If UgHost.Children.Count <> 2 Then Err.Raise vbObjectError, , "B expected 2 children"
    Set Btn0 = UgHost.Children(0)
    Set Btn1 = UgHost.Children(1)
    If CStr(Btn0.Content) <> "OK" Then Err.Raise vbObjectError, , "B Btn0 Content expected OK"
    If CStr(Btn1.Content) <> "Cancel" Then Err.Raise vbObjectError, , "B Btn1 Content expected Cancel"
    Debug.Print "P7c-PANEL B OK"
    Park Shell.Base, IC
    Set Btn0 = Nothing
    Set Btn1 = Nothing
    Set UgHost = Nothing
    Set IC = Nothing
    Set Coll = Nothing
    Set OkItem = Nothing
    Set CancelItem = Nothing

    EndSession Shell, AppHost
    LogResult "P7c-PANEL", 0, "OK ItemsPanel + Unload"
    Debug.Print "PASS  P7c-PANEL ItemsPanelTemplate (framework Unload)"
    Bench_ItemsPanelUniformGrid = True
    Exit Function
Fail:
    LogResult "P7c-PANEL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7c-PANEL - " & Err.Description
    EndSession Shell, AppHost
    Bench_ItemsPanelUniformGrid = False
End Function

Public Function Bench_WindowChrome() As Boolean
    Dim AppHost As Phase0App
    Dim Borderless As Phase0Shell
    Dim Bordered As Phase0BorderedShell
    Dim Win As Window

    On Error GoTo Fail
    VCF.ClearApplication
    Set AppHost = New Phase0App

    Set Borderless = New Phase0Shell
    Set Win = Borderless.Base
    If CLng(Win.DependencyProperties.GetValue("BorderStyle")) <> 0 Then Err.Raise vbObjectError, , "Borderless DP expected 0"
    If Win.Form.BorderStyle <> 0 Then Err.Raise vbObjectError, , "Borderless Form.BorderStyle expected 0"
    Win.Unload
    Set Borderless = Nothing
    Set Win = Nothing

    Set Bordered = New Phase0BorderedShell
    Set Win = Bordered.Base
    If CLng(Win.DependencyProperties.GetValue("BorderStyle")) <> 2 Then Err.Raise vbObjectError, , "Bordered DP expected 2"
    If Win.Form.BorderStyle <> 2 Then Err.Raise vbObjectError, , "Bordered Form.BorderStyle expected 2"
    Win.Unload
    Set Bordered = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication

    LogResult "B-CHROME", 0, "OK"
    Debug.Print "PASS  B-CHROME first-frame BorderStyle sync"
    Bench_WindowChrome = True
    Exit Function
Fail:
    LogResult "B-CHROME", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-CHROME - " & Err.Description
    On Error Resume Next
    If Not Win Is Nothing Then Win.Unload
    Set Borderless = Nothing
    Set Bordered = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication
    Bench_WindowChrome = False
End Function

' Visibility swap + Unload; proves Windows registry clears without KeepAlive / manual detach cleanup.
Public Function Bench_ViewNavUnload() As Boolean
    Dim AppHost As Phase0App
    Dim Shell As Phase0Shell
    Dim Win As Window
    Dim ViewA As StackPanel
    Dim ViewB As StackPanel
    Dim TbA As TextBlock
    Dim TbB As TextBlock
    Dim VmA As Phase0ViewModel
    Dim VmB As Phase0ViewModel
    Dim B As Binding
    Dim i As Long
    Dim WinCount As Long

    On Error GoTo Fail
    BeginSession Shell, AppHost
    Set Win = Shell.Base

    Set VmA = New Phase0ViewModel
    VmA.Title = "ViewA"
    Set VmB = New Phase0ViewModel
    VmB.Title = "ViewB"

    Set TbA = New TextBlock
    Set B = New Binding
    Set B.TargetProperty = TbA.DependencyProperties.GetProperty("Text")
    Set B.Source = TbA.DependencyProperties.GetProperty("DataContext")
    B.Path = "Title"
    B.Mode = OneWay
    Set B.Target = TbA
    TbA.Bindings.Add B
    Set TbA.DataContext = VmA

    Set TbB = New TextBlock
    Set B = New Binding
    Set B.TargetProperty = TbB.DependencyProperties.GetProperty("Text")
    Set B.Source = TbB.DependencyProperties.GetProperty("DataContext")
    B.Path = "Title"
    B.Mode = OneWay
    Set B.Target = TbB
    TbB.Bindings.Add B
    Set TbB.DataContext = VmB

    Set ViewA = New StackPanel
    ViewA.Children.Add TbA
    Set ViewB = New StackPanel
    ViewB.Children.Add TbB
    Win.Children.Add ViewA
    Win.Children.Add ViewB
    ViewA.Visibility = VisibilityVisible
    ViewB.Visibility = VisibilityCollapsed

    For i = 1 To 20
        If (i Mod 2) = 0 Then
            ViewA.Visibility = VisibilityVisible
            ViewB.Visibility = VisibilityCollapsed
        Else
            ViewB.Visibility = VisibilityVisible
            ViewA.Visibility = VisibilityCollapsed
        End If
    Next

    Set TbA = Nothing
    Set TbB = Nothing
    Set ViewA = Nothing
    Set ViewB = Nothing
    Set VmA = Nothing
    Set VmB = Nothing
    Set B = Nothing

    Win.Unload
    Set Win = Nothing
    Set Shell = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication

    WinCount = 0
    If Not Application.Current Is Nothing Then
        If Not Application.Current.Base Is Nothing Then
            WinCount = Application.Current.Base.Windows.Count
        End If
    End If
    ' After ClearApplication, Current may be Nothing - that is OK.
    LogResult "B-NAV", 0, "OK Unload teardown"
    Debug.Print "PASS  B-NAV Visibility nav + Window.Unload"
    Bench_ViewNavUnload = True
    Exit Function
Fail:
    LogResult "B-NAV", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-NAV - " & Err.Description
    EndSession Shell, AppHost
    Bench_ViewNavUnload = False
End Function

Private Sub ClearLog()
    On Error Resume Next
    If Len(Dir(App.Path & "\" & LOG_FILE)) > 0 Then Kill App.Path & "\" & LOG_FILE
End Sub

Private Sub LogResult(ByVal Name As String, ByVal Ms As Long, ByVal Detail As String)
    Dim F As Integer
    On Error Resume Next
    F = FreeFile
    Open App.Path & "\" & LOG_FILE For Append As #F
    Print #F, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Name & vbTab & CStr(Ms) & vbTab & Detail
    Close #F
End Sub

Private Function LoadTextFile(ByVal Path As String) As String
    Dim F As Integer
    Dim Content As String
    F = FreeFile
    Open Path For Input As #F
    Content = Input$(LOF(F), F)
    Close #F
    LoadTextFile = Content
End Function