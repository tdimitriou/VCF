Attribute VB_Name = "modPhase0Bench"
Option Explicit

Private Const LOG_FILE As String = "Phase0_bench.log"

' Hold Button/ItemsControl trees for the IDE session ? releasing them (Terminate /
' WidgetForms.RemoveAll) hangs or silently crashes VB6 after a successful RunAll.
Private m_KeepAlive As Collection

Private Sub KeepAlive(ByVal Obj As Object)
    If Obj Is Nothing Then Exit Sub
    If m_KeepAlive Is Nothing Then Set m_KeepAlive = New Collection
    m_KeepAlive.Add Obj
End Sub

Public Sub RunAll()
    Dim Failed As Long
    Failed = 0

    Debug.Print "=== Demac.VCF Phase 0 benchmarks ==="
    ClearLog
    If m_KeepAlive Is Nothing Then Set m_KeepAlive = New Collection

    If Not Phase0Bench_GoldenXamlLoad() Then Failed = Failed + 1
    If Not Phase0Bench_CollectionAdd1000() Then Failed = Failed + 1
    If Not Phase0Bench_DualListCollectionView() Then Failed = Failed + 1
    If Not Phase0Bench_StrictMalformedXaml() Then Failed = Failed + 1
    If Not Phase0Bench_StrictUnknownType() Then Failed = Failed + 1
    If Not Phase1Bench_LayoutWidthXaml() Then Failed = Failed + 1
    If Not Phase1Bench_PanelVisibilityCollapsed() Then Failed = Failed + 1
    If Not Phase1Bench_VisibilityHiddenReserves() Then Failed = Failed + 1
    If Not Phase1Bench_VisibilityCollapsedReclaims() Then Failed = Failed + 1
    If Not Phase1Bench_VisibleBoolCollapsed() Then Failed = Failed + 1
    If Not Phase1Bench_TextBlockVisibility() Then Failed = Failed + 1
    If Not Phase1Bench_ImageVisibility() Then Failed = Failed + 1
    If Not Phase1Bench_IsHitTestVisible() Then Failed = Failed + 1
    If Not Phase1Bench_BorderWidthXaml() Then Failed = Failed + 1
    If Not Phase1Bench_BorderMeasure() Then Failed = Failed + 1
    If Not Phase1Bench_MinWidthFloor() Then Failed = Failed + 1
    If Not Phase1Bench_MaxWidthCeiling() Then Failed = Failed + 1
    If Not Phase2Bench_StackPanelXaml() Then Failed = Failed + 1
    If Not Phase2Bench_StackPanelLayout() Then Failed = Failed + 1
    If Not Phase2Bench_StackPanelMeasure() Then Failed = Failed + 1
    If Not Phase2Bench_MeasureOverrideAlias() Then Failed = Failed + 1
    If Not Phase2Bench_GridRowDefinitionsXaml() Then Failed = Failed + 1
    If Not Phase2Bench_GridAttachedCode() Then Failed = Failed + 1
    If Not Phase2Bench_GridAttachedXaml() Then Failed = Failed + 1
    If Not Phase2Bench_GridMeasure() Then Failed = Failed + 1
    If Not Phase2Bench_GridAlign() Then Failed = Failed + 1
    If Not Phase2Bench_GridAttachedDpBag() Then Failed = Failed + 1
    If Not Phase2Bench_DockPanelXaml() Then Failed = Failed + 1
    If Not Phase2Bench_DockPanelLayout() Then Failed = Failed + 1
    If Not Phase2Bench_CanvasXaml() Then Failed = Failed + 1
    If Not Phase2Bench_CanvasLayout() Then Failed = Failed + 1
    If Not Phase3Bench_MergedDictionaryLookup() Then Failed = Failed + 1
    If Not Phase3Bench_ResourceSourceLoad() Then Failed = Failed + 1
    If Not Phase3Bench_DynamicResourceExtension() Then Failed = Failed + 1
    If Not Phase3Bench_StrictUnknownProperty() Then Failed = Failed + 1
    If Not Phase4Bench_BindingOneWay() Then Failed = Failed + 1
    If Not Phase4Bench_BindingAttached() Then Failed = Failed + 1
    If Not Phase4Bench_BindingAttachedPath() Then Failed = Failed + 1
    If Not Phase4Bench_BindingAttachedLayout() Then Failed = Failed + 1
    If Not Phase4Bench_DataContextRebind() Then Failed = Failed + 1
    If Not Phase4Bench_BindingDetach() Then Failed = Failed + 1
    If Not Phase4Bench_DpPrecedence() Then Failed = Failed + 1
    If Not Phase4Bench_RelativeSourceSelf() Then Failed = Failed + 1
    If Not Phase4Bench_ElementName() Then Failed = Failed + 1
    If Not Phase4Bench_RelativeSourceTemplatedParent() Then Failed = Failed + 1
    If Not Phase4Bench_ElementNameCommand() Then Failed = Failed + 1
    If Not Phase4Bench_UpdateSourceTrigger() Then Failed = Failed + 1
    If Not Phase4Bench_UpdateSourceDelay() Then Failed = Failed + 1
    If Not Phase4Bench_TextCaretPreserve() Then Failed = Failed + 1
    If Not Phase4Bench_CanExecuteChanged() Then Failed = Failed + 1
    If Not Phase4bBench_BeginUpdateDefer() Then Failed = Failed + 1
    If Not Phase4bBench_Move() Then Failed = Failed + 1
    If Not Phase4bBench_ItemsControl() Then Failed = Failed + 1
    If Not Phase4dBench_Selector() Then Failed = Failed + 1
    If Not Phase4dBench_SelectionChanged() Then Failed = Failed + 1
    If Not Phase5aBench_OwnerDrawListView() Then Failed = Failed + 1
    If Not Phase5bBench_MeasureRow() Then Failed = Failed + 1
    If Not Phase5cBench_RowLevel() Then Failed = Failed + 1
    If Not Phase6aBench_ButtonContent() Then Failed = Failed + 1
    If Not Phase6bBench_PropertyTrigger() Then Failed = Failed + 1
    If Not Phase6cBench_ControlTemplate() Then Failed = Failed + 1
    If Not Phase6dBench_RenderCoalesce() Then Failed = Failed + 1
    If Not Phase6eBench_ContentPresenter() Then Failed = Failed + 1
    If Not Phase6eBench_ContentAlignment() Then Failed = Failed + 1
    If Not Phase6eBench_ContentControlContent() Then Failed = Failed + 1
    If Not Phase6eBench_ContentControlMeasure() Then Failed = Failed + 1
    If Not Phase6eBench_ContentTemplate() Then Failed = Failed + 1
    If Not Phase6fBench_TemplateBindingSlot() Then Failed = Failed + 1
    If Not Phase6gBench_LiveTemplateChrome() Then Failed = Failed + 1
    If Not Phase6hBench_LiveContentPresenter() Then Failed = Failed + 1
    If Not Phase6iBench_NestedContentPresenter() Then Failed = Failed + 1
    If Not Phase6jBench_MultiNodeTemplate() Then Failed = Failed + 1
    If Not Phase6kBench_TemplateBindingMarkup() Then Failed = Failed + 1
    If Not Phase2aBench_ThemeDictionarySwap() Then Failed = Failed + 1
    If Not Phase2aBench_SystemThemeResolve() Then Failed = Failed + 1
    If Not Phase7aBench_PosSalesOrderShell() Then Failed = Failed + 1
    If Not Phase7cBench_LegacyLayoutShim() Then Failed = Failed + 1
    If Not Phase7dBench_PanelResize() Then Failed = Failed + 1
    If Not Phase8Bench_InheritanceBatch() Then Failed = Failed + 1
    If Not Phase2aBench_NestedUniformGridResize() Then Failed = Failed + 1
    If Not Phase2aBench_ViewNavLeak() Then Failed = Failed + 1
    If Not Phase2aBench_WindowChrome() Then Failed = Failed + 1
    If Not Phase2aBench_ListViewBindHotspot() Then Failed = Failed + 1
    If Not Phase2aBench_ListViewPaddingDefaults() Then Failed = Failed + 1
    If Not Phase2aBench_TextBoxButtonPaddingDefaults() Then Failed = Failed + 1
    If Not Phase2aBench_UniformGridPaddingDefault() Then Failed = Failed + 1
    If Not Phase2aBench_UniformGridMeasure() Then Failed = Failed + 1
    If Not Phase7cBench_DialogDataTemplate() Then Failed = Failed + 1
    If Not Phase7cBench_ItemsPanelUniformGrid() Then Failed = Failed + 1

    ' Report only ? do not RemoveAll / release KeepAlive here (Button ItemsHost
    ' Terminate after MsgBox silently crashes the IDE).
    Debug.Print "=== Done: " & (89 - Failed) & " passed, " & Failed & " failed ==="
    If Failed > 0 Then
        MsgBox Failed & " Phase 0/1/2/3/4/5/6/7/2a test(s) failed. See Immediate window and " & LOG_FILE, vbExclamation, "Phase0"
    Else
        MsgBox "All Phase 0/1/2/3/4/5/6/7/2a tests passed.", vbInformation, "Phase0"
    End If
End Sub

Public Function Phase0Bench_GoldenXamlLoad() As Boolean
    Dim Reader As XAMLReader
    Dim Xml As String
    Dim Root As Object
    Dim Started As Single
    Dim ElapsedMs As Long

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\GoldenPanel.xml")

    Started = Timer
    Set Root = Reader.Load(Xml)
    ElapsedMs = CLng((Timer - Started) * 1000#)

    If Root Is Nothing Then Err.Raise vbObjectError, , "Golden XAML returned Nothing"

    LogResult "B-GOLD", ElapsedMs, "OK"
    Debug.Print "PASS  B-GOLD Golden XAML load (" & ElapsedMs & " ms)"
    Phase0Bench_GoldenXamlLoad = True
    Exit Function

Fail:
    LogResult "B-GOLD", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-GOLD ? " & Err.Description
    Phase0Bench_GoldenXamlLoad = False
End Function

Public Function Phase0Bench_CollectionAdd1000() As Boolean
    Dim Coll As ObservableCollection
    Dim i As Long
    Dim Started As Single
    Dim ElapsedMs As Long

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    Started = Timer
    For i = 1 To 1000
        Coll.Add "item" & i
    Next
    ElapsedMs = CLng((Timer - Started) * 1000#)

    If Coll.Count <> 1000 Then Err.Raise vbObjectError, , "Expected 1000 items"

    LogResult "B-COLL", ElapsedMs, "OK count=" & Coll.Count
    Debug.Print "PASS  B-COLL 1000x Add (" & ElapsedMs & " ms)"
    Phase0Bench_CollectionAdd1000 = True
    Exit Function

Fail:
    LogResult "B-COLL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-COLL ? " & Err.Description
    Phase0Bench_CollectionAdd1000 = False
End Function

Public Function Phase0Bench_DualListCollectionView() As Boolean
    Dim Coll1 As ObservableCollection
    Dim Coll2 As ObservableCollection
    Dim View1 As ListCollectionView
    Dim View2 As ListCollectionView

    On Error GoTo Fail

    Set Coll1 = New ObservableCollection
    Set Coll2 = New ObservableCollection
    Coll1.Add "a"
    Coll2.Add "b"

    Set View1 = VCF.CollectionViewSource.GetDefaultView(Coll1)
    Set View2 = VCF.CollectionViewSource.GetDefaultView(Coll2)

    If View1.Count <> 1 Or View2.Count <> 1 Then
        Err.Raise vbObjectError, , "Dual view counts wrong"
    End If

    LogResult "B-LCV", 0, "OK view1=" & View1.Count & " view2=" & View2.Count
    Debug.Print "PASS  B-LCV Dual ListCollectionView"
    Phase0Bench_DualListCollectionView = True
    Exit Function

Fail:
    LogResult "B-LCV", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-LCV ? " & Err.Description
    Phase0Bench_DualListCollectionView = False
End Function

Public Function Phase0Bench_StrictMalformedXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Object
    Dim SavedStrict As Boolean

    On Error GoTo Fail

    SavedStrict = VCF.StrictXamlLoad
    VCF.StrictXamlLoad = True

    Set Reader = New XAMLReader
    Set Root = Reader.Load("<Panel><Unclosed>")

    VCF.StrictXamlLoad = SavedStrict
    Err.Raise vbObjectError, , "Expected XamlLoadException for malformed XML"
    Exit Function

Fail:
    VCF.StrictXamlLoad = SavedStrict
    If Err.Source = "VCF.XamlLoadException" Then
        LogResult "B-STRICT-MALFORM", 0, "OK raised XamlLoadException"
        Debug.Print "PASS  B-STRICT Malformed XAML raises"
        Phase0Bench_StrictMalformedXaml = True
    Else
        LogResult "B-STRICT-MALFORM", 0, "FAIL: " & Err.Number & " " & Err.Description
        Debug.Print "FAIL  B-STRICT Malformed ? " & Err.Description
        Phase0Bench_StrictMalformedXaml = False
    End If
End Function

Public Function Phase0Bench_StrictUnknownType() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Object
    Dim SavedStrict As Boolean

    On Error GoTo Fail

    SavedStrict = VCF.StrictXamlLoad
    VCF.StrictXamlLoad = True

    Set Reader = New XAMLReader
    Set Root = Reader.Load("<NotARealVcfType/>")

    VCF.StrictXamlLoad = SavedStrict
    Err.Raise vbObjectError, , "Expected XamlLoadException for unknown type"
    Exit Function

Fail:
    VCF.StrictXamlLoad = SavedStrict
    If Err.Source = "VCF.XamlLoadException" Then
        LogResult "B-STRICT-UNKNOWN", 0, "OK raised XamlLoadException"
        Debug.Print "PASS  B-STRICT Unknown type raises"
        Phase0Bench_StrictUnknownType = True
    Else
        LogResult "B-STRICT-UNKNOWN", 0, "FAIL: " & Err.Number & " " & Err.Description
        Debug.Print "FAIL  B-STRICT Unknown ? " & Err.Description
        Phase0Bench_StrictUnknownType = False
    End If
End Function

Public Function Phase1Bench_LayoutWidthXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Panel
    Dim Xml As String

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\LayoutPanelWidth.xml")
    Set Root = Reader.Load(Xml)

    If Root Is Nothing Then Err.Raise vbObjectError, , "Layout XAML returned Nothing"
    If Root.Width <> 400# Then Err.Raise vbObjectError, , "Expected Width=400, got " & Root.Width
    If Root.Height <> 200# Then Err.Raise vbObjectError, , "Expected Height=200, got " & Root.Height

    LogResult "P1-WIDTH", 0, "OK Width=" & Root.Width & " Height=" & Root.Height
    Debug.Print "PASS  P1-WIDTH Layout Width/Height XAML"
    Phase1Bench_LayoutWidthXaml = True
    Exit Function

Fail:
    LogResult "P1-WIDTH", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-WIDTH ? " & Err.Description
    Phase1Bench_LayoutWidthXaml = False
End Function

Public Function Phase1Bench_PanelVisibilityCollapsed() As Boolean
    Dim P As Panel

    On Error GoTo Fail

    Set P = New Panel
    P.Visibility = VisibilityCollapsed

    If P.Visibility <> VisibilityCollapsed Then
        Err.Raise vbObjectError, , "Visibility DP not set to Collapsed"
    End If

    LogResult "P1-VIS", 0, "OK Collapsed stored"
    Debug.Print "PASS  P1-VIS Panel Visibility=Collapsed"
    Phase1Bench_PanelVisibilityCollapsed = True
    Exit Function

Fail:
    LogResult "P1-VIS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-VIS - " & Err.Description
    Phase1Bench_PanelVisibilityCollapsed = False
End Function

' Hidden keeps layout slot; Collapsed removes it (StackPanel).
Public Function Phase1Bench_VisibilityHiddenReserves() As Boolean
    Dim Sp As StackPanel
    Dim A As Panel
    Dim B As Panel
    Dim C As Panel

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Sp.Widget.Move 0, 0, 200, 300

    Set A = New Panel
    A.Width = 180
    A.Height = 20
    Set B = New Panel
    B.Width = 180
    B.Height = 20
    Set C = New Panel
    C.Width = 180
    C.Height = 20

    Sp.Children.Add A
    Sp.Children.Add B
    Sp.Children.Add C

    If Abs(C.Widget.Top - 40!) > 1! Then Err.Raise vbObjectError, , "Baseline C.Top expected 40, got " & C.Widget.Top

    B.Visibility = VisibilityHidden

    If Abs(C.Widget.Top - 40!) > 1! Then Err.Raise vbObjectError, , "Hidden must keep slot; C.Top expected 40, got " & C.Widget.Top
    If B.Widget.Visible Then Err.Raise vbObjectError, , "Hidden child Widget.Visible must be False"
    If Abs(B.Widget.Top - 20!) > 1! Then Err.Raise vbObjectError, , "Hidden child must still be arranged at Top=20"

    KeepAlive Sp
    LogResult "P1-VIS-HIDDEN", 0, "OK Hidden reserves slot C.Top=" & C.Widget.Top
    Debug.Print "PASS  P1-VIS-HIDDEN StackPanel Hidden reserves space"
    Phase1Bench_VisibilityHiddenReserves = True
    Exit Function

Fail:
    LogResult "P1-VIS-HIDDEN", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-VIS-HIDDEN - " & Err.Description
    Phase1Bench_VisibilityHiddenReserves = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

Public Function Phase1Bench_VisibilityCollapsedReclaims() As Boolean
    Dim Sp As StackPanel
    Dim A As Panel
    Dim B As Panel
    Dim C As Panel

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Sp.Widget.Move 0, 0, 200, 300

    Set A = New Panel
    A.Width = 180
    A.Height = 20
    Set B = New Panel
    B.Width = 180
    B.Height = 20
    Set C = New Panel
    C.Width = 180
    C.Height = 20

    Sp.Children.Add A
    Sp.Children.Add B
    Sp.Children.Add C

    B.Visibility = VisibilityCollapsed

    If Abs(C.Widget.Top - 20!) > 1! Then Err.Raise vbObjectError, , "Collapsed must reclaim; C.Top expected 20, got " & C.Widget.Top
    If Sp.Widget.Widgets.Exists("_" & ObjPtr(B)) Then Err.Raise vbObjectError, , "Collapsed child must be detached from parent Widgets"

    KeepAlive Sp
    LogResult "P1-VIS-COLLAPSE", 0, "OK Collapsed reclaims C.Top=" & C.Widget.Top
    Debug.Print "PASS  P1-VIS-COLLAPSE StackPanel Collapsed reclaims space"
    Phase1Bench_VisibilityCollapsedReclaims = True
    Exit Function

Fail:
    LogResult "P1-VIS-COLLAPSE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-VIS-COLLAPSE - " & Err.Description
    Phase1Bench_VisibilityCollapsedReclaims = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

' Button.Visible=False maps to VisibilityCollapsed (UniformGrid cell reclaim).
Public Function Phase1Bench_VisibleBoolCollapsed() As Boolean
    Dim Ug As UniformGrid
    Dim B1 As Button
    Dim B2 As Button
    Dim B3 As Button

    On Error GoTo Fail

    Set Ug = New UniformGrid
    Ug.Rows = 1
    Ug.Columns = 3
    Ug.Widget.Move 0, 0, 300, 40

    Set B1 = New Button
    B1.Content = "1"
    Set B2 = New Button
    B2.Content = "2"
    Set B3 = New Button
    B3.Content = "3"

    Ug.Children.Add B1
    Ug.Children.Add B2
    Ug.Children.Add B3

    If Abs(B3.Widget.Left - 200!) > 2! Then Err.Raise vbObjectError, , "Baseline B3.Left expected ~200, got " & B3.Widget.Left

    B2.Visible = False

    If B2.Visibility <> VisibilityCollapsed Then Err.Raise vbObjectError, , "Visible=False must map to VisibilityCollapsed, got " & B2.Visibility
    If Abs(B3.Widget.Left - 100!) > 2! Then Err.Raise vbObjectError, , "After Visible=False, B3.Left expected ~100, got " & B3.Widget.Left
    If Ug.Widget.Widgets.Exists("_" & ObjPtr(B2)) Then Err.Raise vbObjectError, , "Collapsed Button must leave UniformGrid Widgets"

    B2.Visible = True
    If B2.Visibility <> VisibilityVisible Then Err.Raise vbObjectError, , "Visible=True must restore VisibilityVisible"
    If Abs(B3.Widget.Left - 200!) > 2! Then Err.Raise vbObjectError, , "After Visible=True, B3.Left expected ~200, got " & B3.Widget.Left

    KeepAlive Ug
    LogResult "P1-VIS-BOOL", 0, "OK Visible bool maps to Collapsed cell reclaim"
    Debug.Print "PASS  P1-VIS-BOOL Button.Visible maps to Collapsed"
    Phase1Bench_VisibleBoolCollapsed = True
    Exit Function

Fail:
    LogResult "P1-VIS-BOOL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-VIS-BOOL - " & Err.Description
    Phase1Bench_VisibleBoolCollapsed = False
    On Error Resume Next
    KeepAlive Ug
    Err.Clear
End Function

' TextBlock Visibility Hidden reserves / Collapsed reclaims; Visible=False -> Collapsed.
Public Function Phase1Bench_TextBlockVisibility() As Boolean
    Dim Sp As StackPanel
    Dim A As TextBlock
    Dim B As TextBlock
    Dim C As TextBlock

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Sp.Widget.Move 0, 0, 200, 300

    Set A = New TextBlock
    A.Width = 180
    A.Height = 20
    A.Text = "A"
    Set B = New TextBlock
    B.Width = 180
    B.Height = 20
    B.Text = "B"
    Set C = New TextBlock
    C.Width = 180
    C.Height = 20
    C.Text = "C"

    Sp.Children.Add A
    Sp.Children.Add B
    Sp.Children.Add C

    If Abs(C.Widget.Top - 40!) > 1! Then Err.Raise vbObjectError, , "Baseline C.Top expected 40, got " & C.Widget.Top

    B.Visibility = VisibilityHidden
    If Abs(C.Widget.Top - 40!) > 1! Then Err.Raise vbObjectError, , "Hidden must keep slot; C.Top expected 40, got " & C.Widget.Top
    If B.Widget.Visible Then Err.Raise vbObjectError, , "Hidden TextBlock Widget.Visible must be False"

    B.Visibility = VisibilityVisible
    B.Visibility = VisibilityCollapsed
    If Abs(C.Widget.Top - 20!) > 1! Then Err.Raise vbObjectError, , "Collapsed must reclaim; C.Top expected 20, got " & C.Widget.Top
    If Sp.Widget.Widgets.Exists("_" & ObjPtr(B)) Then Err.Raise vbObjectError, , "Collapsed TextBlock must leave parent Widgets"

    B.Visible = True
    If B.Visibility <> VisibilityVisible Then Err.Raise vbObjectError, , "Visible=True must restore VisibilityVisible"
    B.Visible = False
    If B.Visibility <> VisibilityCollapsed Then Err.Raise vbObjectError, , "Visible=False must map to VisibilityCollapsed"
    If Abs(C.Widget.Top - 20!) > 1! Then Err.Raise vbObjectError, , "Visible=False must reclaim; C.Top expected 20, got " & C.Widget.Top

    KeepAlive Sp
    LogResult "P1-VIS-TB", 0, "OK TextBlock Hidden/Collapsed/Visible bool"
    Debug.Print "PASS  P1-VIS-TB TextBlock Visibility + Visible bool"
    Phase1Bench_TextBlockVisibility = True
    Exit Function

Fail:
    LogResult "P1-VIS-TB", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-VIS-TB - " & Err.Description
    Phase1Bench_TextBlockVisibility = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

' Image Visibility applies to widget; Visible=False -> Collapsed.
Public Function Phase1Bench_ImageVisibility() As Boolean
    Dim Img As VCF.Image

    On Error GoTo Fail

    Set Img = New VCF.Image
    If Img.Visibility <> VisibilityVisible Then Err.Raise vbObjectError, , "Default Visibility expected Visible"
    If Not Img.Widget.Visible Then Err.Raise vbObjectError, , "Default Widget.Visible expected True"

    Img.Visibility = VisibilityHidden
    If Img.Widget.Visible Then Err.Raise vbObjectError, , "Hidden Image Widget.Visible must be False"
    If Img.Visible Then Err.Raise vbObjectError, , "Hidden must sync Visible=False"

    Img.Visible = True
    If Img.Visibility <> VisibilityVisible Then Err.Raise vbObjectError, , "Visible=True must restore VisibilityVisible"
    If Not Img.Widget.Visible Then Err.Raise vbObjectError, , "Visible=True Widget.Visible expected True"

    Img.Visible = False
    If Img.Visibility <> VisibilityCollapsed Then Err.Raise vbObjectError, , "Visible=False must map to VisibilityCollapsed"
    If Img.Widget.Visible Then Err.Raise vbObjectError, , "Collapsed Image Widget.Visible must be False"

    KeepAlive Img
    LogResult "P1-VIS-IMG", 0, "OK Image Visibility + Visible bool"
    Debug.Print "PASS  P1-VIS-IMG Image Visibility + Visible bool"
    Phase1Bench_ImageVisibility = True
    Exit Function

Fail:
    LogResult "P1-VIS-IMG", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-VIS-IMG - " & Err.Description
    Phase1Bench_ImageVisibility = False
    On Error Resume Next
    KeepAlive Img
    Err.Clear
End Function

' IsHitTestVisible=False clears ImplementsHitTest (clicks pass through); still Visible.
Public Function Phase1Bench_IsHitTestVisible() As Boolean
    Dim B As Button
    Dim Tb As TextBlock

    On Error GoTo Fail

    Set B = New Button
    B.Content = "Hit"
    If Not B.IsHitTestVisible Then Err.Raise vbObjectError, , "Button default IsHitTestVisible expected True"
    If Not B.Widget.ImplementsHitTest Then Err.Raise vbObjectError, , "Button default ImplementsHitTest expected True"

    B.IsHitTestVisible = False
    If B.IsHitTestVisible Then Err.Raise vbObjectError, , "IsHitTestVisible=False did not stick"
    If B.Widget.ImplementsHitTest Then Err.Raise vbObjectError, , "ImplementsHitTest must be False when IsHitTestVisible=False"
    If Not B.Widget.Visible Then Err.Raise vbObjectError, , "IsHitTestVisible must not hide the widget"

    B.IsHitTestVisible = True
    If Not B.Widget.ImplementsHitTest Then Err.Raise vbObjectError, , "ImplementsHitTest must restore True"

    Set Tb = New TextBlock
    Tb.Text = "Pass"
    If Not Tb.IsHitTestVisible Then Err.Raise vbObjectError, , "TextBlock default IsHitTestVisible expected True"
    Tb.IsHitTestVisible = False
    If Tb.Widget.ImplementsHitTest Then Err.Raise vbObjectError, , "TextBlock ImplementsHitTest must be False"

    KeepAlive B
    KeepAlive Tb
    LogResult "P1-HITTEST", 0, "OK IsHitTestVisible toggles ImplementsHitTest"
    Debug.Print "PASS  P1-HITTEST IsHitTestVisible + ImplementsHitTest"
    Phase1Bench_IsHitTestVisible = True
    Exit Function

Fail:
    LogResult "P1-HITTEST", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-HITTEST - " & Err.Description
    Phase1Bench_IsHitTestVisible = False
    On Error Resume Next
    KeepAlive B
    KeepAlive Tb
    Err.Clear
End Function

Public Function Phase1Bench_BorderWidthXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Border
    Dim Xml As String

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\LayoutBorderWidth.xml")
    Set Root = Reader.Load(Xml)

    If Root Is Nothing Then Err.Raise vbObjectError, , "Border XAML returned Nothing"
    If Root.Width <> 320# Then Err.Raise vbObjectError, , "Expected Width=320, got " & Root.Width

    LogResult "P1-BORDER", 0, "OK Width=" & Root.Width
    Debug.Print "PASS  P1-BORDER Border Width XAML"
    Phase1Bench_BorderWidthXaml = True
    Exit Function

Fail:
    LogResult "P1-BORDER", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-BORDER - " & Err.Description
    Phase1Bench_BorderWidthXaml = False
End Function

Public Function Phase1Bench_BorderMeasure() As Boolean
    Dim B As Border
    Dim Child As Panel
    Dim Marg As Thickness

    On Error GoTo Fail

    Set B = New Border
    B.Width = 0
    B.Height = 0
    B.Widget.Move 0, 0, 200, 150

    Set Child = New Panel
    Child.Width = 80
    Child.Height = 40
    Set Marg = New Thickness
    Marg.Left = 10
    Marg.Top = 5
    Marg.Right = 10
    Marg.Bottom = 5
    Set Child.Margin = Marg

    Set B.Child = Child

    ' Child Desired 80x40 + Margin insets -> Border Desired 100x50
    If Abs(B.DesiredWidth - 100#) > 0.5 Then Err.Raise vbObjectError, , "DesiredWidth expected 100, got " & B.DesiredWidth
    If Abs(B.DesiredHeight - 50#) > 0.5 Then Err.Raise vbObjectError, , "DesiredHeight expected 50, got " & B.DesiredHeight
    If Abs(B.ActualWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualWidth expected 200, got " & B.ActualWidth
    If Abs(B.ActualHeight - 150#) > 0.5 Then Err.Raise vbObjectError, , "ActualHeight expected 150, got " & B.ActualHeight
    If Abs(Child.Widget.Left - 10!) > 1! Then Err.Raise vbObjectError, , "Child.Left expected 10, got " & Child.Widget.Left
    If Abs(Child.Widget.Top - 5!) > 1! Then Err.Raise vbObjectError, , "Child.Top expected 5, got " & Child.Widget.Top

    KeepAlive B
    LogResult "P1-BORDER-MEAS", 0, "OK Desired=" & B.DesiredWidth & "x" & B.DesiredHeight
    Debug.Print "PASS  P1-BORDER-MEAS Border Measure/Actual decorator child"
    Phase1Bench_BorderMeasure = True
    Exit Function

Fail:
    LogResult "P1-BORDER-MEAS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-BORDER-MEAS - " & Err.Description
    Phase1Bench_BorderMeasure = False
    On Error Resume Next
    KeepAlive B
    Err.Clear
End Function

Public Function Phase1Bench_MinWidthFloor() As Boolean
    Dim Sp As StackPanel
    Dim Child As Panel

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Sp.Widget.Move 0, 0, 300, 100

    Set Child = New Panel
    Child.Width = 50
    Child.Height = 20
    Child.MinWidth = 120
    Sp.Children.Add Child

    If Abs(Child.Widget.Width - 120!) > 1! Then
        Err.Raise vbObjectError, , "MinWidth floor expected Widget.Width=120, got " & Child.Widget.Width
    End If

    KeepAlive Sp
    LogResult "P1-MINMAX-MIN", 0, "OK MinWidth floors arranged width"
    Debug.Print "PASS  P1-MINMAX-MIN MinWidth floors arranged size"
    Phase1Bench_MinWidthFloor = True
    Exit Function

Fail:
    LogResult "P1-MINMAX-MIN", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-MINMAX-MIN - " & Err.Description
    Phase1Bench_MinWidthFloor = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

Public Function Phase1Bench_MaxWidthCeiling() As Boolean
    Dim Sp As StackPanel
    Dim Child As Panel

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Sp.Widget.Move 0, 0, 300, 100

    Set Child = New Panel
    Child.Width = 200
    Child.Height = 20
    Child.MaxWidth = 80
    Sp.Children.Add Child

    If Abs(Child.Widget.Width - 80!) > 1! Then
        Err.Raise vbObjectError, , "MaxWidth ceiling expected Widget.Width=80, got " & Child.Widget.Width
    End If

    KeepAlive Sp
    LogResult "P1-MINMAX-MAX", 0, "OK MaxWidth ceilings arranged width"
    Debug.Print "PASS  P1-MINMAX-MAX MaxWidth ceilings arranged size"
    Phase1Bench_MaxWidthCeiling = True
    Exit Function

Fail:
    LogResult "P1-MINMAX-MAX", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P1-MINMAX-MAX - " & Err.Description
    Phase1Bench_MaxWidthCeiling = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

Public Function Phase2Bench_StackPanelXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Object
    Dim Xml As String

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\LayoutStackPanel.xml")
    Set Root = Reader.Load(Xml)

    If Root Is Nothing Then Err.Raise vbObjectError, , "StackPanel XAML returned Nothing"
    If TypeName(Root) <> "StackPanel" Then Err.Raise vbObjectError, , "Expected StackPanel, got " & TypeName(Root)
    If CDbl(Root.Width) <> 240# Then Err.Raise vbObjectError, , "Expected Width=240, got " & Root.Width
    If CLng(Root.Orientation) <> OrientationVertical Then Err.Raise vbObjectError, , "Expected Vertical orientation"

    LogResult "P2-STACK", 0, "OK Width=" & Root.Width
    Debug.Print "PASS  P2-STACK StackPanel Width/Orientation XAML"
    Phase2Bench_StackPanelXaml = True
    Exit Function

Fail:
    LogResult "P2-STACK", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-STACK ? " & Err.Description
    Phase2Bench_StackPanelXaml = False
End Function

Public Function Phase2Bench_StackPanelLayout() As Boolean
    Dim Sp As Object
    Dim P1 As Panel
    Dim P2 As Panel

    On Error GoTo Fail

    Set Sp = CreateObject("VCF.StackPanel")
    Sp.Orientation = OrientationVertical
    Sp.Widget.Move 0, 0, 200, 300

    Set P1 = New Panel
    P1.Width = 180
    P1.Height = 50
    Set P2 = New Panel
    P2.Width = 180
    P2.Height = 80

    Sp.Children.Add P1
    Sp.Children.Add P2

    If Abs(P1.Widget.Top - 0!) > 1! Then Err.Raise vbObjectError, , "P1.Top expected 0, got " & P1.Widget.Top
    If Abs(P2.Widget.Top - 50!) > 1! Then Err.Raise vbObjectError, , "P2.Top expected 50, got " & P2.Widget.Top

    LogResult "P2-STACK-LAY", 0, "OK P2.Top=" & P2.Widget.Top
    Debug.Print "PASS  P2-STACK-LAY vertical stack positions"
    Phase2Bench_StackPanelLayout = True
    Exit Function

Fail:
    LogResult "P2-STACK-LAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-STACK-LAY ? " & Err.Description
    Phase2Bench_StackPanelLayout = False
End Function

Public Function Phase2Bench_StackPanelMeasure() As Boolean
    Dim Sp As StackPanel
    Dim P1 As Panel
    Dim P2 As Panel

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    ' Content-driven Desired*: unset Width/Height DPs (0) so MeasureLayout reports child sum.
    Sp.Width = 0
    Sp.Height = 0
    Sp.Widget.Move 0, 0, 200, 300

    Set P1 = New Panel
    P1.Width = 180
    P1.Height = 50
    Set P2 = New Panel
    P2.Width = 180
    P2.Height = 80

    Sp.Children.Add P1
    Sp.Children.Add P2

    If Abs(Sp.DesiredWidth - 180#) > 0.5 Then Err.Raise vbObjectError, , "DesiredWidth expected 180, got " & Sp.DesiredWidth
    If Abs(Sp.DesiredHeight - 130#) > 0.5 Then Err.Raise vbObjectError, , "DesiredHeight expected 130, got " & Sp.DesiredHeight
    If Abs(Sp.ActualWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualWidth expected 200, got " & Sp.ActualWidth
    If Abs(Sp.ActualHeight - 300#) > 0.5 Then Err.Raise vbObjectError, , "ActualHeight expected 300, got " & Sp.ActualHeight
    If Abs(P2.Widget.Top - 50!) > 1! Then Err.Raise vbObjectError, , "P2.Top expected 50, got " & P2.Widget.Top

    KeepAlive Sp
    LogResult "P2-STACK-MEAS", 0, "OK DesiredH=" & Sp.DesiredHeight & " ActualW=" & Sp.ActualWidth
    Debug.Print "PASS  P2-STACK-MEAS StackPanel Measure/Actual"
    Phase2Bench_StackPanelMeasure = True
    Exit Function

Fail:
    LogResult "P2-STACK-MEAS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-STACK-MEAS - " & Err.Description
    Phase2Bench_StackPanelMeasure = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

Public Function Phase2Bench_MeasureOverrideAlias() As Boolean
    Dim Sp As StackPanel
    Dim P1 As Panel
    Dim P2 As Panel
    Dim Dw As Double
    Dim Dh As Double

    On Error GoTo Fail

    Set Sp = New StackPanel
    Sp.Orientation = OrientationVertical
    Sp.Width = 0
    Sp.Height = 0

    Set P1 = New Panel
    P1.Width = 100
    P1.Height = 40
    Set P2 = New Panel
    P2.Width = 100
    P2.Height = 60

    Sp.Children.Add P1
    Sp.Children.Add P2

    ' Explicit MeasureOverride (WPF name / CallByName path).
    CallByName Sp, "MeasureOverride", VbMethod, 200#, 300#
    Dw = Sp.DesiredWidth
    Dh = Sp.DesiredHeight
    If Abs(Dw - 100#) > 0.5 Then Err.Raise vbObjectError, , "MeasureOverride DesiredWidth expected 100, got " & Dw
    If Abs(Dh - 100#) > 0.5 Then Err.Raise vbObjectError, , "MeasureOverride DesiredHeight expected 100, got " & Dh

    ' MeasureLayout alias must match.
    CallByName Sp, "MeasureLayout", VbMethod, 200#, 300#
    If Abs(Sp.DesiredWidth - Dw) > 0.01 Then Err.Raise vbObjectError, , "MeasureLayout DesiredWidth mismatch"
    If Abs(Sp.DesiredHeight - Dh) > 0.01 Then Err.Raise vbObjectError, , "MeasureLayout DesiredHeight mismatch"

    ' ArrangeOverride positions children.
    Sp.Widget.Move 0, 0, 200, 300
    Sp.ArrangeOverride 200#, 300#
    If Abs(P2.Widget.Top - 40!) > 1! Then Err.Raise vbObjectError, , "ArrangeOverride P2.Top expected 40, got " & P2.Widget.Top
    If Abs(Sp.ActualWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualWidth expected 200, got " & Sp.ActualWidth

    KeepAlive Sp
    LogResult "P2-MEAS-OVR", 0, "OK MeasureOverride/MeasureLayout alias + ArrangeOverride"
    Debug.Print "PASS  P2-MEAS-OVR MeasureOverride alias + ArrangeOverride"
    Phase2Bench_MeasureOverrideAlias = True
    Exit Function

Fail:
    LogResult "P2-MEAS-OVR", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-MEAS-OVR - " & Err.Description
    Phase2Bench_MeasureOverrideAlias = False
    On Error Resume Next
    KeepAlive Sp
    Err.Clear
End Function

Public Function Phase2Bench_GridRowDefinitionsXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Object
    Dim Xml As String

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\LayoutGridRows.xml")
    Set Root = Reader.Load(Xml)

    If Root Is Nothing Then Err.Raise vbObjectError, , "Grid XAML returned Nothing"
    If TypeName(Root) <> "Grid" Then Err.Raise vbObjectError, , "Expected Grid, got " & TypeName(Root)
    If Root.RowDefinitions.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 row definitions"
    If Root.ColumnDefinitions.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 column definitions"
    If CDbl(Root.Width) <> 300# Then Err.Raise vbObjectError, , "Expected Width=300"

    LogResult "P2-GRID", 0, "OK rows=" & Root.RowDefinitions.Count
    Debug.Print "PASS  P2-GRID Grid RowDefinitions/ColumnDefinitions XAML"
    Phase2Bench_GridRowDefinitionsXaml = True
    Exit Function

Fail:
    LogResult "P2-GRID", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-GRID - " & Err.Description
    Phase2Bench_GridRowDefinitionsXaml = False
End Function

Public Function Phase2Bench_GridAttachedCode() As Boolean
    Dim G As Grid
    Dim Rd As RowDefinition
    Dim Cd As ColumnDefinition
    Dim P00 As Panel
    Dim P01 As Panel
    Dim P10 As Panel
    Dim P11 As Panel

    On Error GoTo Fail

    Set G = New Grid
    G.Width = 200
    G.Height = 200

    Set Rd = New RowDefinition
    Rd.Height = "*"
    G.RowDefinitions.Add Rd
    Set Rd = New RowDefinition
    Rd.Height = "*"
    G.RowDefinitions.Add Rd

    Set Cd = New ColumnDefinition
    Cd.Width = "*"
    G.ColumnDefinitions.Add Cd
    Set Cd = New ColumnDefinition
    Cd.Width = "*"
    G.ColumnDefinitions.Add Cd

    Set P00 = New Panel
    P00.Width = 40
    P00.Height = 40
    Set P01 = New Panel
    P01.Width = 40
    P01.Height = 40
    Set P10 = New Panel
    P10.Width = 40
    P10.Height = 40
    Set P11 = New Panel
    P11.Width = 40
    P11.Height = 40

    G.SetRow P00, 0
    G.SetColumn P00, 0
    G.SetRow P01, 0
    G.SetColumn P01, 1
    G.SetRow P10, 1
    G.SetColumn P10, 0
    G.SetRow P11, 1
    G.SetColumn P11, 1

    If G.GetRow(P11) <> 1 Then Err.Raise vbObjectError, , "GetRow(P11) expected 1"
    If G.GetColumn(P01) <> 1 Then Err.Raise vbObjectError, , "GetColumn(P01) expected 1"

    G.Children.Add P00
    G.Children.Add P01
    G.Children.Add P10
    G.Children.Add P11
    G.Widget.Move 0, 0, 200, 200

    If Abs(P00.Widget.Left - 0!) > 2! Or Abs(P00.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "P00 expected (0,0), got (" & P00.Widget.Left & "," & P00.Widget.Top & ")"
    End If
    If Abs(P01.Widget.Left - 100!) > 2! Or Abs(P01.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "P01 expected (100,0), got (" & P01.Widget.Left & "," & P01.Widget.Top & ")"
    End If
    If Abs(P10.Widget.Left - 0!) > 2! Or Abs(P10.Widget.Top - 100!) > 2! Then
        Err.Raise vbObjectError, , "P10 expected (0,100), got (" & P10.Widget.Left & "," & P10.Widget.Top & ")"
    End If
    If Abs(P11.Widget.Left - 100!) > 2! Or Abs(P11.Widget.Top - 100!) > 2! Then
        Err.Raise vbObjectError, , "P11 expected (100,100), got (" & P11.Widget.Left & "," & P11.Widget.Top & ")"
    End If

    KeepAlive G
    LogResult "P2-GRID-ATTACH", 0, "OK Get/SetRow Column cell positions"
    Debug.Print "PASS  P2-GRID-ATTACH Grid.SetRow/SetColumn positions"
    Phase2Bench_GridAttachedCode = True
    Exit Function

Fail:
    LogResult "P2-GRID-ATTACH", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-GRID-ATTACH - " & Err.Description
    Phase2Bench_GridAttachedCode = False
    On Error Resume Next
    KeepAlive G
    Err.Clear
End Function

Public Function Phase2Bench_GridAttachedXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Grid
    Dim Xml As String
    Dim P00 As Panel
    Dim P01 As Panel
    Dim P10 As Panel
    Dim P11 As Panel

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Xml = LoadTextFile(App.Path & "\Resources\LayoutGridAttach.xml")
    Set Root = Reader.Load(Xml)

    If Root Is Nothing Then Err.Raise vbObjectError, , "Grid XAML returned Nothing"
    If Root.Children.Count <> 4 Then Err.Raise vbObjectError, , "Expected 4 children, got " & Root.Children.Count

    Set P00 = Root.Children(0)
    Set P01 = Root.Children(1)
    Set P10 = Root.Children(2)
    Set P11 = Root.Children(3)

    If Root.GetRow(P11) <> 1 Or Root.GetColumn(P11) <> 1 Then
        Err.Raise vbObjectError, , "XAML attached GetRow/GetColumn mismatch for P11"
    End If

    Root.Widget.Move 0, 0, 200, 200

    If Abs(P00.Widget.Left - 0!) > 2! Or Abs(P00.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "XAML P00 expected (0,0), got (" & P00.Widget.Left & "," & P00.Widget.Top & ")"
    End If
    If Abs(P01.Widget.Left - 100!) > 2! Or Abs(P01.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "XAML P01 expected (100,0), got (" & P01.Widget.Left & "," & P01.Widget.Top & ")"
    End If
    If Abs(P10.Widget.Left - 0!) > 2! Or Abs(P10.Widget.Top - 100!) > 2! Then
        Err.Raise vbObjectError, , "XAML P10 expected (0,100), got (" & P10.Widget.Left & "," & P10.Widget.Top & ")"
    End If
    If Abs(P11.Widget.Left - 100!) > 2! Or Abs(P11.Widget.Top - 100!) > 2! Then
        Err.Raise vbObjectError, , "XAML P11 expected (100,100), got (" & P11.Widget.Left & "," & P11.Widget.Top & ")"
    End If

    KeepAlive Root
    LogResult "P2-GRID-XAML", 0, "OK Grid.Row/Column XAML positions"
    Debug.Print "PASS  P2-GRID-XAML Grid.Row/Column XAML positions"
    Phase2Bench_GridAttachedXaml = True
    Exit Function

Fail:
    LogResult "P2-GRID-XAML", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-GRID-XAML - " & Err.Description
    Phase2Bench_GridAttachedXaml = False
    On Error Resume Next
    KeepAlive Root
    Err.Clear
End Function

Public Function Phase2Bench_GridMeasure() As Boolean
    Dim G As Grid
    Dim Rd As RowDefinition
    Dim Cd As ColumnDefinition
    Dim PAuto As Panel
    Dim PPixel As Panel

    On Error GoTo Fail

    Set G = New Grid
    ' Content-driven Desired*: unset Width/Height DPs so MeasureLayout reports track sum.
    G.Width = 0
    G.Height = 0
    G.Widget.Move 0, 0, 200, 200

    Set Rd = New RowDefinition
    Rd.Height = "Auto"
    G.RowDefinitions.Add Rd
    Set Rd = New RowDefinition
    Rd.Height = "80"
    G.RowDefinitions.Add Rd

    Set Cd = New ColumnDefinition
    Cd.Width = "*"
    G.ColumnDefinitions.Add Cd

    Set PAuto = New Panel
    PAuto.Width = 120
    PAuto.Height = 50
    Set PPixel = New Panel
    PPixel.Width = 120
    PPixel.Height = 40

    G.SetRow PAuto, 0
    G.SetColumn PAuto, 0
    G.SetRow PPixel, 1
    G.SetColumn PPixel, 0

    G.Children.Add PAuto
    G.Children.Add PPixel

    ' Auto(50) + Pixel(80) = 130; star column eats AvailableWidth=200
    If Abs(G.DesiredHeight - 130#) > 0.5 Then Err.Raise vbObjectError, , "DesiredHeight expected 130, got " & G.DesiredHeight
    If Abs(G.DesiredWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "DesiredWidth expected 200, got " & G.DesiredWidth
    If Abs(G.ActualWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualWidth expected 200, got " & G.ActualWidth
    If Abs(G.ActualHeight - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualHeight expected 200, got " & G.ActualHeight
    If Abs(PAuto.Widget.Top - 0!) > 1! Then Err.Raise vbObjectError, , "PAuto.Top expected 0, got " & PAuto.Widget.Top
    If Abs(PPixel.Widget.Top - 50!) > 1! Then Err.Raise vbObjectError, , "PPixel.Top expected 50, got " & PPixel.Widget.Top

    KeepAlive G
    LogResult "P2-GRID-MEAS", 0, "OK DesiredH=" & G.DesiredHeight & " ActualW=" & G.ActualWidth
    Debug.Print "PASS  P2-GRID-MEAS Grid Measure/Actual Auto+Pixel+Star"
    Phase2Bench_GridMeasure = True
    Exit Function

Fail:
    LogResult "P2-GRID-MEAS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-GRID-MEAS - " & Err.Description
    Phase2Bench_GridMeasure = False
    On Error Resume Next
    KeepAlive G
    Err.Clear
End Function

Public Function Phase2Bench_GridAlign() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Grid
    Dim Child As Panel

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\LayoutGridAlign.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "LayoutGridAlign returned Nothing"
    If Root.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 child, got " & Root.Children.Count

    Set Child = Root.Children(0)
    Root.Widget.Move 0, 0, 200, 200

    ' 40x40 centered in 200x200 => Left/Top ~80
    If Abs(Child.Widget.Left - 80!) > 2! Then Err.Raise vbObjectError, , "Left expected ~80, got " & Child.Widget.Left
    If Abs(Child.Widget.Top - 80!) > 2! Then Err.Raise vbObjectError, , "Top expected ~80, got " & Child.Widget.Top
    If Abs(Child.Widget.Width - 40!) > 2! Then Err.Raise vbObjectError, , "Width expected ~40, got " & Child.Widget.Width
    If Abs(Child.Widget.Height - 40!) > 2! Then Err.Raise vbObjectError, , "Height expected ~40, got " & Child.Widget.Height

    ' Stretch (default): fill cell after clearing Center
    Child.DependencyProperties.SetValue "HorizontalAlignment", "Stretch"
    Child.DependencyProperties.SetValue "VerticalAlignment", "Stretch"
    Child.Width = 0
    Child.Height = 0
    Root.Widget.Move 0, 0, 200, 200
    If Abs(Child.Widget.Left - 0!) > 2! Then Err.Raise vbObjectError, , "Stretch Left expected 0, got " & Child.Widget.Left
    If Abs(Child.Widget.Width - 200!) > 2! Then Err.Raise vbObjectError, , "Stretch Width expected ~200, got " & Child.Widget.Width
    If Abs(Child.Widget.Height - 200!) > 2! Then Err.Raise vbObjectError, , "Stretch Height expected ~200, got " & Child.Widget.Height

    KeepAlive Root
    LogResult "P2-GRID-ALIGN", 0, "OK Center + Stretch in Grid cell"
    Debug.Print "PASS  P2-GRID-ALIGN Grid Horizontal/VerticalAlignment"
    Phase2Bench_GridAlign = True
    Exit Function

Fail:
    LogResult "P2-GRID-ALIGN", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-GRID-ALIGN - " & Err.Description
    Phase2Bench_GridAlign = False
    On Error Resume Next
    KeepAlive Root
    Err.Clear
End Function

Public Function Phase2Bench_GridAttachedDpBag() As Boolean
    Dim G As Grid
    Dim P As Panel
    Dim Unset As Variant

    On Error GoTo Fail

    Set G = New Grid
    Set P = New Panel
    P.Width = 40
    P.Height = 40

    ' Lazy EnsureAttached on SetRow - DP bag + Get*
    G.SetRow P, 2
    G.SetColumn P, 1
    If Not P.DependencyProperties.Exists("Grid.Row") Then Err.Raise vbObjectError, , "Expected Grid.Row registered on target"
    If Not P.DependencyProperties.Exists("Grid.Column") Then Err.Raise vbObjectError, , "Expected Grid.Column registered on target"
    If G.GetRow(P) <> 2 Then Err.Raise vbObjectError, , "GetRow expected 2, got " & G.GetRow(P)
    If CLng(P.DependencyProperties.GetValue("Grid.Row")) <> 2 Then Err.Raise vbObjectError, , "DP Grid.Row expected 2"

    ' ClearValue restores metadata default (0) for layout Get*
    P.DependencyProperties.ClearValue "Grid.Row"
    If G.GetRow(P) <> 0 Then Err.Raise vbObjectError, , "GetRow after ClearValue expected 0, got " & G.GetRow(P)
    If CLng(P.DependencyProperties.GetValue("Grid.Row")) <> 0 Then Err.Raise vbObjectError, , "DP GetValue after Clear expected default 0"

    Unset = P.DependencyProperties.ReadLocalValue("Grid.Row")
    ' Local slot should be unset sentinel (not the old 2)
    If IsNumeric(Unset) Then
        If CLng(Unset) = 2 Then Err.Raise vbObjectError, , "ReadLocalValue still holds 2 after ClearValue"
    End If

    G.SetRow P, 1
    If G.GetRow(P) <> 1 Then Err.Raise vbObjectError, , "GetRow after re-Set expected 1"

    KeepAlive G
    KeepAlive P
    LogResult "P2-GRID-ATTACH-DP", 0, "OK EnsureAttached + ClearValue default"
    Debug.Print "PASS  P2-GRID-ATTACH-DP Grid.* DP bag EnsureAttached/ClearValue"
    Phase2Bench_GridAttachedDpBag = True
    Exit Function

Fail:
    LogResult "P2-GRID-ATTACH-DP", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-GRID-ATTACH-DP - " & Err.Description
    Phase2Bench_GridAttachedDpBag = False
    On Error Resume Next
    KeepAlive G
    KeepAlive P
    Err.Clear
End Function

Public Function Phase2Bench_DockPanelXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As DockPanel
    Dim LeftP As Panel
    Dim TopP As Panel
    Dim FillP As Panel

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\LayoutDockPanel.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "LayoutDockPanel returned Nothing"
    If Root.Children.Count <> 3 Then Err.Raise vbObjectError, , "Expected 3 children, got " & Root.Children.Count
    If Not Root.LastChildFill Then Err.Raise vbObjectError, , "LastChildFill expected True"

    Set LeftP = Root.Children(0)
    Set TopP = Root.Children(1)
    Set FillP = Root.Children(2)

    If Root.GetDock(LeftP) <> DockLeft Then Err.Raise vbObjectError, , "LeftDock expected DockLeft, got " & Root.GetDock(LeftP)
    If Root.GetDock(TopP) <> DockTop Then Err.Raise vbObjectError, , "TopDock expected DockTop, got " & Root.GetDock(TopP)

    Root.Widget.Move 0, 0, 200, 200

    ' Left: 40x200 at (0,0); Top: 160x30 at (40,0); Fill: 160x170 at (40,30)
    If Abs(LeftP.Widget.Left - 0!) > 2! Or Abs(LeftP.Widget.Width - 40!) > 2! Then
        Err.Raise vbObjectError, , "LeftDock expected L=0 W=40, got L=" & LeftP.Widget.Left & " W=" & LeftP.Widget.Width
    End If
    If Abs(LeftP.Widget.Height - 200!) > 2! Then
        Err.Raise vbObjectError, , "LeftDock Height expected ~200, got " & LeftP.Widget.Height
    End If
    If Abs(TopP.Widget.Left - 40!) > 2! Or Abs(TopP.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "TopDock expected (40,0), got (" & TopP.Widget.Left & "," & TopP.Widget.Top & ")"
    End If
    If Abs(TopP.Widget.Width - 160!) > 2! Or Abs(TopP.Widget.Height - 30!) > 2! Then
        Err.Raise vbObjectError, , "TopDock size expected 160x30, got " & TopP.Widget.Width & "x" & TopP.Widget.Height
    End If
    If Abs(FillP.Widget.Left - 40!) > 2! Or Abs(FillP.Widget.Top - 30!) > 2! Then
        Err.Raise vbObjectError, , "Fill expected (40,30), got (" & FillP.Widget.Left & "," & FillP.Widget.Top & ")"
    End If
    If Abs(FillP.Widget.Width - 160!) > 2! Or Abs(FillP.Widget.Height - 170!) > 2! Then
        Err.Raise vbObjectError, , "Fill size expected 160x170, got " & FillP.Widget.Width & "x" & FillP.Widget.Height
    End If

    KeepAlive Root
    LogResult "P2-DOCK-XAML", 0, "OK DockPanel.Dock XAML + LastChildFill"
    Debug.Print "PASS  P2-DOCK-XAML DockPanel.Dock XAML + LastChildFill"
    Phase2Bench_DockPanelXaml = True
    Exit Function

Fail:
    LogResult "P2-DOCK-XAML", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-DOCK-XAML - " & Err.Description
    Phase2Bench_DockPanelXaml = False
    On Error Resume Next
    KeepAlive Root
    Err.Clear
End Function

Public Function Phase2Bench_DockPanelLayout() As Boolean
    Dim Dp As DockPanel
    Dim LeftP As Panel
    Dim TopP As Panel
    Dim FillP As Panel

    On Error GoTo Fail

    Set Dp = New DockPanel
    Dp.Width = 200
    Dp.Height = 200
    Dp.LastChildFill = True
    Dp.Widget.Move 0, 0, 200, 200

    Set LeftP = New Panel
    LeftP.Width = 50
    LeftP.Height = 0
    Set TopP = New Panel
    TopP.Width = 0
    TopP.Height = 40
    Set FillP = New Panel
    FillP.Width = 0
    FillP.Height = 0

    Dp.SetDock LeftP, DockLeft
    Dp.SetDock TopP, DockTop

    Dp.Children.Add LeftP
    Dp.Children.Add TopP
    Dp.Children.Add FillP

    Dp.Widget.Move 0, 0, 200, 200

    If Abs(LeftP.Widget.Left - 0!) > 2! Or Abs(LeftP.Widget.Width - 50!) > 2! Then
        Err.Raise vbObjectError, , "Left expected L=0 W=50, got L=" & LeftP.Widget.Left & " W=" & LeftP.Widget.Width
    End If
    If Abs(LeftP.Widget.Height - 200!) > 2! Then
        Err.Raise vbObjectError, , "Left Height expected 200, got " & LeftP.Widget.Height
    End If
    If Abs(TopP.Widget.Left - 50!) > 2! Or Abs(TopP.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "Top expected (50,0), got (" & TopP.Widget.Left & "," & TopP.Widget.Top & ")"
    End If
    If Abs(TopP.Widget.Height - 40!) > 2! Then
        Err.Raise vbObjectError, , "Top Height expected 40, got " & TopP.Widget.Height
    End If
    If Abs(FillP.Widget.Left - 50!) > 2! Or Abs(FillP.Widget.Top - 40!) > 2! Then
        Err.Raise vbObjectError, , "Fill expected (50,40), got (" & FillP.Widget.Left & "," & FillP.Widget.Top & ")"
    End If
    If Abs(FillP.Widget.Width - 150!) > 2! Or Abs(FillP.Widget.Height - 160!) > 2! Then
        Err.Raise vbObjectError, , "Fill size expected 150x160, got " & FillP.Widget.Width & "x" & FillP.Widget.Height
    End If

    KeepAlive Dp
    LogResult "P2-DOCK-LAY", 0, "OK Left+Top+Fill arrange"
    Debug.Print "PASS  P2-DOCK-LAY DockPanel Left+Top+Fill arrange"
    Phase2Bench_DockPanelLayout = True
    Exit Function

Fail:
    LogResult "P2-DOCK-LAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-DOCK-LAY - " & Err.Description
    Phase2Bench_DockPanelLayout = False
    On Error Resume Next
    KeepAlive Dp
    Err.Clear
End Function

Public Function Phase2Bench_CanvasXaml() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Canvas
    Dim AtOrigin As Panel
    Dim Offset As Panel

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\LayoutCanvas.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "LayoutCanvas returned Nothing"
    If Root.Children.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 children, got " & Root.Children.Count

    Set AtOrigin = Root.Children(0)
    Set Offset = Root.Children(1)

    If Abs(Root.GetLeft(AtOrigin) - 10#) > 0.01 Then Err.Raise vbObjectError, , "GetLeft AtOrigin expected 10"
    If Abs(Root.GetTop(AtOrigin) - 20#) > 0.01 Then Err.Raise vbObjectError, , "GetTop AtOrigin expected 20"
    If Abs(Root.GetLeft(Offset) - 80#) > 0.01 Then Err.Raise vbObjectError, , "GetLeft Offset expected 80"

    Root.Widget.Move 0, 0, 200, 200

    If Abs(AtOrigin.Widget.Left - 10!) > 2! Or Abs(AtOrigin.Widget.Top - 20!) > 2! Then
        Err.Raise vbObjectError, , "AtOrigin expected (10,20), got (" & AtOrigin.Widget.Left & "," & AtOrigin.Widget.Top & ")"
    End If
    If Abs(Offset.Widget.Left - 80!) > 2! Or Abs(Offset.Widget.Top - 100!) > 2! Then
        Err.Raise vbObjectError, , "Offset expected (80,100), got (" & Offset.Widget.Left & "," & Offset.Widget.Top & ")"
    End If

    KeepAlive Root
    LogResult "P2-CANVAS-XAML", 0, "OK Canvas.Left/Top XAML positions"
    Debug.Print "PASS  P2-CANVAS-XAML Canvas.Left/Top XAML positions"
    Phase2Bench_CanvasXaml = True
    Exit Function

Fail:
    LogResult "P2-CANVAS-XAML", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-CANVAS-XAML - " & Err.Description
    Phase2Bench_CanvasXaml = False
    On Error Resume Next
    KeepAlive Root
    Err.Clear
End Function

Public Function Phase2Bench_CanvasLayout() As Boolean
    Dim Cv As Canvas
    Dim A As Panel
    Dim B As Panel

    On Error GoTo Fail

    Set Cv = New Canvas
    Cv.Width = 200
    Cv.Height = 200
    Cv.Widget.Move 0, 0, 200, 200

    Set A = New Panel
    A.Width = 30
    A.Height = 20
    Set B = New Panel
    B.Width = 40
    B.Height = 25

    Cv.SetLeft A, 15
    Cv.SetTop A, 25
    Cv.SetLeft B, 100
    Cv.SetTop B, 50

    Cv.Children.Add A
    Cv.Children.Add B
    Cv.Widget.Move 0, 0, 200, 200

    If Abs(A.Widget.Left - 15!) > 2! Or Abs(A.Widget.Top - 25!) > 2! Then
        Err.Raise vbObjectError, , "A expected (15,25), got (" & A.Widget.Left & "," & A.Widget.Top & ")"
    End If
    If Abs(B.Widget.Left - 100!) > 2! Or Abs(B.Widget.Top - 50!) > 2! Then
        Err.Raise vbObjectError, , "B expected (100,50), got (" & B.Widget.Left & "," & B.Widget.Top & ")"
    End If

    KeepAlive Cv
    LogResult "P2-CANVAS-LAY", 0, "OK SetLeft/SetTop arrange"
    Debug.Print "PASS  P2-CANVAS-LAY Canvas SetLeft/SetTop arrange"
    Phase2Bench_CanvasLayout = True
    Exit Function

Fail:
    LogResult "P2-CANVAS-LAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2-CANVAS-LAY - " & Err.Description
    Phase2Bench_CanvasLayout = False
    On Error Resume Next
    KeepAlive Cv
    Err.Clear
End Function

Public Function Phase3Bench_MergedDictionaryLookup() As Boolean
    Dim Root As ResourceDictionary
    Dim Child As ResourceDictionary
    Dim Value As Variant

    On Error GoTo Fail

    Set Root = New ResourceDictionary
    Set Child = New ResourceDictionary
    Child.Add "TestKey", "hello"

    Root.Merge Child

    If Not Root.TryGetResource("TestKey", Value) Then
        Err.Raise vbObjectError, , "Merged key not found"
    End If
    If Value <> "hello" Then
        Err.Raise vbObjectError, , "Expected 'hello', got " & CStr(Value)
    End If

    LogResult "P3-MERGE", 0, "OK TryGetResource=hello"
    Debug.Print "PASS  P3-MERGE Merged dictionary lookup"
    Phase3Bench_MergedDictionaryLookup = True
    Exit Function

Fail:
    LogResult "P3-MERGE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P3-MERGE ? " & Err.Description
    Phase3Bench_MergedDictionaryLookup = False
End Function

Public Function Phase3Bench_ResourceSourceLoad() As Boolean
    Dim Resolver As XamlResourceResolver
    Dim Dict As ResourceDictionary
    Dim Value As Variant

    On Error GoTo Fail

    Set Resolver = New XamlResourceResolver
    Resolver.BasePath = App.Path & "\Resources"
    Set Dict = Resolver.LoadFromSource("P3ChildDict.xml")

    If Dict Is Nothing Then Err.Raise vbObjectError, , "LoadFromSource returned Nothing"
    If Not Dict.TryGetResource("Greeting", Value) Then
        Err.Raise vbObjectError, , "Greeting key not found in sourced dictionary"
    End If
    If Value <> "Phase3" Then
        Err.Raise vbObjectError, , "Expected 'Phase3', got " & CStr(Value)
    End If

    LogResult "P3-SOURCE", 0, "OK Greeting=Phase3"
    Debug.Print "PASS  P3-SOURCE ResourceDictionary Source load"
    Phase3Bench_ResourceSourceLoad = True
    Exit Function

Fail:
    LogResult "P3-SOURCE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P3-SOURCE ? " & Err.Description
    Phase3Bench_ResourceSourceLoad = False
End Function

Public Function Phase3Bench_DynamicResourceExtension() As Boolean
    Dim P As Panel
    Dim El As IUIElement
    Dim Value As Variant

    On Error GoTo Fail

    Set P = New Panel
    Set El = P
    El.Base.Resources.Add "BgColor", 12345

    API.CopyVariable El.Base.TryFindResource("BgColor"), Value

    If IsEmpty(Value) Then Err.Raise vbObjectError, , "TryFindResource returned Empty"
    If CLng(Value) <> 12345 Then
        Err.Raise vbObjectError, , "Expected 12345, got " & CStr(Value)
    End If

    LogResult "P3-DYNAMIC", 0, "OK BgColor=12345"
    Debug.Print "PASS  P3-DYNAMIC element TryFindResource (DynamicResource lookup path)"
    Phase3Bench_DynamicResourceExtension = True
    Exit Function

Fail:
    LogResult "P3-DYNAMIC", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P3-DYNAMIC ? " & Err.Description
    Phase3Bench_DynamicResourceExtension = False
End Function

Public Function Phase3Bench_StrictUnknownProperty() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Object
    Dim SavedStrict As Boolean

    On Error GoTo Fail

    SavedStrict = VCF.StrictXamlLoad
    VCF.StrictXamlLoad = True

    Set Reader = New XAMLReader
    Set Root = Reader.Load("<Panel NotARealProperty=""1""/>")

    VCF.StrictXamlLoad = SavedStrict
    Err.Raise vbObjectError, , "Expected XamlLoadException for unknown property"
    Exit Function

Fail:
    VCF.StrictXamlLoad = SavedStrict
    If Err.Source = "VCF.XamlLoadException" Then
        LogResult "P3-STRICT-PROP", 0, "OK raised XamlLoadException"
        Debug.Print "PASS  P3-STRICT Unknown property raises"
        Phase3Bench_StrictUnknownProperty = True
    Else
        LogResult "P3-STRICT-PROP", 0, "FAIL: " & Err.Number & " " & Err.Description
        Debug.Print "FAIL  P3-STRICT ? " & Err.Description
        Phase3Bench_StrictUnknownProperty = False
    End If
End Function

Public Function Phase4Bench_BindingOneWay() As Boolean
    Dim Vm As Phase0ViewModel
    Dim Tb As TextBlock
    Dim Expr As BindingExpression

    On Error GoTo Fail

    Set Vm = New Phase0ViewModel
    Vm.Title = "Hello"
    Set Tb = New TextBlock
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Vm, "Title", OneWay

    If Tb.Text <> "Hello" Then Err.Raise vbObjectError, , "Expected Hello, got " & Tb.Text

    Vm.Title = "World"
    If Tb.Text <> "World" Then Err.Raise vbObjectError, , "Expected World after INPC, got " & Tb.Text

    LogResult "P4-BIND", 0, "OK OneWay Title binding"
    Debug.Print "PASS  P4-BIND OneWay binding + INPC"
    Expr.Detach
    Set Expr = Nothing
    Set Tb = Nothing
    Set Vm = Nothing
    Phase4Bench_BindingOneWay = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-BIND", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-BIND ? " & Err.Description
    Phase4Bench_BindingOneWay = False
End Function

Public Function Phase4Bench_BindingAttached() As Boolean
    Dim Vm As Phase0ViewModel
    Dim G As Grid
    Dim Child As Panel
    Dim Expr As BindingExpression

    On Error GoTo Fail

    Set Vm = New Phase0ViewModel
    Vm.RowIndex = 1

    Set G = New Grid
    Set Child = New Panel
    ' Grid.Row not EnsureAttached yet - Attach must lazy-register then bind.
    Set Expr = New BindingExpression
    Expr.Attach Child, "Grid.Row", Vm, "RowIndex", OneWay
    If Expr.Binding Is Nothing Then Err.Raise vbObjectError, , "Attach to Grid.Row failed (no Binding)"

    If Not Child.DependencyProperties.Exists("Grid.Row") Then
        Err.Raise vbObjectError, , "Grid.Row not registered on target after Attach"
    End If
    If CLng(Child.DependencyProperties.GetValue("Grid.Row")) <> 1 Then
        Err.Raise vbObjectError, , "Expected Grid.Row=1, got " & Child.DependencyProperties.GetValue("Grid.Row")
    End If
    If G.GetRow(Child) <> 1 Then Err.Raise vbObjectError, , "GetRow expected 1, got " & G.GetRow(Child)

    Vm.RowIndex = 2
    If CLng(Child.DependencyProperties.GetValue("Grid.Row")) <> 2 Then
        Err.Raise vbObjectError, , "Expected Grid.Row=2 after INPC, got " & Child.DependencyProperties.GetValue("Grid.Row")
    End If
    If G.GetRow(Child) <> 2 Then Err.Raise vbObjectError, , "GetRow expected 2 after INPC, got " & G.GetRow(Child)

    KeepAlive G
    KeepAlive Child
    LogResult "P4-BIND-ATTACH", 0, "OK OneWay Grid.Row binding + INPC"
    Debug.Print "PASS  P4-BIND-ATTACH OneWay binding to Grid.Row + INPC"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_BindingAttached = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-BIND-ATTACH", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-BIND-ATTACH - " & Err.Description
    Phase4Bench_BindingAttached = False
    KeepAlive G
    KeepAlive Child
End Function

Public Function Phase4Bench_BindingAttachedPath() As Boolean
    Dim G As Grid
    Dim Src As Panel
    Dim Dst As Panel
    Dim Expr As BindingExpression

    On Error GoTo Fail

    Set G = New Grid
    Set Src = New Panel
    Set Dst = New Panel

    G.SetRow Src, 2

    ' Source Path=(Grid.Row) reads attached DP from Src into Dst.Grid.Column
    Set Expr = New BindingExpression
    Expr.Attach Dst, "Grid.Column", Src, "(Grid.Row)", OneWay
    If Expr.Binding Is Nothing Then Err.Raise vbObjectError, , "Attach Path=(Grid.Row) failed"

    If G.GetColumn(Dst) <> 2 Then Err.Raise vbObjectError, , "Expected Column=2 from Path=(Grid.Row), got " & G.GetColumn(Dst)

    ' Live update via attached DP PropertyChanged on source
    G.SetRow Src, 3
    If G.GetColumn(Dst) <> 3 Then Err.Raise vbObjectError, , "Expected Column=3 after SetRow, got " & G.GetColumn(Dst)

    KeepAlive G
    KeepAlive Src
    KeepAlive Dst
    LogResult "P4-BIND-APATH", 0, "OK Path=(Grid.Row) + live update"
    Debug.Print "PASS  P4-BIND-APATH Path=(Grid.Row) source + live update"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_BindingAttachedPath = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-BIND-APATH", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-BIND-APATH - " & Err.Description
    Phase4Bench_BindingAttachedPath = False
    KeepAlive G
    KeepAlive Src
    KeepAlive Dst
End Function

' Binding Grid.Row + INPC must rearrange the parent (widget Top), not only bag GetValue.
Public Function Phase4Bench_BindingAttachedLayout() As Boolean
    Dim Vm As Phase0ViewModel
    Dim G As Grid
    Dim Rd As RowDefinition
    Dim Child As Panel
    Dim Expr As BindingExpression

    On Error GoTo Fail

    Set Vm = New Phase0ViewModel
    Vm.RowIndex = 0

    Set G = New Grid
    G.Width = 100
    G.Height = 200
    Set Rd = New RowDefinition
    Rd.Height = "*"
    G.RowDefinitions.Add Rd
    Set Rd = New RowDefinition
    Rd.Height = "*"
    G.RowDefinitions.Add Rd

    Set Child = New Panel
    Child.Width = 40
    Child.Height = 40

    Set Expr = New BindingExpression
    Expr.Attach Child, "Grid.Row", Vm, "RowIndex", OneWay
    If Expr.Binding Is Nothing Then Err.Raise vbObjectError, , "Attach to Grid.Row failed"

    G.Children.Add Child
    G.Widget.Move 0, 0, 100, 200

    If Abs(Child.Widget.Top - 0!) > 2! Then
        Err.Raise vbObjectError, , "Row0 Top expected 0, got " & Child.Widget.Top
    End If

    Vm.RowIndex = 1
    If CLng(Child.DependencyProperties.GetValue("Grid.Row")) <> 1 Then
        Err.Raise vbObjectError, , "Expected Grid.Row=1 after INPC, got " & Child.DependencyProperties.GetValue("Grid.Row")
    End If
    If Abs(Child.Widget.Top - 100!) > 2! Then
        Err.Raise vbObjectError, , "Row1 Top expected 100 after INPC, got " & Child.Widget.Top
    End If

    KeepAlive G
    KeepAlive Child
    LogResult "P4-BIND-LAY", 0, "OK Grid.Row binding rearranges Top"
    Debug.Print "PASS  P4-BIND-LAY Grid.Row INPC rearranges widget Top"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_BindingAttachedLayout = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-BIND-LAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-BIND-LAY - " & Err.Description
    Phase4Bench_BindingAttachedLayout = False
    KeepAlive G
    KeepAlive Child
End Function

Public Function Phase4Bench_DataContextRebind() As Boolean
    Dim Vm1 As Phase0ViewModel
    Dim Vm2 As Phase0ViewModel
    Dim Tb As TextBlock
    Dim Expr As BindingExpression
    Dim DataContextProp As DependencyProperty

    On Error GoTo Fail

    Set Vm1 = New Phase0ViewModel
    Vm1.Title = "One"
    Set Vm2 = New Phase0ViewModel
    Vm2.Title = "Two"

    Set Tb = New TextBlock
    Set DataContextProp = Tb.DependencyProperties.GetProperty("DataContext")

    Set Tb.DataContext = Vm1
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", DataContextProp, "Title", OneWay

    If Tb.Text <> "One" Then Err.Raise vbObjectError, , "Expected One, got " & Tb.Text

    Set Tb.DataContext = Vm2
    If Tb.Text <> "Two" Then Err.Raise vbObjectError, , "Expected Two after DataContext swap, got " & Tb.Text

    LogResult "P4-DCTX", 0, "OK DataContext rebind"
    Debug.Print "PASS  P4-DCTX DataContext rebind"
    Expr.Detach
    Set Expr = Nothing
    Set Tb = Nothing
    Set Vm1 = Nothing
    Set Vm2 = Nothing
    Phase4Bench_DataContextRebind = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-DCTX", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-DCTX ? " & Err.Description
    Phase4Bench_DataContextRebind = False
End Function

Public Function Phase4Bench_BindingDetach() As Boolean
    Dim Vm As Phase0ViewModel
    Dim Tb As TextBlock
    Dim Expr As BindingExpression

    On Error GoTo Fail

    Set Vm = New Phase0ViewModel
    Vm.Title = "Before"
    Set Tb = New TextBlock
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Vm, "Title", OneWay

    If Tb.Text <> "Before" Then Err.Raise vbObjectError, , "Expected Before, got " & Tb.Text

    Expr.Detach
    Vm.Title = "After"
    ' Detach clears local binding value (WPF ClearBinding); Text falls back to unset/default.
    If Len(Tb.Text) <> 0 Then Err.Raise vbObjectError, , "Expected empty Text after Detach+ClearValue, got [" & Tb.Text & "]"

    LogResult "P4-DETACH", 0, "OK Detach clears local + stops updates"
    Debug.Print "PASS  P4-DETACH Binding Detach"
    Set Expr = Nothing
    Set Tb = Nothing
    Set Vm = Nothing
    Phase4Bench_BindingDetach = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-DETACH", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-DETACH - " & Err.Description
    Phase4Bench_BindingDetach = False
End Function

' Locked two-slot precedence + ClearValue + metadata default (conflict #4 / 3.2.0).
Public Function Phase4Bench_DpPrecedence() As Boolean
    Dim Tb As TextBlock
    Dim CC As ContentControl
    Dim LocalRaw As Variant
    Dim Meta As DependencyPropertyMetadata

    On Error GoTo Fail

    Set Tb = New TextBlock
    Tb.DependencyProperties.SetCurrentValue "Text", "FromStyle"
    Tb.DependencyProperties.SetValue "Text", "FromLocal"
    If Tb.Text <> "FromLocal" Then Err.Raise vbObjectError, , "Local SetValue must beat SetCurrentValue, got " & Tb.Text

    Tb.DependencyProperties.SetCurrentValue "Text", "StyleAgain"
    If Tb.Text <> "FromLocal" Then Err.Raise vbObjectError, , "SetCurrentValue must not pierce local, got " & Tb.Text

    Tb.DependencyProperties.ClearValue "Text"
    If Tb.Text <> "StyleAgain" Then Err.Raise vbObjectError, , "ClearValue must fall back to current, got " & Tb.Text

    Call API.CopyVariable(Tb.DependencyProperties.ReadLocalValue("Text"), LocalRaw)
    If Not Object.Equals(LocalRaw, Tb.DependencyProperties.GetProperty("Text").UnsetValue) Then
        Err.Raise vbObjectError, , "ReadLocalValue expected Unset after ClearValue"
    End If

    Set CC = New ContentControl
    Set Meta = CC.DependencyProperties.GetProperty("Content").Metadata
    If Not Meta.HasDefaultValue Then Err.Raise vbObjectError, , "Content metadata HasDefaultValue expected True"
    If CStr(Meta.DefaultValue) <> "" Then Err.Raise vbObjectError, , "Content DefaultValue expected empty string"
    CC.DependencyProperties.SetValue "Content", "LocalContent"
    If CStr(CC.Content) <> "LocalContent" Then Err.Raise vbObjectError, , "Local Content expected"
    CC.DependencyProperties.ClearValue "Content"
    ' No style current -> metadata default "" (Window is not creatable; BorderStyle default gated in product code).
    If CStr(CC.Content) <> "" Then Err.Raise vbObjectError, , "ClearValue Content expected metadata default empty, got " & CStr(CC.Content)

    KeepAlive Tb
    KeepAlive CC
    Set Tb = Nothing
    Set CC = Nothing

    LogResult "P4-PREC", 0, "OK local>current ClearValue + metadata default"
    Debug.Print "PASS  P4-PREC DP value precedence"
    Phase4Bench_DpPrecedence = True
    Exit Function

Fail:
    LogResult "P4-PREC", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-PREC - " & Err.Description
    Phase4Bench_DpPrecedence = False
    On Error Resume Next
    KeepAlive Tb
    KeepAlive CC
    Err.Clear
End Function

Public Function Phase4Bench_RelativeSourceSelf() As Boolean
    Dim Tb As TextBlock
    Dim Expr As BindingExpression
    Dim RS As RelativeSource
    Dim Src As Object

    On Error GoTo Fail

    Set Tb = New TextBlock
    Tb.Name = "SelfOk"
    Set RS = New RelativeSource
    RS.Mode = RelativeSourceSelf
    Set Src = RS.Resolve(Tb)
    If Not Src Is Tb Then Err.Raise vbObjectError, , "RelativeSource Self must resolve to Target"

    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Src, "Name", OneWay
    If Tb.Text <> "SelfOk" Then Err.Raise vbObjectError, , "Expected SelfOk, got " & Tb.Text

    LogResult "P4-RSELF", 0, "OK RelativeSource Self"
    Debug.Print "PASS  P4-RSELF RelativeSource Self"
    Expr.Detach
    Set Expr = Nothing
    Set Tb = Nothing
    Phase4Bench_RelativeSourceSelf = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-RSELF", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-RSELF - " & Err.Description
    Phase4Bench_RelativeSourceSelf = False
End Function

Public Function Phase4Bench_ElementName() As Boolean
    Dim Sp As StackPanel
    Dim TbSrc As TextBlock
    Dim TbDst As TextBlock
    Dim Expr As BindingExpression
    Dim Src As Object
    Dim Anc As Object
    Dim RS As RelativeSource

    On Error GoTo Fail

    Set Sp = New StackPanel
    Set TbSrc = New TextBlock
    TbSrc.Name = "Src"
    TbSrc.Text = "FromName"
    Set TbDst = New TextBlock
    TbDst.Name = "Dst"
    Sp.Children.Add TbSrc
    Sp.Children.Add TbDst

    Set Src = VCF.NamingManager.ResolveElementName(TbDst, "Src")
    If Src Is Nothing Then Err.Raise vbObjectError, , "ElementName Src not found"
    If Not Src Is TbSrc Then Err.Raise vbObjectError, , "ElementName must resolve to named sibling"

    Set Expr = New BindingExpression
    Expr.Attach TbDst, "Text", Src, "Text", OneWay
    If TbDst.Text <> "FromName" Then Err.Raise vbObjectError, , "Expected FromName, got " & TbDst.Text

    Set RS = New RelativeSource
    RS.Mode = RelativeSourceFindAncestor
    RS.AncestorType = "StackPanel"
    RS.AncestorLevel = 1
    Set Anc = RS.Resolve(TbDst)
    If Not Anc Is Sp Then Err.Raise vbObjectError, , "FindAncestor StackPanel expected"

    LogResult "P4-ENAME", 0, "OK ElementName + FindAncestor"
    Debug.Print "PASS  P4-ENAME ElementName + FindAncestor"
    Expr.Detach
    Set Expr = Nothing
    KeepAlive Sp
    Phase4Bench_ElementName = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-ENAME", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-ENAME - " & Err.Description
    KeepAlive Sp
    Phase4Bench_ElementName = False
End Function

Public Function Phase4Bench_RelativeSourceTemplatedParent() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim CP As ContentPresenter
    Dim Live As Border
    Dim Nested As ContentPresenter
    Dim Tb As TextBlock
    Dim Expr As BindingExpression
    Dim RS As RelativeSource
    Dim Src As Object
    Dim Rad As VCF.CornerRadius

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"
    Set B = New Border
    Rad.TopLeft = 4
    Rad.TopRight = 4
    Rad.BottomLeft = 4
    Rad.BottomRight = 4
    B.CornerRadius = Rad
    Set CP = New ContentPresenter
    Set B.Child = CP
    Tmpl.Children.Add B

    Set St = NewStyle("Button")
    Set St.Template = Tmpl

    Set Btn = New Button
    Btn.Content = "TP-OK"
    Set Btn.Style = St

    If Btn.ContentPresenter Is Nothing Then Err.Raise vbObjectError, , "Expected live ContentPresenter"
    Set Nested = Btn.ContentPresenter
    If Nested.TemplatedParent Is Nothing Then Err.Raise vbObjectError, , "TemplatedParent not stamped"
    If Not Nested.TemplatedParent Is Btn Then Err.Raise vbObjectError, , "TemplatedParent must be Button"

    Set Live = Btn.Children(0)
    If Live.TemplatedParent Is Nothing Then Err.Raise vbObjectError, , "Chrome TemplatedParent missing"
    If Not Live.TemplatedParent Is Btn Then Err.Raise vbObjectError, , "Chrome TemplatedParent must be Button"

    Set RS = New RelativeSource
    RS.Mode = RelativeSourceTemplatedParent
    Set Src = RS.Resolve(Nested)
    If Not Src Is Btn Then Err.Raise vbObjectError, , "RelativeSource TemplatedParent resolve failed"

    Set Tb = New TextBlock
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Src, "Content", OneWay
    If Tb.Text <> "TP-OK" Then Err.Raise vbObjectError, , "Expected TP-OK, got " & Tb.Text

    KeepAlive Btn
    LogResult "P4-RTP", 0, "OK RelativeSource TemplatedParent"
    Debug.Print "PASS  P4-RTP RelativeSource TemplatedParent"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_RelativeSourceTemplatedParent = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-RTP", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-RTP - " & Err.Description
    KeepAlive Btn
    Phase4Bench_RelativeSourceTemplatedParent = False
End Function

' Window-level Command pattern: named host exposes ICommand; action Button binds via ElementName.
Public Function Phase4Bench_ElementNameCommand() As Boolean
    Dim Sp As StackPanel
    Dim Host As Button
    Dim Act As Button
    Dim SharedCmd As Phase0DialogCommand
    Dim Expr As BindingExpression
    Dim Src As Object
    Dim Cmd As ICommand

    On Error GoTo Fail

    Set SharedCmd = New Phase0DialogCommand
    SharedCmd.Reset

    Set Sp = New StackPanel

    Set Host = New Button
    Host.Name = "CmdHost"
    Host.Content = "Host"
    Set Host.Command = SharedCmd
    Host.CommandParameter = "from-host"

    Set Act = New Button
    Act.Name = "Act"
    Act.Content = "Go"
    Act.CommandParameter = "from-act"

    Sp.Children.Add Host
    Sp.Children.Add Act

    Set Src = VCF.NamingManager.ResolveElementName(Act, "CmdHost")
    If Src Is Nothing Then Err.Raise vbObjectError, , "ElementName CmdHost not found"
    If Not Src Is Host Then Err.Raise vbObjectError, , "ElementName must resolve to CmdHost"

    Set Expr = New BindingExpression
    Expr.Attach Act, "Command", Src, "Command", OneWay

    If Act.Command Is Nothing Then Err.Raise vbObjectError, , "Act.Command not bound"
    If Not Act.Command Is SharedCmd Then Err.Raise vbObjectError, , "Act.Command must be SharedCmd"

    Set Cmd = Act.Command
    Cmd.Execute Act.CommandParameter
    If SharedCmd.ExecuteCount <> 1 Then Err.Raise vbObjectError, , "ExecuteCount expected 1"
    If CStr(SharedCmd.LastParameter) <> "from-act" Then Err.Raise vbObjectError, , "LastParameter expected from-act"

    KeepAlive Sp
    LogResult "P4-ECMD", 0, "OK ElementName Command binding"
    Debug.Print "PASS  P4-ECMD ElementName Command binding"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_ElementNameCommand = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-ECMD", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-ECMD - " & Err.Description
    KeepAlive Sp
    Phase4Bench_ElementNameCommand = False
End Function

Public Function Phase4Bench_UpdateSourceTrigger() As Boolean
    Dim Vm As Phase0ViewModel
    Dim Tb As VCF.TextBox
    Dim Expr As BindingExpression
    Dim FailMsg As String

    On Error GoTo Fail

    ' Qualify VCF.TextBox ? bare TextBox is VB.TextBox in this EXE.
    ' LostFocus: target change must not push source until UpdateSource/Flush.
    Set Vm = New Phase0ViewModel
    Vm.Title = "Initial"
    Set Tb = New VCF.TextBox
    Tb.Text = "Initial"
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Vm, "Title", TwoWay, ustLostFocus

    Tb.Text = "Deferred"
    If Vm.Title <> "Initial" Then Err.Raise vbObjectError, , "LostFocus: VM updated too early, got " & Vm.Title
    If Tb.Text <> "Deferred" Then Err.Raise vbObjectError, , "LostFocus: TextBox Text expected Deferred"

    Expr.UpdateSource
    If Vm.Title <> "Deferred" Then Err.Raise vbObjectError, , "LostFocus: VM expected Deferred after UpdateSource, got " & Vm.Title

    Expr.Detach
    Set Expr = Nothing
    Set Tb = Nothing
    Set Vm = Nothing

    ' PropertyChanged (default): immediate Target -> Source (regression).
    Set Vm = New Phase0ViewModel
    Vm.Title = "A"
    Set Tb = New VCF.TextBox
    Tb.Text = "A"
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Vm, "Title", TwoWay

    Tb.Text = "B"
    If Vm.Title <> "B" Then Err.Raise vbObjectError, , "PropertyChanged: VM expected B, got " & Vm.Title

    KeepAlive Tb
    LogResult "P4-UST", 0, "OK UpdateSourceTrigger LostFocus + PropertyChanged"
    Debug.Print "PASS  P4-UST UpdateSourceTrigger LostFocus + PropertyChanged"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_UpdateSourceTrigger = True
    Exit Function

Fail:
    FailMsg = Err.Description
    If Len(FailMsg) = 0 Then FailMsg = "Error " & CStr(Err.Number)
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-UST", 0, "FAIL: " & FailMsg
    Debug.Print "FAIL  P4-UST - " & FailMsg
    KeepAlive Tb
    Phase4Bench_UpdateSourceTrigger = False
End Function

Public Function Phase4Bench_UpdateSourceDelay() As Boolean
    Dim Vm As Phase0ViewModel
    Dim Tb As VCF.TextBox
    Dim Expr As BindingExpression
    Dim B As Binding
    Dim FailMsg As String

    On Error GoTo Fail

    Set Vm = New Phase0ViewModel
    Vm.Title = "Initial"
    Set Tb = New VCF.TextBox
    Tb.Text = "Initial"
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Vm, "Title", TwoWay
    Set B = Expr.Binding
    ' Long delay so timer cannot fire during the gate; assert debounce + Flush.
    B.UpdateSourceDelay = 60000

    Tb.Text = "Deferred"
    If Vm.Title <> "Initial" Then Err.Raise vbObjectError, , "Delay: VM updated too early, got " & Vm.Title
    If Tb.Text <> "Deferred" Then Err.Raise vbObjectError, , "Delay: TextBox.Text expected Deferred"

    Expr.UpdateSource
    If Vm.Title <> "Deferred" Then Err.Raise vbObjectError, , "Delay: VM expected Deferred after UpdateSource, got " & Vm.Title

    Tb.Text = "Again"
    If Vm.Title <> "Deferred" Then Err.Raise vbObjectError, , "Delay: VM should stay Deferred until flush, got " & Vm.Title
    Expr.UpdateSource
    If Vm.Title <> "Again" Then Err.Raise vbObjectError, , "Delay: VM expected Again after UpdateSource, got " & Vm.Title

    KeepAlive Tb
    LogResult "P4-UDELAY", 0, "OK UpdateSourceDelay debounce + Flush"
    Debug.Print "PASS  P4-UDELAY UpdateSourceDelay debounce + Flush"
    Expr.Detach
    Set Expr = Nothing
    Phase4Bench_UpdateSourceDelay = True
    Exit Function

Fail:
    FailMsg = Err.Description
    If Len(FailMsg) = 0 Then FailMsg = "Error " & CStr(Err.Number)
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P4-UDELAY", 0, "FAIL: " & FailMsg
    Debug.Print "FAIL  P4-UDELAY - " & FailMsg
    KeepAlive Tb
    Phase4Bench_UpdateSourceDelay = False
End Function

Public Function Phase4Bench_TextCaretPreserve() As Boolean
    Dim Tb As VCF.TextBox
    Dim FailMsg As String

    On Error GoTo Fail

    Set Tb = New VCF.TextBox
    Tb.Text = "AB"
    Tb.SelStart = 2
    Tb.SelLength = 0

    ' Prefix extension (binding echo): caret at end must move to new end, not 0.
    Tb.Text = "ABC"
    If Tb.Text <> "ABC" Then Err.Raise vbObjectError, , "Text expected ABC"
    If Tb.SelStart <> 3 Then Err.Raise vbObjectError, , "Caret expected 3 after append, got " & CStr(Tb.SelStart)

    Tb.Text = "AB"
    Tb.SelStart = 1
    Tb.Text = "ABX"
    ' Not a pure extension of "AB" -> "ABX" wait - "ABX" Left 2 = "AB", Len >=, OldStart=1 < Len(Old)=2
    ' So NewStart = OldStart = 1 ? preserve mid caret on extension
    If Tb.SelStart <> 1 Then Err.Raise vbObjectError, , "Mid caret expected 1 on extension, got " & CStr(Tb.SelStart)

    Tb.Text = "ZZ"
    ' Unrelated replace ? caret resets to 0
    If Tb.SelStart <> 0 Then Err.Raise vbObjectError, , "Unrelated replace expected caret 0, got " & CStr(Tb.SelStart)

    KeepAlive Tb
    LogResult "P4-CARET", 0, "OK TextBox caret preserve on prefix extension"
    Debug.Print "PASS  P4-CARET TextBox caret preserve on prefix extension"
    Phase4Bench_TextCaretPreserve = True
    Exit Function

Fail:
    FailMsg = Err.Description
    If Len(FailMsg) = 0 Then FailMsg = "Error " & CStr(Err.Number)
    On Error Resume Next
    LogResult "P4-CARET", 0, "FAIL: " & FailMsg
    Debug.Print "FAIL  P4-CARET - " & FailMsg
    KeepAlive Tb
    Phase4Bench_TextCaretPreserve = False
End Function

Public Function Phase4Bench_CanExecuteChanged() As Boolean
    Dim Btn As Button
    Dim Cmd As Phase0DialogCommand
    Dim FailMsg As String

    On Error GoTo Fail

    Set Cmd = New Phase0DialogCommand
    Cmd.Reset
    Cmd.CanExecute = True

    Set Btn = New Button
    Set Btn.Command = Cmd
    If Not Btn.Widget.Enabled Then Err.Raise vbObjectError, , "Expected Enabled=True when CanExecute=True"

    Cmd.CanExecute = False
    If Btn.Widget.Enabled Then Err.Raise vbObjectError, , "Expected Enabled=False after CanExecuteChanged"

    Cmd.CanExecute = True
    If Not Btn.Widget.Enabled Then Err.Raise vbObjectError, , "Expected Enabled=True after re-enable"

    KeepAlive Btn
    LogResult "P4-CCMD", 0, "OK CanExecuteChanged syncs Button.Enabled"
    Debug.Print "PASS  P4-CCMD CanExecuteChanged syncs Button.Enabled"
    Phase4Bench_CanExecuteChanged = True
    Exit Function

Fail:
    FailMsg = Err.Description
    If Len(FailMsg) = 0 Then FailMsg = "Error " & CStr(Err.Number)
    On Error Resume Next
    LogResult "P4-CCMD", 0, "FAIL: " & FailMsg
    Debug.Print "FAIL  P4-CCMD - " & FailMsg
    KeepAlive Btn
    Phase4Bench_CanExecuteChanged = False
End Function

Public Function Phase4bBench_BeginUpdateDefer() As Boolean
    Dim Coll As ObservableCollection
    Dim Sink As Phase0CollectionSink
    Dim i As Long

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    Set Sink = New Phase0CollectionSink
    Sink.Attach Coll

    Coll.BeginUpdate
    For i = 1 To 100
        Coll.Add "item" & i
    Next
    Coll.EndUpdate

    If Coll.Count <> 100 Then Err.Raise vbObjectError, , "Expected 100 items after batch"
    If Sink.NotifyCount <> 1 Then Err.Raise vbObjectError, , "Expected 1 notification, got " & Sink.NotifyCount
    If Sink.LastAction <> CollectionChangedActionReset Then Err.Raise vbObjectError, , "Expected Reset notification"

    Sink.Detach
    LogResult "P4b-DEFER", 0, "OK BeginUpdate coalesced 100 adds"
    Debug.Print "PASS  P4b-DEFER BeginUpdate defers notifications"
    Phase4bBench_BeginUpdateDefer = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Sink Is Nothing Then Sink.Detach
    LogResult "P4b-DEFER", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4b-DEFER ? " & Err.Description
    Phase4bBench_BeginUpdateDefer = False
End Function

Public Function Phase4bBench_Move() As Boolean
    Dim Coll As ObservableCollection
    Dim Sink As Phase0CollectionSink

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    Coll.Add "a"
    Coll.Add "b"
    Coll.Add "c"

    Set Sink = New Phase0CollectionSink
    Sink.Attach Coll

    Coll.Move 0, 2

    If Coll(0) <> "b" Then Err.Raise vbObjectError, , "Index 0 expected b"
    If Coll(1) <> "c" Then Err.Raise vbObjectError, , "Index 1 expected c"
    If Coll(2) <> "a" Then Err.Raise vbObjectError, , "Index 2 expected a"
    If Sink.LastAction <> CollectionChangedActionMove Then Err.Raise vbObjectError, , "Expected Move notification"

    Sink.Detach
    LogResult "P4b-MOVE", 0, "OK Move(0,2)"
    Debug.Print "PASS  P4b-MOVE ObservableCollection Move"
    Phase4bBench_Move = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Sink Is Nothing Then Sink.Detach
    LogResult "P4b-MOVE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4b-MOVE ? " & Err.Description
    Phase4bBench_Move = False
End Function

Public Function Phase4bBench_ItemsControl() As Boolean
    Dim IC As ItemsControl
    Dim Coll As ObservableCollection
    Dim Tmpl As DataTemplate
    Dim Tb As TextBlock

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    Coll.Add "one"
    Coll.Add "two"

    Set Tmpl = New DataTemplate
    Set Tb = New TextBlock
    Tb.Text = "Item"
    Tmpl.Children.Add Tb

    Set IC = New ItemsControl
    Set IC.ItemTemplate = Tmpl
    Set IC.ItemsSource = Coll

    If IC.ItemCount <> 2 Then Err.Raise vbObjectError, , "Expected ItemCount=2"
    If IC.ItemsHost Is Nothing Then Err.Raise vbObjectError, , "ItemsHost is Nothing"
    If IC.ItemsHost.Children.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 generated items"

    Coll.Add "three"
    If IC.ItemCount <> 3 Then Err.Raise vbObjectError, , "Expected ItemCount=3 after Add"
    If IC.ItemsHost.Children.Count <> 3 Then Err.Raise vbObjectError, , "Expected 3 host children after Add"

    LogResult "P4b-ICtrl", 0, "OK ItemsControl generates items"
    Debug.Print "PASS  P4b-ICtrl ItemsControl ItemTemplate + ItemsSource"
    Phase4bBench_ItemsControl = True
    Exit Function

Fail:
    LogResult "P4b-ICtrl", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4b-ICtrl ? " & Err.Description
    Phase4bBench_ItemsControl = False
End Function

Public Function Phase4dBench_Selector() As Boolean
    Dim LV As ListView
    Dim Sel As Selector
    Dim Coll As ObservableCollection
    Dim Tmpl As DataTemplate
    Dim Tb As TextBlock
    Dim Bad As Panel

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    Coll.Add "alpha"
    Coll.Add "beta"
    Coll.Add "gamma"

    On Error GoTo FailNew
    Set LV = New ListView
    On Error GoTo FailSource
    Set LV.ItemsSource = Coll
    On Error GoTo FailIndex1
    LV.SelectedIndex = 1
    If LV.SelectedIndex <> 1 Then Err.Raise vbObjectError, , "ListView SelectedIndex expected 1"
    If LV.SelectedValue <> "beta" Then Err.Raise vbObjectError, , "ListView SelectedValue expected beta"

    On Error GoTo FailIndex2
    LV.SelectedIndex = 2
    If LV.SelectedValue <> "gamma" Then Err.Raise vbObjectError, , "ListView SelectedValue expected gamma"

    On Error GoTo FailSelector
    Set Tmpl = New DataTemplate
    Set Tb = New TextBlock
    Tb.Text = "Item"
    Tmpl.Children.Add Tb

    Set Sel = New Selector
    Set Sel.ItemTemplate = Tmpl
    Set Sel.ItemsSource = Coll
    Sel.SelectedIndex = 0
    If Sel.SelectedIndex <> 0 Then Err.Raise vbObjectError, , "Selector SelectedIndex expected 0"
    If Sel.SelectedValue <> "alpha" Then Err.Raise vbObjectError, , "Selector SelectedValue expected alpha"

    Set Bad = New Panel

    Dim BadErr As Long
    On Error Resume Next
    Set LV.ItemsSource = Bad
    BadErr = Err.Number
    Err.Clear
    If BadErr <> vbObjectError + 4 Then
        Err.Raise vbObjectError, , "ListView ItemsSource expected type error, got " & BadErr
    End If

    On Error Resume Next
    Set Sel.ItemsSource = Bad
    BadErr = Err.Number
    Err.Clear
    If BadErr <> vbObjectError + 4 Then
        Err.Raise vbObjectError, , "Selector ItemsSource expected type error, got " & BadErr
    End If

    On Error GoTo Fail

    LogResult "P4d-SEL", 0, "OK Selector DPs on ListView + Selector"
    Debug.Print "PASS  P4d-SEL Selector SelectedIndex/Value"
    Phase4dBench_Selector = True
    Exit Function

FailNew:
    LogResult "P4d-SEL", 0, "FAIL at New ListView: " & Err.Description
    Debug.Print "FAIL  P4d-SEL ? New ListView: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailSource:
    LogResult "P4d-SEL", 0, "FAIL at ListView ItemsSource: " & Err.Description
    Debug.Print "FAIL  P4d-SEL ? ListView ItemsSource: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailIndex1:
    LogResult "P4d-SEL", 0, "FAIL at ListView SelectedIndex=1: " & Err.Description
    Debug.Print "FAIL  P4d-SEL ? ListView SelectedIndex=1: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailIndex2:
    LogResult "P4d-SEL", 0, "FAIL at ListView SelectedIndex=2: " & Err.Description
    Debug.Print "FAIL  P4d-SEL ? ListView SelectedIndex=2: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailSelector:
    LogResult "P4d-SEL", 0, "FAIL at Selector: " & Err.Description
    Debug.Print "FAIL  P4d-SEL ? Selector: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

Fail:
    LogResult "P4d-SEL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4d-SEL ? " & Err.Description
    Phase4dBench_Selector = False
End Function

Public Function Phase4dBench_SelectionChanged() As Boolean
    Dim Host As Phase0SelectionHost
    Dim Coll As ObservableCollection

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    Coll.Add "alpha"
    Coll.Add "beta"
    Coll.Add "gamma"

    Set Host = New Phase0SelectionHost
    Host.Setup Coll
    KeepAlive Host

    Host.ResetCounts
    Host.ListView.Base.ListIndex = 1
    If Host.ListView.SelectedIndex <> 1 Then Err.Raise vbObjectError, , "SelectedIndex expected 1 after Base.ListIndex"
    If Host.SelectionChangedCount <> 1 Then Err.Raise vbObjectError, , "SelectionChanged expected 1, got " & Host.SelectionChangedCount
    If Host.ListIndexChangedCount <> 1 Then Err.Raise vbObjectError, , "ListIndexChanged expected 1, got " & Host.ListIndexChangedCount

    Host.ResetCounts
    Host.ListView.Base.ListIndex = 2
    If Host.ListView.SelectedIndex <> 2 Then Err.Raise vbObjectError, , "SelectedIndex expected 2 after Base.ListIndex"
    If Host.SelectionChangedCount <> 1 Then Err.Raise vbObjectError, , "SelectionChanged expected 1 on second change, got " & Host.SelectionChangedCount
    If Host.ListIndexChangedCount <> 1 Then Err.Raise vbObjectError, , "ListIndexChanged expected 1 on second change, got " & Host.ListIndexChangedCount

    LogResult "P4d-SELCHG", 0, "OK SelectionChanged + ListIndexChanged dual-raise"
    Debug.Print "PASS  P4d-SELCHG SelectionChanged naming"
    Phase4dBench_SelectionChanged = True
    Exit Function

Fail:
    LogResult "P4d-SELCHG", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4d-SELCHG ? " & Err.Description
    Phase4dBench_SelectionChanged = False
End Function

Public Function Phase5aBench_OwnerDrawListView() As Boolean
    Dim LV As ListView
    Dim Reader As XAMLReader
    Dim Root As Object

    On Error GoTo Fail

    Set LV = New ListView
    If Not LV.ItemsSource Is Nothing Then Err.Raise vbObjectError, , "Expected ItemsSource=Nothing for owner-draw"

    LV.Base.ListCount = 5
    If LV.Base.ListCount <> 5 Then Err.Raise vbObjectError, , "Expected ListCount=5"

    LV.SelectedIndex = 2
    If LV.SelectedIndex <> 2 Then Err.Raise vbObjectError, , "Expected SelectedIndex=2"
    If LV.Base.ListIndex <> 2 Then Err.Raise vbObjectError, , "Expected ListIndex=2"

    LV.Refresh

    Set Reader = New XAMLReader
    Set Root = Reader.Load("<UnboundListView/>")
    If Root Is Nothing Then Err.Raise vbObjectError, , "UnboundListView XAML alias failed"
    If Not TypeOf Root Is ListView Then Err.Raise vbObjectError, , "UnboundListView XAML must create ListView"

    LogResult "P5a-OWN", 0, "OK owner-draw ListView + XAML alias"
    Debug.Print "PASS  P5a-OWN owner-draw ListView + UnboundListView XAML alias"
    Phase5aBench_OwnerDrawListView = True
    Exit Function

Fail:
    LogResult "P5a-OWN", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P5a-OWN ? " & Err.Description
    Phase5aBench_OwnerDrawListView = False
End Function

Public Function Phase5bBench_MeasureRow() As Boolean
    Dim Host As Phase0MeasureRowHost

    On Error GoTo Fail

    Set Host = New Phase0MeasureRowHost
    Host.Setup 40, 20, 3

    If Host.MeasuredHeight(0) <> 40 Then Err.Raise vbObjectError, , "Expected row 0 height 40"
    If Host.MeasuredHeight(1) <> 20 Then Err.Raise vbObjectError, , "Expected row 1 height 20"
    If Host.MeasuredHeight(2) <> 20 Then Err.Raise vbObjectError, , "Expected row 2 height 20"

    LogResult "P5b-MSR", 0, "OK MeasureRow parent 40 / child 20"
    Debug.Print "PASS  P5b-MSR MeasureRow variable row heights"
    Phase5bBench_MeasureRow = True
    Exit Function

Fail:
    LogResult "P5b-MSR", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P5b-MSR ? " & Err.Description
    Phase5bBench_MeasureRow = False
End Function

Public Function Phase5cBench_RowLevel() As Boolean
    Dim Host As Phase0MeasureRowHost

    On Error GoTo Fail

    Set Host = New Phase0MeasureRowHost
    Host.Setup 40, 20, 3, 16

    If Host.MeasuredLevel(0) <> 0 Then Err.Raise vbObjectError, , "Expected row 0 level 0"
    If Host.MeasuredLevel(1) <> 1 Then Err.Raise vbObjectError, , "Expected row 1 level 1"
    If Host.MeasuredLevel(2) <> 1 Then Err.Raise vbObjectError, , "Expected row 2 level 1"
    If Host.MeasuredIndent(0) <> 0 Then Err.Raise vbObjectError, , "Expected row 0 indent 0"
    If Host.MeasuredIndent(1) <> 16 Then Err.Raise vbObjectError, , "Expected row 1 indent 16"
    If Host.MeasuredIndent(2) <> 16 Then Err.Raise vbObjectError, , "Expected row 2 indent 16"
    If Host.MeasuredHeight(0) <> 40 Then Err.Raise vbObjectError, , "Expected row 0 height 40 with indent"
    If Host.MeasuredHeight(1) <> 20 Then Err.Raise vbObjectError, , "Expected row 1 height 20 with indent"

    LogResult "P5c-HIER", 0, "OK QueryRowLevel parent/child indent"
    Debug.Print "PASS  P5c-HIER QueryRowLevel parent/child indent"
    Phase5cBench_RowLevel = True
    Exit Function

Fail:
    LogResult "P5c-HIER", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P5c-HIER ? " & Err.Description
    Phase5cBench_RowLevel = False
End Function

Public Function Phase6aBench_ButtonContent() As Boolean
    Dim Btn As Button
    Dim Vm As Phase0ViewModel
    Dim Expr As BindingExpression
    Dim Reader As XAMLReader
    Dim Root As Panel
    Dim XamlBtn As Button
    Dim LegacyBtn As Button
    Dim SavedStrict As Boolean

    On Error GoTo Fail

    SavedStrict = VCF.StrictXamlLoad

    Set Btn = New Button
    Btn.Content = "Static"
    If Btn.Content <> "Static" Then Err.Raise vbObjectError, , "Expected Static, got " & CStr(Btn.Content)

    Set Vm = New Phase0ViewModel
    Vm.Title = "Bound"
    Set Expr = New BindingExpression
    Expr.Attach Btn, "Content", Vm, "Title", OneWay

    If Btn.Content <> "Bound" Then Err.Raise vbObjectError, , "Expected Bound, got " & CStr(Btn.Content)

    Vm.Title = "Updated"
    If Btn.Content <> "Updated" Then Err.Raise vbObjectError, , "Expected Updated after INPC, got " & CStr(Btn.Content)

    Set Reader = New XAMLReader
    Set Root = Reader.Load("<Panel><Button Content=""OK""/></Panel>")
    If Root Is Nothing Then Err.Raise vbObjectError, , "Content XAML returned Nothing"
    Set XamlBtn = Root.Children(0)
    If XamlBtn.Content <> "OK" Then Err.Raise vbObjectError, , "Expected OK from Content XAML, got " & CStr(XamlBtn.Content)

    SavedStrict = VCF.StrictXamlLoad
    VCF.StrictXamlLoad = True
    Set LegacyBtn = Nothing
    Set Root = Reader.Load("<Panel><Button Text=""Legacy""/></Panel>")
    VCF.StrictXamlLoad = SavedStrict
    Set LegacyBtn = Root.Children(0)
    If LegacyBtn.Content <> "Legacy" Then Err.Raise vbObjectError, , "Expected Legacy from Text alias, got " & CStr(LegacyBtn.Content)

    LogResult "P6a-CONTENT", 0, "OK Content DP + Text alias + OneWay bind"
    Debug.Print "PASS  P6a-CONTENT Button Content DP"
    Expr.Detach
    Set Expr = Nothing
    Set Btn = Nothing
    Set Vm = Nothing
    Phase6aBench_ButtonContent = True
    Exit Function

Fail:
    On Error Resume Next
    VCF.StrictXamlLoad = SavedStrict
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P6a-CONTENT", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6a-CONTENT ? " & Err.Description
    Phase6aBench_ButtonContent = False
End Function

Public Function Phase6bBench_PropertyTrigger() As Boolean
    Dim St As Style
    Dim Btn As Button
    Dim Trig As PropertyTrigger
    Dim Tmpl As ControlTemplate
    Dim Chrome As Border
    Dim Pres As ContentPresenter
    Dim PtrBefore As Long
    Dim Rad As VCF.CornerRadius

    On Error GoTo Fail

    Set St = NewStyle("Button")
    St.SetSetter "BackColor", CLng(16777215)
    St.SetSetter "HoverColor", CLng(-1)

    Set Trig = New PropertyTrigger
    Trig.Initialize "IsMouseOver", "True"
    Trig.SetSetter "BackColor", CLng(255)
    St.AddTrigger Trig

    ' Live template: hover must not re-clone chrome (ReapplyStyleValues path).
    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"
    Set Chrome = New Border
    Rad.TopLeft = 4
    Rad.TopRight = 4
    Rad.BottomLeft = 4
    Rad.BottomRight = 4
    Chrome.CornerRadius = Rad
    Set Pres = New ContentPresenter
    Set Chrome.Child = Pres
    Tmpl.Children.Add Chrome
    Set St.Template = Tmpl

    Set Btn = New Button
    Set Btn.Style = St

    If Btn.Widget.BackColor <> 16777215 Then Err.Raise vbObjectError, , "Expected base BackColor 16777215, got " & Btn.Widget.BackColor
    If Btn.ContentPresenter Is Nothing Then Err.Raise vbObjectError, , "Expected live ContentPresenter after Style+Template"
    PtrBefore = ObjPtr(Btn.ContentPresenter)

    Btn.IsMouseOver = True
    If Btn.Widget.BackColor <> 255 Then Err.Raise vbObjectError, , "Expected hover BackColor 255, got " & Btn.Widget.BackColor
    If ObjPtr(Btn.ContentPresenter) <> PtrBefore Then Err.Raise vbObjectError, , "Hover re-cloned ContentPresenter (expected ReapplyStyleValues)"

    Btn.IsMouseOver = False
    If Btn.Widget.BackColor <> 16777215 Then Err.Raise vbObjectError, , "Expected restored BackColor 16777215, got " & Btn.Widget.BackColor
    If ObjPtr(Btn.ContentPresenter) <> PtrBefore Then Err.Raise vbObjectError, , "Unhover re-cloned ContentPresenter"

    ' DP condition path: NotifyConditionPropertyChanged from DependencyProperties (Selected).
    Set Trig = New PropertyTrigger
    Trig.Initialize "Selected", "True"
    Trig.SetSetter "BackColor", CLng(65280)
    St.AddTrigger Trig
    ' StyleChanged re-runs full ApplyStyle (template ok once).
    Set Btn.Style = St

    If Btn.Widget.BackColor <> 16777215 Then Err.Raise vbObjectError, , "Expected base after Selected trigger add"
    Btn.Selected = True
    If Btn.Widget.BackColor <> 65280 Then Err.Raise vbObjectError, , "Expected Selected trigger BackColor 65280, got " & Btn.Widget.BackColor
    Btn.Selected = False
    If Btn.Widget.BackColor <> 16777215 Then Err.Raise vbObjectError, , "Expected restore after Selected=False, got " & Btn.Widget.BackColor

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Trig = Nothing
    Set Tmpl = Nothing
    Set Chrome = Nothing
    Set Pres = Nothing

    LogResult "P6b-TRIG", 0, "OK IsMouseOver+Selected trigger Notify/ReapplyStyleValues"
    Debug.Print "PASS  P6b-TRIG Style PropertyTrigger IsMouseOver"
    Phase6bBench_PropertyTrigger = True
    Exit Function

Fail:
    LogResult "P6b-TRIG", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6b-TRIG - " & Err.Description
    Phase6bBench_PropertyTrigger = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

Public Function Phase6cBench_ControlTemplate() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim Rad As VCF.CornerRadius
    Dim Reader As XAMLReader
    Dim Dom As cSimpleDOM
    Dim XamlTmpl As ControlTemplate

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"
    Set B = New Border
    Rad.TopLeft = 12
    Rad.TopRight = 12
    Rad.BottomLeft = 12
    Rad.BottomRight = 12
    B.CornerRadius = Rad
    Tmpl.Children.Add B

    Set St = NewStyle("Button")
    Set St.Template = Tmpl

    Set Btn = New Button
    Set Btn.Style = St

    If Btn.CornerRadius <> 12# Then Err.Raise vbObjectError, , "Expected CornerRadius 12, got " & Btn.CornerRadius

    Set Reader = New XAMLReader
    Set Dom = New_c.SimpleDOM
    Dom.Xml = "<ControlTemplate TargetType=""Button""><Border CornerRadius=""12""/></ControlTemplate>"
    If Not Dom.WellFormed Then Err.Raise vbObjectError, , "ControlTemplate XAML not well formed"
    Set XamlTmpl = Reader.LoadElement(Dom.Root)
    If XamlTmpl Is Nothing Then Err.Raise vbObjectError, , "ControlTemplate XAML load returned Nothing"
    If XamlTmpl.Children.Count = 0 Then Err.Raise vbObjectError, , "ControlTemplate XAML has no visual child"

    Set St = NewStyle("Button")
    Set St.Template = XamlTmpl
    Set Btn = New Button
    Set Btn.Style = St
    If Btn.CornerRadius <> 12# Then Err.Raise vbObjectError, , "Expected XAML template CornerRadius 12, got " & Btn.CornerRadius

    LogResult "P6c-TMPL", 0, "OK ControlTemplate Border chrome on Button"
    Debug.Print "PASS  P6c-TMPL ControlTemplate Button chrome"
    Phase6cBench_ControlTemplate = True
    Exit Function

Fail:
    LogResult "P6c-TMPL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6c-TMPL ? " & Err.Description
    Phase6cBench_ControlTemplate = False
End Function

Public Function Phase6dBench_RenderCoalesce() As Boolean
    Dim Btn As Button
    Dim St As Style
    Dim RC As RenderCoalescer
    Dim i As Long

    On Error GoTo Fail

    Set RC = New RenderCoalescer
    Set Btn = New Button

    RC.BeginRenderUpdate
    For i = 1 To 20
        RC.RequestWidgetRefresh Btn.Widget
    Next

    If RC.PendingCount <> 1 Then
        Err.Raise vbObjectError, , "Expected 1 pending refresh, got " & RC.PendingCount
    End If

    RC.EndRenderUpdate

    If RC.LastFlushCount <> 1 Then
        Err.Raise vbObjectError, , "Expected flush count 1, got " & RC.LastFlushCount
    End If

    Set St = NewStyle("Button")
    St.SetSetter "BackColor", CLng(16777215)
    St.SetSetter "BorderColor", CLng(255)
    St.SetSetter "ToolTip", "coalesce"

    RC.BeginRenderUpdate
    Set Btn.Style = St

    If RC.PendingCount <> 1 Then
        Err.Raise vbObjectError, , "Expected 1 pending refresh after nested style apply, got " & RC.PendingCount
    End If

    RC.EndRenderUpdate

    If RC.LastFlushCount <> 1 Then
        Err.Raise vbObjectError, , "Expected style batch flush count 1, got " & RC.LastFlushCount
    End If

    LogResult "P6d-COAL", 0, "OK BeginRenderUpdate dedupe + StyleManager batch"
    Debug.Print "PASS  P6d-COAL render refresh coalescing"
    Phase6dBench_RenderCoalesce = True
    Exit Function

Fail:
    LogResult "P6d-COAL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6d-COAL ? " & Err.Description
    Phase6dBench_RenderCoalesce = False
End Function

' ?2.11 ContentPresenter paint-only path (Button caption delegates here).
Public Function Phase6eBench_ContentPresenter() As Boolean
    Dim CP As ContentPresenter
    Dim Btn As Button
    Dim Tb As TextBlock

    On Error GoTo Fail

    Set CP = New ContentPresenter
    If CP Is Nothing Then Err.Raise vbObjectError, , "ContentPresenter New returned Nothing"
    CP.Content = "OK"
    If CStr(CP.Content) <> "OK" Then Err.Raise vbObjectError, , "Content round-trip expected OK"
    If CP.ContentCaption <> "OK" Then Err.Raise vbObjectError, , "ContentCaption expected OK"
    If CP.SuppressContent Then Err.Raise vbObjectError, , "SuppressContent default expected False"
    If Not CP.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected True for OK"
    CP.SuppressContent = True
    If CP.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected False when suppressed"

    Set Btn = New Button
    If Not TypeOf Btn Is IContentControl Then Err.Raise vbObjectError, , "Button should Implement IContentControl"
    If Btn.ContentPresenter Is Nothing Then Err.Raise vbObjectError, , "Button.ContentPresenter is Nothing"
    Btn.Content = "Save"
    Btn.SyncContentPresenter
    If CStr(Btn.ContentPresenter.Content) <> "Save" Then Err.Raise vbObjectError, , "Presenter Content expected Save"
    If Btn.ContentPresenter.SuppressContent Then Err.Raise vbObjectError, , "Suppress expected False with no children"
    If Not Btn.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected True for Save"

    Set Tb = New TextBlock
    Tb.Text = "child"
    Btn.Children.Add Tb
    Btn.SyncContentPresenter
    If Not Btn.ContentPresenter.SuppressContent Then Err.Raise vbObjectError, , "Suppress expected True with child"
    If Btn.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected False with child"

    KeepAlive Btn
    Set Tb = Nothing
    Set Btn = Nothing
    Set CP = Nothing

    LogResult "P6e-PRES", 0, "OK ContentPresenter paint-only + Button suppress"
    Debug.Print "PASS  P6e-PRES ContentPresenter paint path"
    Phase6eBench_ContentPresenter = True
    Exit Function

Fail:
    LogResult "P6e-PRES", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6e-PRES ? " & Err.Description
    Phase6eBench_ContentPresenter = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' ?2.11 ContentPresenter / Button HorizontalContentAlignment + VerticalContentAlignment.
Public Function Phase6eBench_ContentAlignment() As Boolean
    Dim CP As ContentPresenter
    Dim Btn As Button
    Dim Reader As XAMLReader
    Dim Root As Button

    On Error GoTo Fail

    Set CP = New ContentPresenter
    If CP.HorizontalContentAlignment <> AlignmentConstants.vbCenter Then
        Err.Raise vbObjectError, , "Presenter HAlign default expected Center"
    End If
    If CP.VerticalContentAlignment <> 2 Then
        Err.Raise vbObjectError, , "Presenter VAlign default expected Center(2)"
    End If
    CP.HorizontalContentAlignment = AlignmentConstants.vbLeftJustify
    CP.VerticalContentAlignment = 0
    If CP.HorizontalContentAlignment <> AlignmentConstants.vbLeftJustify Then
        Err.Raise vbObjectError, , "Presenter HAlign expected Left"
    End If
    If CP.VerticalContentAlignment <> 0 Then
        Err.Raise vbObjectError, , "Presenter VAlign expected Top(0)"
    End If

    Set Btn = New Button
    If Btn.HorizontalContentAlignment <> AlignmentConstants.vbCenter Then
        Err.Raise vbObjectError, , "Button HAlign default expected Center"
    End If
    If Btn.VerticalContentAlignment <> 2 Then
        Err.Raise vbObjectError, , "Button VAlign default expected Center(2)"
    End If
    Btn.Content = "Go"
    Btn.HorizontalContentAlignment = AlignmentConstants.vbRightJustify
    Btn.VerticalContentAlignment = 1
    Btn.SyncContentPresenter
    If Btn.ContentPresenter.HorizontalContentAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Presenter HAlign expected Right after sync"
    End If
    If Btn.ContentPresenter.VerticalContentAlignment <> 1 Then
        Err.Raise vbObjectError, , "Presenter VAlign expected Bottom(1) after sync"
    End If
    If Not Btn.ContentPresenter.WouldDrawCaption Then
        Err.Raise vbObjectError, , "WouldDrawCaption expected True"
    End If

    Set Reader = New XAMLReader
    Set Root = Reader.Load( _
        "<Button Content=""X"" HorizontalContentAlignment=""0"" VerticalContentAlignment=""0""/>")
    If Root Is Nothing Then Err.Raise vbObjectError, , "XAML Button returned Nothing"
    If Root.HorizontalContentAlignment <> AlignmentConstants.vbLeftJustify Then
        Err.Raise vbObjectError, , "XAML HAlign expected Left(0) got " & Root.HorizontalContentAlignment
    End If
    If Root.VerticalContentAlignment <> 0 Then
        Err.Raise vbObjectError, , "XAML VAlign expected 0 got " & Root.VerticalContentAlignment
    End If

    KeepAlive Btn
    KeepAlive Root
    Set CP = Nothing
    Set Btn = Nothing
    Set Root = Nothing
    Set Reader = Nothing

    LogResult "P6e-ALIGN", 0, "OK ContentAlignment H/V on presenter + Button + XAML"
    Debug.Print "PASS  P6e-ALIGN ContentAlignment H/V"
    Phase6eBench_ContentAlignment = True
    Exit Function

Fail:
    LogResult "P6e-ALIGN", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6e-ALIGN ? " & Err.Description
    Phase6eBench_ContentAlignment = False
    On Error Resume Next
    KeepAlive Btn
    KeepAlive Root
    Err.Clear
End Function

' ?2.11 ContentControl shares Button Content model (string presenter + IUIElement child).
Public Function Phase6eBench_ContentControlContent() As Boolean
    Dim CC As ContentControl
    Dim Tb As TextBlock
    Dim Reader As XAMLReader
    Dim Root As ContentControl
    Dim Child As Object

    On Error GoTo Fail

    Set CC = New ContentControl
    If Not TypeOf CC Is IContentControl Then Err.Raise vbObjectError, , "ContentControl should Implement IContentControl"
    If CC.ContentPresenter Is Nothing Then Err.Raise vbObjectError, , "ContentControl.ContentPresenter is Nothing"

    CC.Content = "Hello"
    CC.SyncContentPresenter
    If CStr(CC.Content) <> "Hello" Then Err.Raise vbObjectError, , "Content string expected Hello"
    If Not CC.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected True for string"
    If CC.Children.Count <> 0 Then Err.Raise vbObjectError, , "String Content should not add children"

    Set Tb = New TextBlock
    Tb.Text = "Child"
    Set CC.Content = Tb
    CC.SyncContentPresenter
    If CC.Children.Count <> 1 Then Err.Raise vbObjectError, , "Object Content expected 1 child"
    Set Child = CC.Content
    If Not Child Is Tb Then Err.Raise vbObjectError, , "Content Get expected TextBlock child"
    If Not CC.ContentPresenter.SuppressContent Then Err.Raise vbObjectError, , "Suppress expected True with child"
    If CC.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected False with child"

    Set Reader = New XAMLReader
    Set Root = Reader.Load("<ContentControl Content=""Hi"" Width=""100"" Height=""30""/>")
    If Root Is Nothing Then Err.Raise vbObjectError, , "XAML ContentControl returned Nothing"
    If CStr(Root.Content) <> "Hi" Then Err.Raise vbObjectError, , "XAML Content expected Hi"
    Root.SyncContentPresenter
    If Not Root.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "XAML WouldDrawCaption expected True"

    KeepAlive CC
    KeepAlive Root
    Set Tb = Nothing
    Set Child = Nothing
    Set CC = Nothing
    Set Root = Nothing
    Set Reader = Nothing

    LogResult "P6e-CC", 0, "OK ContentControl Content string+IUIElement + XAML"
    Debug.Print "PASS  P6e-CC ContentControl Content unify"
    Phase6eBench_ContentControlContent = True
    Exit Function

Fail:
    LogResult "P6e-CC", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6e-CC ? " & Err.Description
    Phase6eBench_ContentControlContent = False
    On Error Resume Next
    KeepAlive CC
    KeepAlive Root
    Err.Clear
End Function

Public Function Phase6eBench_ContentControlMeasure() As Boolean
    Dim CC As ContentControl
    Dim Child As Panel

    On Error GoTo Fail

    Set CC = New ContentControl
    CC.Width = 0
    CC.Height = 0
    CC.Widget.Move 0, 0, 200, 100

    Set Child = New Panel
    Child.Width = 80
    Child.Height = 40
    Set CC.Content = Child

    If CC.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 Content child, got " & CC.Children.Count
    If Not CC.Children(0) Is Child Then Err.Raise vbObjectError, , "Children(0) is not Content Panel"
    If Abs(Child.Width - 80#) > 0.5 Then Err.Raise vbObjectError, , "Child.Width expected 80, got " & Child.Width

    CC.MeasureLayout 200, 100
    If Abs(CC.DesiredWidth - 80#) > 0.5 Then Err.Raise vbObjectError, , "after MeasureLayout DesiredWidth expected 80, got " & CC.DesiredWidth

    CC.RelayoutChildren
    If Abs(CC.DesiredWidth - 80#) > 0.5 Then Err.Raise vbObjectError, , "after Relayout DesiredWidth expected 80, got " & CC.DesiredWidth
    If Abs(CC.DesiredHeight - 40#) > 0.5 Then Err.Raise vbObjectError, , "DesiredHeight expected 40, got " & CC.DesiredHeight
    If Abs(CC.ActualWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualWidth expected 200, got " & CC.ActualWidth
    If Abs(CC.ActualHeight - 100#) > 0.5 Then Err.Raise vbObjectError, , "ActualHeight expected 100, got " & CC.ActualHeight
    If Abs(Child.Widget.Left - 0!) > 1! Then Err.Raise vbObjectError, , "Child.Left expected 0, got " & Child.Widget.Left
    If Abs(Child.Widget.Top - 0!) > 1! Then Err.Raise vbObjectError, , "Child.Top expected 0, got " & Child.Widget.Top

    KeepAlive CC
    LogResult "P6e-CC-MEAS", 0, "OK Desired=" & CC.DesiredWidth & "x" & CC.DesiredHeight
    Debug.Print "PASS  P6e-CC-MEAS ContentControl Measure/Actual"
    Phase6eBench_ContentControlMeasure = True
    Exit Function

Fail:
    LogResult "P6e-CC-MEAS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6e-CC-MEAS - " & Err.Description
    Phase6eBench_ContentControlMeasure = False
    On Error Resume Next
    KeepAlive CC
    Err.Clear
End Function
Public Function Phase6eBench_ContentTemplate() As Boolean
    Dim CC As ContentControl
    Dim Tmpl As DataTemplate
    Dim Tb As TextBlock
    Dim Child As TextBlock
    Dim Explicit As TextBlock

    On Error GoTo Fail

    Set CC = New ContentControl
    Set Tmpl = New DataTemplate
    Set Tb = New TextBlock
    Tb.Text = "FromTemplate"
    Tb.Width = 80
    Tb.Height = 20
    Tmpl.Children.Add Tb

    Set CC.ContentTemplate = Tmpl
    CC.Content = "payload"

    If CC.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 generated child, got " & CC.Children.Count
    If TypeName(CC.Children(0)) <> "TextBlock" Then Err.Raise vbObjectError, , "Expected TextBlock, got " & TypeName(CC.Children(0))
    Set Child = CC.Children(0)
    If Child.Text <> "FromTemplate" Then Err.Raise vbObjectError, , "Generated Text expected FromTemplate, got " & Child.Text
    CC.SyncContentPresenter
    If Not CC.ContentPresenter.SuppressContent Then Err.Raise vbObjectError, , "Suppress expected True with ContentTemplate child"
    If CC.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected False with ContentTemplate child"

    ' Explicit child wins over ContentTemplate.
    Set Explicit = New TextBlock
    Explicit.Text = "Explicit"
    Explicit.Width = 40
    Explicit.Height = 16
    CC.Children.Add Explicit
    CC.Content = "again"
    If CC.Children.Count < 1 Then Err.Raise vbObjectError, , "Expected children after explicit add"
    CC.SyncContentPresenter
    If Not CC.ContentPresenter.SuppressContent Then Err.Raise vbObjectError, , "Suppress expected True with explicit child"

    ' No template: string Content paints via presenter.
    Set CC = New ContentControl
    CC.Content = "Plain"
    CC.SyncContentPresenter
    If CC.Children.Count <> 0 Then Err.Raise vbObjectError, , "Plain Content should not create children"
    If Not CC.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected True for Plain"

    LogResult "P6e-CTMPL", 0, "OK ContentControl ContentTemplate clone + precedence"
    Debug.Print "PASS  P6e-CTMPL ContentTemplate"
    Phase6eBench_ContentTemplate = True
    Exit Function

Fail:
    LogResult "P6e-CTMPL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6e-CTMPL - " & Err.Description
    Phase6eBench_ContentTemplate = False
End Function

' Lookless: ControlTemplate Border chrome + ContentPresenter marker (+ live clone gated in P6g).
Public Function Phase6fBench_TemplateBindingSlot() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim CP As ContentPresenter
    Dim Rad As VCF.CornerRadius

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"

    Set B = New Border
    Rad.TopLeft = 8
    Rad.TopRight = 8
    Rad.BottomLeft = 8
    Rad.BottomRight = 8
    B.CornerRadius = Rad
    Tmpl.Children.Add B

    Set CP = New ContentPresenter
    ' Use Right(1) not Left(0) first so a failed apply cannot be confused with Long default 0.
    CP.HorizontalContentAlignment = AlignmentConstants.vbRightJustify
    CP.VerticalContentAlignment = 0
    Tmpl.Children.Add CP
    ' Explicit slot ? TypeOf ContentPresenter is unreliable across EXE/DLL boundary.
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbRightJustify, 0
    If Tmpl.Children.Count <> 2 Then Err.Raise vbObjectError, , "Template expected Border+ContentPresenter children, count=" & Tmpl.Children.Count
    If Not Tmpl.HasContentAlignmentMarker Then Err.Raise vbObjectError, , "HasContentAlignmentMarker expected True"
    If Tmpl.ContentHorizontalAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Template slot HAlign not Right before Style apply"
    End If

    Set St = NewStyle("Button")
    Set St.Template = Tmpl
    If Not St.Template Is Tmpl Then Err.Raise vbObjectError, , "Style.Template must be same instance as Tmpl"
    If Not St.Template.HasContentAlignmentMarker Then Err.Raise vbObjectError, , "Style.Template marker missing"

    Set Btn = New Button
    Btn.Content = "OK"
    Set Btn.Style = St
    ' Align comes from StyleManager.PushTemplateContentAlignment (in-style path).

    If Btn.CornerRadius <> 8# Then Err.Raise vbObjectError, , "Expected CornerRadius 8, got " & Btn.CornerRadius
    If Btn.HorizontalContentAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Expected HAlign Right from template marker, got " & Btn.HorizontalContentAlignment & _
            " (slot=" & Tmpl.ContentHorizontalAlignment & ")"
    End If
    If Btn.VerticalContentAlignment <> 0 Then
        Err.Raise vbObjectError, , "Expected VAlign Top(0) from template marker, got " & Btn.VerticalContentAlignment
    End If

    ' Second pass: Left(0) via marker; clear Style so SetValue re-fires ApplyStyle.
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbLeftJustify, 0
    Set Btn.Style = Nothing
    Set Btn.Style = St
    If Btn.HorizontalContentAlignment <> AlignmentConstants.vbLeftJustify Then
        Err.Raise vbObjectError, , "Expected HAlign Left after marker update, got " & Btn.HorizontalContentAlignment
    End If
    Btn.SyncContentPresenter
    If Not Btn.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected True"
    If CStr(Btn.Content) <> "OK" Then Err.Raise vbObjectError, , "Content expected OK (TemplateBinding from parent)"

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Tmpl = Nothing
    Set B = Nothing
    Set CP = Nothing

    LogResult "P6f-TBIND", 0, "OK ControlTemplate Border+ContentPresenter marker"
    Debug.Print "PASS  P6f-TBIND template ContentPresenter slot"
    Phase6fBench_TemplateBindingSlot = True
    Exit Function

Fail:
    LogResult "P6f-TBIND", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6f-TBIND ? " & Err.Description
    Phase6fBench_TemplateBindingSlot = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' ?2.11 lookless: live cloned Border under Button.Children; caption not suppressed.
Public Function Phase6gBench_LiveTemplateChrome() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim Live As Border
    Dim CP As ContentPresenter
    Dim Rad As VCF.CornerRadius
    Dim LiveRad As VCF.CornerRadius

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"

    Set B = New Border
    Rad.TopLeft = 8
    Rad.TopRight = 8
    Rad.BottomLeft = 8
    Rad.BottomRight = 8
    B.CornerRadius = Rad
    Tmpl.Children.Add B

    Set CP = New ContentPresenter
    CP.HorizontalContentAlignment = AlignmentConstants.vbRightJustify
    CP.VerticalContentAlignment = 0
    Tmpl.Children.Add CP
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbRightJustify, 0

    Set St = NewStyle("Button")
    Set St.Template = Tmpl

    Set Btn = New Button
    Btn.Content = "OK"
    Set Btn.Style = St

    If Btn.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 live child (template Border), got " & Btn.Children.Count
    If TypeName(Btn.Children(0)) <> "Border" Then Err.Raise vbObjectError, , "Expected live Border, got " & TypeName(Btn.Children(0))
    Set Live = Btn.Children(0)
    If Live Is B Then Err.Raise vbObjectError, , "Live Border must be a clone, not template-bag instance"
    If Btn.CornerRadius <> 8# Then Err.Raise vbObjectError, , "Host CornerRadius expected 8, got " & Btn.CornerRadius
    Call API.CopyVariable(Live.DependencyProperties.GetValue("CornerRadius"), LiveRad)
    If LiveRad.TopLeft <> 8# Then Err.Raise vbObjectError, , "Live Border CornerRadius expected 8, got " & LiveRad.TopLeft

    Btn.SyncContentPresenter
    If Btn.ContentPresenter.SuppressContent Then Err.Raise vbObjectError, , "SuppressContent must be False with template chrome only"
    If Not Btn.ContentPresenter.WouldDrawCaption Then Err.Raise vbObjectError, , "WouldDrawCaption expected True with live template chrome"

    ' Re-apply must replace chrome without stacking children.
    Set Btn.Style = Nothing
    If Btn.Children.Count <> 0 Then Err.Raise vbObjectError, , "Clear Style must remove live template chrome, count=" & Btn.Children.Count
    Set Btn.Style = St
    If Btn.Children.Count <> 1 Then Err.Raise vbObjectError, , "Re-apply expected 1 live child, got " & Btn.Children.Count
    If TypeName(Btn.Children(0)) <> "Border" Then Err.Raise vbObjectError, , "Re-apply expected Border"

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Tmpl = Nothing
    Set B = Nothing
    Set Live = Nothing
    Set CP = Nothing

    LogResult "P6g-LIVE", 0, "OK live ControlTemplate Border chrome + caption"
    Debug.Print "PASS  P6g-LIVE live template Border chrome"
    Phase6gBench_LiveTemplateChrome = True
    Exit Function

Fail:
    LogResult "P6g-LIVE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6g-LIVE ? " & Err.Description
    Phase6gBench_LiveTemplateChrome = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' ?2.11 lookless: live paint-only ContentPresenter slot + Content TemplateBinding.
Public Function Phase6hBench_LiveContentPresenter() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim CP As ContentPresenter
    Dim LiveCP As ContentPresenter
    Dim HostCP As ContentPresenter
    Dim Rad As VCF.CornerRadius

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"

    Set B = New Border
    Rad.TopLeft = 8
    Rad.TopRight = 8
    Rad.BottomLeft = 8
    Rad.BottomRight = 8
    B.CornerRadius = Rad
    Tmpl.Children.Add B

    Set CP = New ContentPresenter
    CP.HorizontalContentAlignment = AlignmentConstants.vbRightJustify
    CP.VerticalContentAlignment = 0
    Tmpl.Children.Add CP
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbRightJustify, 0

    Set St = NewStyle("Button")
    Set St.Template = Tmpl

    Set Btn = New Button
    Btn.Content = "OK"
    Set Btn.Style = St

    Set LiveCP = Btn.ContentPresenter
    If LiveCP Is Nothing Then Err.Raise vbObjectError, , "Live ContentPresenter expected"
    If LiveCP Is CP Then Err.Raise vbObjectError, , "Live ContentPresenter must be a clone, not template-bag instance"
    Btn.SyncContentPresenter
    If CStr(LiveCP.Content) <> "OK" Then Err.Raise vbObjectError, , "TemplateBinding Content expected OK, got " & CStr(LiveCP.Content)
    If LiveCP.HorizontalContentAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Live CP HAlign expected Right, got " & LiveCP.HorizontalContentAlignment
    End If
    If LiveCP.VerticalContentAlignment <> 0 Then
        Err.Raise vbObjectError, , "Live CP VAlign expected Top(0), got " & LiveCP.VerticalContentAlignment
    End If
    If LiveCP.SuppressContent Then Err.Raise vbObjectError, , "Live CP SuppressContent expected False"
    If Not LiveCP.WouldDrawCaption Then Err.Raise vbObjectError, , "Live CP WouldDrawCaption expected True"

    ' Content TemplateBinding: host Content change flows to live slot.
    Btn.Content = "Save"
    Btn.SyncContentPresenter
    If CStr(Btn.ContentPresenter.Content) <> "Save" Then
        Err.Raise vbObjectError, , "TemplateBinding Content expected Save after change"
    End If

    ' Clear Style drops live slot; host presenter resumes.
    Set Btn.Style = Nothing
    Set HostCP = Btn.ContentPresenter
    If HostCP Is LiveCP Then Err.Raise vbObjectError, , "After clear Style, ContentPresenter must fall back to host slot"
    Btn.SyncContentPresenter
    If CStr(HostCP.Content) <> "Save" Then Err.Raise vbObjectError, , "Host CP Content expected Save after clear"
    If Not HostCP.WouldDrawCaption Then Err.Raise vbObjectError, , "Host WouldDrawCaption expected True after clear"

    ' Re-apply restores a fresh live clone.
    Set Btn.Style = St
    If Btn.ContentPresenter Is HostCP Then Err.Raise vbObjectError, , "Re-apply expected new live ContentPresenter"
    If Btn.ContentPresenter Is CP Then Err.Raise vbObjectError, , "Re-apply live CP must not be template-bag instance"
    Btn.SyncContentPresenter
    If CStr(Btn.ContentPresenter.Content) <> "Save" Then Err.Raise vbObjectError, , "Re-apply TemplateBinding Content expected Save"
    If Btn.ContentPresenter.HorizontalContentAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Re-apply live CP HAlign expected Right"
    End If

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Tmpl = Nothing
    Set B = Nothing
    Set CP = Nothing
    Set LiveCP = Nothing
    Set HostCP = Nothing

    LogResult "P6h-CP", 0, "OK live ContentPresenter TemplateBinding Content+align"
    Debug.Print "PASS  P6h-CP live ContentPresenter TemplateBinding"
    Phase6hBench_LiveContentPresenter = True
    Exit Function

Fail:
    LogResult "P6h-CP", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6h-CP - " & Err.Description
    Phase6hBench_LiveContentPresenter = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' Nested lookless: ContentPresenter widget under live Border.Child (WPF visual tree).
Public Function Phase6iBench_NestedContentPresenter() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim Live As Border
    Dim CP As ContentPresenter
    Dim Nested As ContentPresenter
    Dim Rad As VCF.CornerRadius

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"

    Set B = New Border
    Rad.TopLeft = 8
    Rad.TopRight = 8
    Rad.BottomLeft = 8
    Rad.BottomRight = 8
    B.CornerRadius = Rad
    Tmpl.Children.Add B

    Set CP = New ContentPresenter
    CP.HorizontalContentAlignment = AlignmentConstants.vbRightJustify
    CP.VerticalContentAlignment = 0
    Tmpl.Children.Add CP
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbRightJustify, 0

    Set St = NewStyle("Button")
    Set St.Template = Tmpl

    Set Btn = New Button
    Btn.Content = "OK"
    Set Btn.Style = St

    If Btn.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 live Border child, got " & Btn.Children.Count
    Set Live = Btn.Children(0)
    If Live.Child Is Nothing Then Err.Raise vbObjectError, , "Live Border.Child expected ContentPresenter"
    If TypeName(Live.Child) <> "ContentPresenter" Then Err.Raise vbObjectError, , "Border.Child expected ContentPresenter, got " & TypeName(Live.Child)
    Set Nested = Live.Child
    If Nested Is CP Then Err.Raise vbObjectError, , "Nested CP must be a clone, not template-bag instance"
    If Not Btn.ContentPresenter Is Nested Then Err.Raise vbObjectError, , "Button.ContentPresenter must be nested live slot"
    If Nested.Parent Is Nothing Then Err.Raise vbObjectError, , "Nested CP Parent expected (Border)"
    If Not Nested.Parent Is Live Then Err.Raise vbObjectError, , "Nested CP Parent expected live Border"

    Btn.SyncContentPresenter
    If CStr(Nested.Content) <> "OK" Then Err.Raise vbObjectError, , "TemplateBinding Content expected OK"
    If Nested.HorizontalContentAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Nested CP HAlign expected Right"
    End If
    If Nested.SuppressContent Then Err.Raise vbObjectError, , "Nested CP SuppressContent expected False"
    If Not Nested.WouldDrawCaption Then Err.Raise vbObjectError, , "Nested CP WouldDrawCaption expected True"

    Btn.Content = "Save"
    Btn.SyncContentPresenter
    If CStr(Nested.Content) <> "Save" Then Err.Raise vbObjectError, , "TemplateBinding Content expected Save"

    Set Btn.Style = Nothing
    If Btn.Children.Count <> 0 Then Err.Raise vbObjectError, , "Clear Style must remove live template tree"
    If Btn.ContentPresenter Is Nested Then Err.Raise vbObjectError, , "After clear, ContentPresenter must fall back to host"

    Set Btn.Style = St
    Set Live = Btn.Children(0)
    If Live.Child Is Nothing Then Err.Raise vbObjectError, , "Re-apply Border.Child expected"
    If TypeName(Live.Child) <> "ContentPresenter" Then Err.Raise vbObjectError, , "Re-apply Border.Child expected ContentPresenter"

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Tmpl = Nothing
    Set B = Nothing
    Set Live = Nothing
    Set CP = Nothing
    Set Nested = Nothing

    LogResult "P6i-NEST", 0, "OK nested Border.Child ContentPresenter visual tree"
    Debug.Print "PASS  P6i-NEST nested Border.Child ContentPresenter"
    Phase6iBench_NestedContentPresenter = True
    Exit Function

Fail:
    LogResult "P6i-NEST", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6i-NEST - " & Err.Description
    Phase6iBench_NestedContentPresenter = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' Deeper lookless: Border.Child = Grid with ContentPresenter (multi-node template).
Public Function Phase6jBench_MultiNodeTemplate() As Boolean
    Dim Tmpl As ControlTemplate
    Dim St As Style
    Dim Btn As Button
    Dim B As Border
    Dim G As Grid
    Dim CP As ContentPresenter
    Dim Live As Border
    Dim LiveGrid As Grid
    Dim Nested As ContentPresenter
    Dim Rad As VCF.CornerRadius

    On Error GoTo Fail

    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"

    Set B = New Border
    Rad.TopLeft = 6
    Rad.TopRight = 6
    Rad.BottomLeft = 6
    Rad.BottomRight = 6
    B.CornerRadius = Rad

    Set G = New Grid
    Set CP = New ContentPresenter
    CP.HorizontalContentAlignment = AlignmentConstants.vbRightJustify
    CP.VerticalContentAlignment = 0
    G.Children.Add CP
    Set B.Child = G
    Tmpl.Children.Add B
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbRightJustify, 0

    Set St = NewStyle("Button")
    Set St.Template = Tmpl

    Set Btn = New Button
    Btn.Content = "OK"
    Set Btn.Style = St

    If Btn.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 live Border, got " & Btn.Children.Count
    Set Live = Btn.Children(0)
    If Live.Child Is Nothing Then Err.Raise vbObjectError, , "Live Border.Child expected Grid"
    If TypeName(Live.Child) <> "Grid" Then Err.Raise vbObjectError, , "Border.Child expected Grid, got " & TypeName(Live.Child)
    Set LiveGrid = Live.Child
    If LiveGrid.Children.Count <> 1 Then Err.Raise vbObjectError, , "Grid expected 1 child, got " & LiveGrid.Children.Count
    If TypeName(LiveGrid.Children(0)) <> "ContentPresenter" Then Err.Raise vbObjectError, , "Grid child expected ContentPresenter"
    Set Nested = LiveGrid.Children(0)
    If Nested Is CP Then Err.Raise vbObjectError, , "Live CP must be clone"
    If Not Btn.ContentPresenter Is Nested Then Err.Raise vbObjectError, , "Button.ContentPresenter must be deep live slot"

    Btn.SyncContentPresenter
    If CStr(Nested.Content) <> "OK" Then Err.Raise vbObjectError, , "TemplateBinding Content expected OK"
    If Nested.HorizontalContentAlignment <> AlignmentConstants.vbRightJustify Then
        Err.Raise vbObjectError, , "Deep CP HAlign expected Right"
    End If
    If Not Nested.WouldDrawCaption Then Err.Raise vbObjectError, , "Deep CP WouldDrawCaption expected True"

    ' Nested-only Border.Child=CP (no Grid) still works.
    Set Tmpl = New ControlTemplate
    Tmpl.TargetType = "Button"
    Set B = New Border
    B.CornerRadius = Rad
    Set CP = New ContentPresenter
    CP.HorizontalContentAlignment = AlignmentConstants.vbLeftJustify
    Set B.Child = CP
    Tmpl.Children.Add B
    Tmpl.SetContentAlignmentMarker AlignmentConstants.vbLeftJustify, 0
    Set St = NewStyle("Button")
    Set St.Template = Tmpl
    Set Btn.Style = St
    Set Live = Btn.Children(0)
    If TypeName(Live.Child) <> "ContentPresenter" Then Err.Raise vbObjectError, , "Nested-only Border.Child expected ContentPresenter"
    Btn.SyncContentPresenter
    If Btn.ContentPresenter.HorizontalContentAlignment <> AlignmentConstants.vbLeftJustify Then
        Err.Raise vbObjectError, , "Nested-only CP HAlign expected Left"
    End If

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Tmpl = Nothing

    LogResult "P6j-MULTI", 0, "OK multi-node Border/Grid/ContentPresenter template"
    Debug.Print "PASS  P6j-MULTI multi-node template tree"
    Phase6jBench_MultiNodeTemplate = True
    Exit Function

Fail:
    LogResult "P6j-MULTI", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6j-MULTI - " & Err.Description
    Phase6jBench_MultiNodeTemplate = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' {TemplateBinding} shorthand: OneWay from TemplatedParent Path (P6k).
Public Function Phase6kBench_TemplateBindingMarkup() As Boolean
    Dim Btn As Button
    Dim Nested As ContentPresenter
    Dim TB As TemplateBinding
    Dim Expr As BindingExpression
    Dim FailMsg As String

    On Error GoTo Fail

    Set Btn = New Button
    Btn.Content = "TB-OK"

    ' Stamp TemplatedParent on a free CP (TextBlock has no TemplatedParent).
    Set Nested = New ContentPresenter
    Set Nested.TemplatedParent = Btn
    If Nested.TemplatedParent Is Nothing Then Err.Raise vbObjectError, , "TemplatedParent stamp failed"
    If Not Nested.TemplatedParent Is Btn Then Err.Raise vbObjectError, , "TemplatedParent must be Button"

    Set TB = New TemplateBinding
    Set Expr = TB.Attach(Nested, "Content", "Content")
    If Expr Is Nothing Then Err.Raise vbObjectError, , "TemplateBinding.Attach failed"

    If CStr(Nested.Content) <> "TB-OK" Then
        Err.Raise vbObjectError, , "TemplateBinding expected Content=TB-OK, got " & CStr(Nested.Content)
    End If

    KeepAlive Btn
    LogResult "P6k-TBMK", 0, "OK TemplateBinding Attach TemplatedParent Path"
    Debug.Print "PASS  P6k-TBMK TemplateBinding markup/Attach"
    If Not Expr Is Nothing Then Expr.Detach
    Phase6kBench_TemplateBindingMarkup = True
    Exit Function

Fail:
    FailMsg = Err.Description
    If Len(FailMsg) = 0 Then FailMsg = "Error " & CStr(Err.Number)
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P6k-TBMK", 0, "FAIL: " & FailMsg
    Debug.Print "FAIL  P6k-TBMK - " & FailMsg
    KeepAlive Btn
    Phase6kBench_TemplateBindingMarkup = False
End Function
' Phase 2a: ThemesManager merges active theme into host ResourceDictionary (WPF swap).
Public Function Phase2aBench_ThemeDictionarySwap() As Boolean
    Dim TM As ThemesManager
    Dim Host As ResourceDictionary
    Dim Light As ObservableDictionary
    Dim Dark As ObservableDictionary
    Dim V As Variant

    On Error GoTo Fail

    Set Light = New ObservableDictionary
    Light.Add "AccentToken", "LightAccent"
    Set Dark = New ObservableDictionary
    Dark.Add "AccentToken", "DarkAccent"

    Set TM = New ThemesManager
    TM.Add "Light", Light
    TM.Add "Dark", Dark

    Set Host = New ResourceDictionary
    TM.AttachToResources Host

    TM.ActiveThemeName = "Light"
    If Not Host.TryGetResource("AccentToken", V) Then Err.Raise vbObjectError, , "Light AccentToken missing after merge"
    If CStr(V) <> "LightAccent" Then Err.Raise vbObjectError, , "Expected LightAccent, got " & CStr(V)

    TM.ActiveThemeName = "Dark"
    If Not Host.TryGetResource("AccentToken", V) Then Err.Raise vbObjectError, , "Dark AccentToken missing after swap"
    If CStr(V) <> "DarkAccent" Then Err.Raise vbObjectError, , "Expected DarkAccent, got " & CStr(V)

    ' Prior theme bag unmerged ? only active theme contributes AccentToken.
    If Host.MergedDictionaries.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 merged theme dict, got " & Host.MergedDictionaries.Count

    LogResult "P2a-THEME-SWAP", 0, "OK ThemesManager ActiveThemeName merges ResourceDictionary"
    Debug.Print "PASS  P2a-THEME-SWAP theme dictionary merge/swap"
    Phase2aBench_ThemeDictionarySwap = True
    Exit Function

Fail:
    LogResult "P2a-THEME-SWAP", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2a-THEME-SWAP - " & Err.Description
    Phase2aBench_ThemeDictionarySwap = False
End Function

' Phase 2a: ActiveThemeName=System resolves to OS Light/Dark (override for gate).
Public Function Phase2aBench_SystemThemeResolve() As Boolean
    Dim TM As ThemesManager
    Dim Host As ResourceDictionary
    Dim Light As ObservableDictionary
    Dim Dark As ObservableDictionary
    Dim V As Variant

    On Error GoTo Fail

    Set Light = New ObservableDictionary
    Light.Add "AccentToken", "LightAccent"
    Set Dark = New ObservableDictionary
    Dark.Add "AccentToken", "DarkAccent"

    Set TM = New ThemesManager
    TM.Add "Light", Light
    TM.Add "Dark", Dark

    Set Host = New ResourceDictionary
    TM.AttachToResources Host

    TM.SystemThemeOverride = "Dark"
    TM.ActiveThemeName = "System"
    If TM.ActiveThemeName <> "System" Then Err.Raise vbObjectError, , "ActiveThemeName expected System"
    If TM.EffectiveThemeName <> "Dark" Then Err.Raise vbObjectError, , "EffectiveThemeName expected Dark, got " & TM.EffectiveThemeName
    If Not Host.TryGetResource("AccentToken", V) Then Err.Raise vbObjectError, , "Dark AccentToken missing"
    If CStr(V) <> "DarkAccent" Then Err.Raise vbObjectError, , "Expected DarkAccent, got " & CStr(V)

    TM.SystemThemeOverride = "Light"
    If TM.EffectiveThemeName <> "Light" Then Err.Raise vbObjectError, , "EffectiveThemeName expected Light after override, got " & TM.EffectiveThemeName
    If Not Host.TryGetResource("AccentToken", V) Then Err.Raise vbObjectError, , "Light AccentToken missing after override"
    If CStr(V) <> "LightAccent" Then Err.Raise vbObjectError, , "Expected LightAccent, got " & CStr(V)

    ' Named theme still works alongside System.
    TM.ActiveThemeName = "Dark"
    If TM.EffectiveThemeName <> "Dark" Then Err.Raise vbObjectError, , "Named Dark EffectiveThemeName expected Dark"
    If Not Host.TryGetResource("AccentToken", V) Then Err.Raise vbObjectError, , "Named Dark AccentToken missing"
    If CStr(V) <> "DarkAccent" Then Err.Raise vbObjectError, , "Named Dark expected DarkAccent"

    LogResult "P2a-THEME-OS", 0, "OK System theme resolves Light/Dark via override"
    Debug.Print "PASS  P2a-THEME-OS System Light/Dark resolve"
    Phase2aBench_SystemThemeResolve = True
    Exit Function

Fail:
    LogResult "P2a-THEME-OS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2a-THEME-OS - " & Err.Description
    Phase2aBench_SystemThemeResolve = False
End Function

Public Function Phase7aBench_PosSalesOrderShell() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Scene
    Dim Grid As Object

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\PosSalesOrderShell.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "POS SalesOrder shell returned Nothing"
    If Root.Name <> "SalesOrder" Then Err.Raise vbObjectError, , "Expected Name=SalesOrder, got " & Root.Name
    If Root.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 child, got " & Root.Children.Count

    Set Grid = Root.Children(0)
    If TypeName(Grid) <> "UniformGrid" Then Err.Raise vbObjectError, , "Expected UniformGrid, got " & TypeName(Grid)

    LogResult "P7a-SMOKE", 0, "OK POS SalesOrder shell Scene+UniformGrid"
    Debug.Print "PASS  P7a-SMOKE POS SalesOrder shell XAML"
    Phase7aBench_PosSalesOrderShell = True
    Exit Function

Fail:
    LogResult "P7a-SMOKE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7a-SMOKE ? " & Err.Description
    Phase7aBench_PosSalesOrderShell = False
End Function

Public Function Phase7cBench_LegacyLayoutShim() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Panel
    Dim Tb As TextBlock
    Dim Marg As Thickness

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\PosMigratedTextBlockLayout.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "POS migrated layout returned Nothing"
    If Root.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 child, got " & Root.Children.Count

    Set Tb = Root.Children(0)
    If TypeName(Tb) <> "TextBlock" Then Err.Raise vbObjectError, , "Expected TextBlock, got " & TypeName(Tb)
    Set Marg = Tb.Margin
    If Marg Is Nothing Then Err.Raise vbObjectError, , "Expected Margin"
    If Marg.Left <> 10# Then Err.Raise vbObjectError, , "Expected Margin.Left=10, got " & Marg.Left
    If Marg.Top <> 20# Then Err.Raise vbObjectError, , "Expected Margin.Top=20, got " & Marg.Top
    If Tb.Width <> 200# Then Err.Raise vbObjectError, , "Expected Width=200, got " & Tb.Width
    If Tb.Height <> 30# Then Err.Raise vbObjectError, , "Expected Height=30, got " & Tb.Height

    LogResult "P7c-LAY", 0, "OK TextBlock Margin/Width/Height (no Design*)"
    Debug.Print "PASS  P7c-LAY TextBlock Margin/Width/Height (WPF layout)"
    Phase7cBench_LegacyLayoutShim = True
    Exit Function

Fail:
    LogResult "P7c-LAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7c-LAY - " & Err.Description
    Phase7cBench_LegacyLayoutShim = False
End Function

Public Function Phase7dBench_PanelResize() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Grid
    Dim LeftCol As Border
    Dim RightCol As Border
    Dim Lines As ListView

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\LayoutPanelResize.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "LayoutPanelResize returned Nothing"
    If Root.Children.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 children, got " & Root.Children.Count

    Set LeftCol = Root.Children(0)
    Set RightCol = Root.Children(1)
    If TypeName(LeftCol) <> "Border" Then Err.Raise vbObjectError, , "Expected LeftColumn Border, got " & TypeName(LeftCol)
    If LeftCol.Children.Count < 1 Then Err.Raise vbObjectError, , "Expected LinesList under LeftColumn"
    Set Lines = LeftCol.Children(0)

    ' Full design size: star columns split host (minus margins 4+4).
    Root.Widget.Move 0, 0, 400, 300
    If Abs(LeftCol.Widget.Left - 0!) > 3! Then Err.Raise vbObjectError, , "Left.Left expected ~0, got " & LeftCol.Widget.Left
    If Abs(LeftCol.Widget.Width - 196!) > 6! Then Err.Raise vbObjectError, , "Left.Width expected ~196, got " & LeftCol.Widget.Width
    If Abs(RightCol.Widget.Left - 204!) > 6! Then Err.Raise vbObjectError, , "Right.Left expected ~204, got " & RightCol.Widget.Left
    If Abs(RightCol.Widget.Width - 196!) > 6! Then Err.Raise vbObjectError, , "Right.Width expected ~196, got " & RightCol.Widget.Width
    If Abs(LeftCol.Widget.Height - 300!) > 6! Then Err.Raise vbObjectError, , "Left.Height expected ~300, got " & LeftCol.Widget.Height
    If Lines.Widget.Height < 280! Then Err.Raise vbObjectError, , "LinesList should track column height, got " & Lines.Widget.Height

    ' Half host: columns still fill halves via Grid stars (not 0.5x absolute Margin math).
    Root.Widget.Move 0, 0, 200, 150
    If Abs(LeftCol.Widget.Width - 96!) > 6! Then Err.Raise vbObjectError, , "Left.Width expected ~96 after half, got " & LeftCol.Widget.Width
    If Abs(RightCol.Widget.Left - 104!) > 6! Then Err.Raise vbObjectError, , "Right.Left expected ~104 after half, got " & RightCol.Widget.Left
    If Abs(RightCol.Widget.Width - 96!) > 6! Then Err.Raise vbObjectError, , "Right.Width expected ~96 after half, got " & RightCol.Widget.Width
    If Abs(LeftCol.Widget.Height - 150!) > 6! Then Err.Raise vbObjectError, , "Left.Height expected ~150 after half, got " & LeftCol.Widget.Height
    If Lines.Widget.Height < 130! Then Err.Raise vbObjectError, , "LinesList should track half height, got " & Lines.Widget.Height

    KeepAlive Root
    LogResult "P7d-LAY-PANEL", 0, "OK Grid star columns + ListView fill on resize"
    Debug.Print "PASS  P7d-LAY-PANEL Grid star columns fill on host resize"
    Phase7dBench_PanelResize = True
    Exit Function

Fail:
    LogResult "P7d-LAY-PANEL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7d-LAY-PANEL - " & Err.Description
    Phase7dBench_PanelResize = False
End Function

Public Function Phase8Bench_InheritanceBatch() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Border
    Dim Mid As Border
    Dim Inner As Border
    Dim Tb As TextBlock
    Dim PassDuringLoad As Long
    Dim Vm As Phase0ViewModel
    Dim Vm2 As Phase0ViewModel
    Dim Expr As BindingExpression
    Dim DataContextProp As DependencyProperty

    On Error GoTo Fail

    VCF.ResetInheritanceCounters

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\InheritanceNestedBorder.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "InheritanceNestedBorder returned Nothing"
    If Root.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 child on root"

    PassDuringLoad = VCF.PassPropertyValueCalls
    ' Phase 8b: PassPropertyValue is a no-op (lazy GetValue); keep a soft ceiling.
    If PassDuringLoad > 0 Then Err.Raise vbObjectError, , "PassPropertyValueCalls during load expected 0, got " & PassDuringLoad

    Set Mid = Root.Children(0)
    Set Inner = Mid.Children(0)
    Set Tb = Inner.Children(0)

    Set Vm = New Phase0ViewModel
    Vm.Title = "InheritedCtx"
    Set Root.DataContext = Vm

    If Not Mid.DataContext Is Vm Then Err.Raise vbObjectError, , "Mid DataContext not inherited"
    If Not Inner.DataContext Is Vm Then Err.Raise vbObjectError, , "Inner DataContext not inherited"

    Set DataContextProp = Tb.DependencyProperties.GetProperty("DataContext")
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", DataContextProp, "Title", OneWay
    If Tb.Text <> "InheritedCtx" Then Err.Raise vbObjectError, , "Expected InheritedCtx from pull DataContext, got " & Tb.Text

    Set Vm2 = New Phase0ViewModel
    Vm2.Title = "RebindCtx"
    Set Root.DataContext = Vm2
    If Tb.Text <> "RebindCtx" Then Err.Raise vbObjectError, , "Expected RebindCtx after ancestor DataContext change, got " & Tb.Text

    LogResult "P8-INHERIT", 0, "OK PassDuringLoad=" & PassDuringLoad & " InheritCalls=" & VCF.InheritPropertyValuesCalls & " lazy"
    Debug.Print "PASS  P8-INHERIT lazy GetValue inherit + DataContext"
    Expr.Detach
    Set Expr = Nothing
    Phase8Bench_InheritanceBatch = True
    Exit Function

Fail:
    On Error Resume Next
    If Not Expr Is Nothing Then Expr.Detach
    LogResult "P8-INHERIT", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P8-INHERIT ? " & Err.Description
    Phase8Bench_InheritanceBatch = False
End Function

Public Function Phase2aBench_NestedUniformGridResize() As Boolean
    Dim Reader As XAMLReader
    Dim Outer As UniformGrid
    Dim Inner As UniformGrid
    Dim i As Long
    Dim Started As Single
    Dim ElapsedMs As Long

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Outer = Reader.Load(LoadTextFile(App.Path & "\Resources\NestedUniformGridResize.xml"))
    If Outer Is Nothing Then Err.Raise vbObjectError, , "NestedUniformGridResize returned Nothing"
    If Outer.Children.Count <> 4 Then Err.Raise vbObjectError, , "Expected 4 nested grids, got " & Outer.Children.Count

    Set Inner = Outer.Children(0)

    Outer.Widget.Move 0, 0, 400, 300
    If Abs(Inner.Widget.Width - 200!) > 3! Then Err.Raise vbObjectError, , "Full Inner.Width expected ~200, got " & Inner.Widget.Width

    Started = Timer
    For i = 1 To 50
        If (i Mod 2) = 0 Then
            Outer.Widget.Move 0, 0, 400, 300
        Else
            Outer.Widget.Move 0, 0, 200, 150
        End If
    Next
    ElapsedMs = CLng((Timer - Started) * 1000#)

    ' Loop ends on odd i=49 ? half size; i=50 even ? full. Assert full cell size.
    If Abs(Inner.Widget.Width - 200!) > 3! Then Err.Raise vbObjectError, , "After 50? resize Inner.Width expected ~200, got " & Inner.Widget.Width
    If Abs(Inner.Widget.Height - 150!) > 3! Then Err.Raise vbObjectError, , "After 50? resize Inner.Height expected ~150, got " & Inner.Widget.Height

    LogResult "B-RESZ", ElapsedMs, "OK nested UniformGrid 50x resize"
    Debug.Print "PASS  B-RESZ nested UniformGrid resize x50 (" & ElapsedMs & " ms)"
    Phase2aBench_NestedUniformGridResize = True
    Exit Function

Fail:
    LogResult "B-RESZ", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-RESZ ? " & Err.Description
    Phase2aBench_NestedUniformGridResize = False
End Function

Public Function Phase2aBench_ViewNavLeak() As Boolean
    Dim AppHost As Phase0App
    Dim Shell As Phase0Shell
    Dim Win As Window
    Dim ViewA As Panel
    Dim ViewB As Panel
    Dim TbA As TextBlock
    Dim TbB As TextBlock
    Dim VmA As Phase0ViewModel
    Dim VmB As Phase0ViewModel
    Dim ExprA As BindingExpression
    Dim ExprB As BindingExpression
    Dim i As Long
    Dim Started As Single
    Dim ElapsedMs As Long
    Dim WinCount As Long

    On Error GoTo Fail

    VCF.ClearApplication

    Set AppHost = New Phase0App
    Set Shell = New Phase0Shell
    Set Win = Shell.Base
    If Win Is Nothing Then Err.Raise vbObjectError, , "Phase0Shell.Base is Nothing"

    WinCount = AppHost.Base.Windows.Count
    If WinCount < 1 Then Err.Raise vbObjectError, , "Expected Windows.Count >= 1, got " & WinCount

    Set VmA = New Phase0ViewModel
    Set VmB = New Phase0ViewModel
    VmA.Title = "ViewA"
    VmB.Title = "ViewB"

    ' UserControl is VB_Creatable=False; Panel is the creatable host for Phase0 nav benches.
    Set ViewA = New Panel
    Set ViewB = New Panel
    Set TbA = New TextBlock
    Set TbB = New TextBlock
    ViewA.Children.Add TbA
    ViewB.Children.Add TbB

    Set ExprA = New BindingExpression
    Set ExprB = New BindingExpression
    ExprA.Attach TbA, "Text", VmA, "Title", OneWay
    ExprB.Attach TbB, "Text", VmB, "Title", OneWay

    If TbA.Text <> "ViewA" Then Err.Raise vbObjectError, , "Expected ViewA bind, got " & TbA.Text
    If TbB.Text <> "ViewB" Then Err.Raise vbObjectError, , "Expected ViewB bind, got " & TbB.Text

    Win.Children.Add ViewA
    Win.Children.Add ViewB
    ViewA.Visibility = VisibilityVisible
    ViewB.Visibility = VisibilityCollapsed
    Win.RelayoutChildren
    Win.RebuildNamedItemsList

    Started = Timer
    For i = 1 To 50
        If (i Mod 2) = 0 Then
            ViewA.Visibility = VisibilityVisible
            ViewB.Visibility = VisibilityCollapsed
        Else
            ViewB.Visibility = VisibilityVisible
            ViewA.Visibility = VisibilityCollapsed
        End If
        Win.RelayoutChildren
        Win.RebuildNamedItemsList
    Next
    ElapsedMs = CLng((Timer - Started) * 1000#)

    ' Active is ViewA (i=50 even). Detach trees and prove INPC no longer updates targets.
    VCF.DetachBindingsTree ViewA
    VCF.DetachBindingsTree ViewB
    Set ExprA = Nothing
    Set ExprB = Nothing

    VmA.Title = "LeakedA"
    VmB.Title = "LeakedB"
    If TbA.Text = "LeakedA" Then Err.Raise vbObjectError, , "ViewA binding leaked after DetachBindingsTree"
    If TbB.Text = "LeakedB" Then Err.Raise vbObjectError, , "ViewB binding leaked after DetachBindingsTree"

    Win.Unload
    On Error Resume Next
    Cairo.WidgetForms.RemoveAll
    On Error GoTo Fail
    Err.Clear

    WinCount = 0
    If Not AppHost.Base Is Nothing Then WinCount = AppHost.Base.Windows.Count
    If WinCount <> 0 Then Err.Raise vbObjectError, , "Expected Windows.Count=0 after dispose, got " & WinCount

    Set Shell = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication

    LogResult "B-NAV", ElapsedMs, "OK 50x Visibility nav; Windows=0; no bind leak"
    Debug.Print "PASS  B-NAV view nav x50 + Windows registry (" & ElapsedMs & " ms)"
    Phase2aBench_ViewNavLeak = True
    Exit Function

Fail:
    Dim FailDesc As String
    FailDesc = Err.Description
    On Error Resume Next
    If Not ExprA Is Nothing Then ExprA.Detach
    If Not ExprB Is Nothing Then ExprB.Detach
    If Not Win Is Nothing Then Win.Unload
    Cairo.WidgetForms.RemoveAll
    Set Shell = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication
    LogResult "B-NAV", 0, "FAIL: " & FailDesc
    Debug.Print "FAIL  B-NAV ? " & FailDesc
    Phase2aBench_ViewNavLeak = False
End Function

' First-frame chrome: after NewWindow (before Show), Form.BorderStyle matches local DP
' (Create(0) + SyncFormBorderStyle) - borderless shell and bordered dialog.
Public Function Phase2aBench_WindowChrome() As Boolean
    Dim AppHost As Phase0App
    Dim Borderless As Phase0Shell
    Dim Bordered As Phase0BorderedShell
    Dim Win As Window

    On Error GoTo Fail

    VCF.ClearApplication
    Set AppHost = New Phase0App

    Set Borderless = New Phase0Shell
    Set Win = Borderless.Base
    If Win Is Nothing Then Err.Raise vbObjectError, , "Phase0Shell.Base is Nothing"
    If CLng(Win.DependencyProperties.GetValue("BorderStyle")) <> 0 Then
        Err.Raise vbObjectError, , "Borderless DP BorderStyle expected 0, got " & Win.DependencyProperties.GetValue("BorderStyle")
    End If
    If Win.Form Is Nothing Then Err.Raise vbObjectError, , "Borderless Form is Nothing"
    If Win.Form.BorderStyle <> 0 Then
        Err.Raise vbObjectError, , "Borderless Form.BorderStyle expected 0 after NewWindow, got " & Win.Form.BorderStyle
    End If

    Win.Unload
    Set Borderless = Nothing
    Set Win = Nothing
    On Error Resume Next
    Cairo.WidgetForms.RemoveAll
    On Error GoTo Fail
    Err.Clear

    Set Bordered = New Phase0BorderedShell
    Set Win = Bordered.Base
    If Win Is Nothing Then Err.Raise vbObjectError, , "Phase0BorderedShell.Base is Nothing"
    If CLng(Win.DependencyProperties.GetValue("BorderStyle")) <> 2 Then
        Err.Raise vbObjectError, , "Bordered DP BorderStyle expected 2, got " & Win.DependencyProperties.GetValue("BorderStyle")
    End If
    If Win.Form Is Nothing Then Err.Raise vbObjectError, , "Bordered Form is Nothing"
    If Win.Form.BorderStyle <> 2 Then
        Err.Raise vbObjectError, , "Bordered Form.BorderStyle expected 2 after NewWindow, got " & Win.Form.BorderStyle
    End If

    Win.Unload
    Set Bordered = Nothing
    Set AppHost = Nothing
    On Error Resume Next
    Cairo.WidgetForms.RemoveAll
    On Error GoTo Fail
    VCF.ClearApplication

    LogResult "B-CHROME", 0, "OK borderless=0 + bordered=2 after NewWindow"
    Debug.Print "PASS  B-CHROME first-frame BorderStyle sync (0 + 2)"
    Phase2aBench_WindowChrome = True
    Exit Function

Fail:
    Dim FailDesc As String
    FailDesc = Err.Description
    On Error Resume Next
    If Not Win Is Nothing Then Win.Unload
    Cairo.WidgetForms.RemoveAll
    Set Borderless = Nothing
    Set Bordered = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication
    LogResult "B-CHROME", 0, "FAIL: " & FailDesc
    Debug.Print "FAIL  B-CHROME - " & FailDesc
    Phase2aBench_WindowChrome = False
End Function

' ListView bind hotspot (framework-first): menu-like density = 21 rows ? 6 DataContext bindings/cell.
' Gates CloneDataTemplateForItem binding fidelity (ItemsControl generation is covered by P4b-ICtrl).
Public Function Phase2aBench_ListViewBindHotspot() As Boolean
    Const ROW_COUNT As Long = 21
    Const BIND_PER_CELL As Long = 6

    Dim Coll As ObservableCollection
    Dim Tmpl As DataTemplate
    Dim Tb As TextBlock
    Dim Vm As Phase0ViewModel
    Dim RowItem As Object
    Dim i As Long
    Dim j As Long
    Dim Started As Single
    Dim ElapsedMs As Long
    Dim Cloned As DataTemplate
    Dim FirstClone As DataTemplate
    Dim ChildTb As TextBlock
    Dim FirstCells(0 To BIND_PER_CELL - 1) As TextBlock

    On Error GoTo Fail

    Set Coll = New ObservableCollection
    For i = 1 To ROW_COUNT
        Set Vm = New Phase0ViewModel
        Vm.Title = "R" & CStr(i)
        Coll.Add Vm
    Next

    Set Tmpl = New DataTemplate
    For j = 1 To BIND_PER_CELL
        Set Tb = New TextBlock
        Tb.Text = "?"
        AttachDataContextTitleBinding Tb
        Tmpl.Children.Add Tb
    Next

    Started = Timer
    For i = 0 To ROW_COUNT - 1
        Set RowItem = Coll(i)
        Set Cloned = VCF.CloneDataTemplateForItem(Tmpl, RowItem, Nothing)
        If Cloned Is Nothing Then Err.Raise vbObjectError, , "CloneDataTemplateForItem returned Nothing at " & i
        If Cloned.Children.Count <> BIND_PER_CELL Then Err.Raise vbObjectError, , "Expected " & BIND_PER_CELL & " children, got " & Cloned.Children.Count
        If i = 0 Then
            Set FirstClone = Cloned
            For j = 0 To BIND_PER_CELL - 1
                Set FirstCells(j) = FirstClone.Children(j)
            Next
        Else
            For j = 0 To Cloned.Children.Count - 1
                VCF.DetachBindingsTree Cloned.Children(j)
            Next
            Set Cloned = Nothing
        End If
    Next
    ElapsedMs = CLng((Timer - Started) * 1000#)

    If FirstCells(0).Text <> "R1" Then Err.Raise vbObjectError, , "Expected R1 after clone+DataContext, got [" & FirstCells(0).Text & "]"
    If FirstCells(BIND_PER_CELL - 1).Text <> "R1" Then Err.Raise vbObjectError, , "Expected R1 on last cell binding, got [" & FirstCells(BIND_PER_CELL - 1).Text & "]"

    Set Vm = Coll(0)
    Vm.Title = "Mutated"
    If FirstCells(0).Text <> "Mutated" Then Err.Raise vbObjectError, , "Expected Mutated after INPC, got [" & FirstCells(0).Text & "]"
    If FirstCells(BIND_PER_CELL - 1).Text <> "Mutated" Then Err.Raise vbObjectError, , "Expected Mutated on last cell after INPC, got [" & FirstCells(BIND_PER_CELL - 1).Text & "]"

    For j = 0 To BIND_PER_CELL - 1
        VCF.DetachBindingsTree FirstCells(j)
    Next

    Vm.Title = "AfterDetach"
    If FirstCells(0).Text = "AfterDetach" Then Err.Raise vbObjectError, , "Binding leaked after DetachBindingsTree"

    CleanupBindDenseArtifacts Tmpl, FirstClone, FirstCells, Coll

    LogResult "B-BIND-DENSE", ElapsedMs, "OK 21x6 template clone+bind+INPC+detach"
    Debug.Print "PASS  B-BIND-DENSE ListView template bind hotspot (" & ElapsedMs & " ms)"
    Phase2aBench_ListViewBindHotspot = True
    Exit Function

Fail:
    Dim FailDesc As String
    Dim FailNum As Long
    FailNum = Err.Number
    FailDesc = Err.Description
    On Error Resume Next
    CleanupBindDenseArtifacts Tmpl, FirstClone, FirstCells, Coll
    LogResult "B-BIND-DENSE", 0, "FAIL: " & CStr(FailNum) & " " & FailDesc
    Debug.Print "FAIL  B-BIND-DENSE ? (" & FailNum & ") " & FailDesc
    Phase2aBench_ListViewBindHotspot = False
End Function

Private Sub CleanupBindDenseArtifacts( _
    ByRef Tmpl As DataTemplate, _
    ByRef FirstClone As DataTemplate, _
    ByRef FirstCells() As TextBlock, _
    ByRef Coll As ObservableCollection)

    Dim j As Long
    Dim Child As Object
    Dim El As IUIElement

    On Error Resume Next

    For j = LBound(FirstCells) To UBound(FirstCells)
        If Not FirstCells(j) Is Nothing Then
            VCF.DetachBindingsTree FirstCells(j)
            Set FirstCells(j).DataContext = Nothing
            Set FirstCells(j) = Nothing
        End If
    Next

    If Not FirstClone Is Nothing Then
        For j = 0 To FirstClone.Children.Count - 1
            Set Child = FirstClone.Children(j)
            VCF.DetachBindingsTree Child
            If TypeOf Child Is IUIElement Then
                Set El = Child
                Set El.DataContext = Nothing
            End If
        Next
        Set FirstClone = Nothing
    End If

    If Not Tmpl Is Nothing Then
        For j = 0 To Tmpl.Children.Count - 1
            Set Child = Tmpl.Children(j)
            VCF.DetachBindingsTree Child
            If TypeOf Child Is IUIElement Then
                Set El = Child
                Set El.DataContext = Nothing
            End If
        Next
        Set Tmpl = Nothing
    End If

    If Not Coll Is Nothing Then
        Coll.Clear
        Set Coll = Nothing
    End If

    Cairo.WidgetForms.RemoveAll
End Sub

' Phase 2a Margin/Padding family 1: ListView DPs (Margin=0, Padding=4,1,4,1 Win10 ListBoxItem).
Public Function Phase2aBench_ListViewPaddingDefaults() As Boolean
    Dim LV As ListView
    Dim Marg As Thickness
    Dim Pad As Thickness
    Dim Custom As Thickness

    On Error GoTo Fail

    Set LV = New ListView
    If Not LV.DependencyProperties.Exists("Margin") Then Err.Raise vbObjectError, , "ListView missing Margin DP"
    If Not LV.DependencyProperties.Exists("Padding") Then Err.Raise vbObjectError, , "ListView missing Padding DP"

    Set Marg = LV.Margin
    If Marg Is Nothing Then Err.Raise vbObjectError, , "ListView.Margin is Nothing"
    If Marg.Left <> 0 Or Marg.Top <> 0 Or Marg.Right <> 0 Or Marg.Bottom <> 0 Then
        Err.Raise vbObjectError, , "ListView.Margin default expected 0,0,0,0"
    End If

    Set Pad = LV.Padding
    If Pad Is Nothing Then Err.Raise vbObjectError, , "ListView.Padding is Nothing"
    If Pad.Left <> 4 Or Pad.Top <> 1 Or Pad.Right <> 4 Or Pad.Bottom <> 1 Then
        Err.Raise vbObjectError, , "ListView.Padding default expected 4,1,4,1 got " & _
            Pad.Left & "," & Pad.Top & "," & Pad.Right & "," & Pad.Bottom
    End If

    Set Custom = New Thickness
    Custom.Left = 8
    Custom.Top = 2
    Custom.Right = 8
    Custom.Bottom = 2
    Set LV.Padding = Custom
    Set Pad = LV.Padding
    If Pad.Left <> 8 Or Pad.Top <> 2 Or Pad.Right <> 8 Or Pad.Bottom <> 2 Then
        Err.Raise vbObjectError, , "ListView.Padding set expected 8,2,8,2"
    End If

    Set Custom = New Thickness
    Custom.Left = 4
    Custom.Top = 4
    Custom.Right = 4
    Custom.Bottom = 4
    Set LV.Margin = Custom
    Set Marg = LV.Margin
    If Marg.Left <> 4 Or Marg.Top <> 4 Or Marg.Right <> 4 Or Marg.Bottom <> 4 Then
        Err.Raise vbObjectError, , "ListView.Margin set expected 4,4,4,4"
    End If

    LogResult "P2a-PAD", 0, "OK ListView Margin=0 Padding=4,1,4,1 + setters"
    Debug.Print "PASS  P2a-PAD ListView Margin/Padding defaults"
    Phase2aBench_ListViewPaddingDefaults = True
    Exit Function

Fail:
    LogResult "P2a-PAD", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2a-PAD ? " & Err.Description
    Phase2aBench_ListViewPaddingDefaults = False
End Function

' Phase 2a Margin/Padding family 2: TextBox (Margin=0, Padding=1) + Button (Padding=1 Aero2).
' Qualify VCF.TextBox ? bare TextBox resolves to VB.TextBox (higher in project refs).
Public Function Phase2aBench_TextBoxButtonPaddingDefaults() As Boolean
    Dim Tb As VCF.TextBox
    Dim Btn As Button
    Dim Marg As Thickness
    Dim Pad As Thickness
    Dim Custom As Thickness

    On Error GoTo Fail

    Set Tb = New VCF.TextBox
    If Not Tb.DependencyProperties.Exists("Margin") Then Err.Raise vbObjectError, , "TextBox missing Margin DP"
    If Not Tb.DependencyProperties.Exists("Padding") Then Err.Raise vbObjectError, , "TextBox missing Padding DP"

    Set Marg = Tb.Margin
    If Marg Is Nothing Then Err.Raise vbObjectError, , "TextBox.Margin is Nothing"
    If Marg.Left <> 0 Or Marg.Top <> 0 Or Marg.Right <> 0 Or Marg.Bottom <> 0 Then
        Err.Raise vbObjectError, , "TextBox.Margin default expected 0,0,0,0"
    End If

    Set Pad = Tb.Padding
    If Pad Is Nothing Then Err.Raise vbObjectError, , "TextBox.Padding is Nothing"
    If Pad.Left <> 1 Or Pad.Top <> 1 Or Pad.Right <> 1 Or Pad.Bottom <> 1 Then
        Err.Raise vbObjectError, , "TextBox.Padding default expected 1,1,1,1 got " & _
            Pad.Left & "," & Pad.Top & "," & Pad.Right & "," & Pad.Bottom
    End If
    If Tb.InnerSpace <> 1 Then Err.Raise vbObjectError, , "TextBox.InnerSpace expected 1 after default Padding"

    Set Custom = New Thickness
    Custom.Left = 4
    Custom.Top = 2
    Custom.Right = 6
    Custom.Bottom = 3
    Set Tb.Padding = Custom
    Set Pad = Tb.Padding
    If Pad.Left <> 4 Or Pad.Top <> 2 Or Pad.Right <> 6 Or Pad.Bottom <> 3 Then
        Err.Raise vbObjectError, , "TextBox.Padding set expected 4,2,6,3"
    End If
    If Tb.InnerSpace <> 4 Then Err.Raise vbObjectError, , "TextBox.InnerSpace expected 4 (Left) after asymmetric Padding"

    Set Btn = New Button
    If Not Btn.DependencyProperties.Exists("Padding") Then Err.Raise vbObjectError, , "Button missing Padding DP"
    Set Pad = Btn.Padding
    If Pad Is Nothing Then Err.Raise vbObjectError, , "Button.Padding is Nothing"
    If Pad.Left <> 1 Or Pad.Top <> 1 Or Pad.Right <> 1 Or Pad.Bottom <> 1 Then
        Err.Raise vbObjectError, , "Button.Padding default expected 1,1,1,1 got " & _
            Pad.Left & "," & Pad.Top & "," & Pad.Right & "," & Pad.Bottom
    End If

    Set Custom = New Thickness
    Custom.Left = 8
    Custom.Top = 2
    Custom.Right = 8
    Custom.Bottom = 2
    Set Btn.Padding = Custom
    Set Pad = Btn.Padding
    If Pad.Left <> 8 Or Pad.Top <> 2 Or Pad.Right <> 8 Or Pad.Bottom <> 2 Then
        Err.Raise vbObjectError, , "Button.Padding set expected 8,2,8,2"
    End If

    LogResult "P2a-PAD-TB", 0, "OK TextBox Margin=0 Padding=1 + Button Padding=1 Aero2"
    Debug.Print "PASS  P2a-PAD-TB TextBox/Button Margin/Padding defaults"
    Phase2aBench_TextBoxButtonPaddingDefaults = True
    Exit Function

Fail:
    LogResult "P2a-PAD-TB", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2a-PAD-TB ? " & Err.Description
    Phase2aBench_TextBoxButtonPaddingDefaults = False
End Function

' Phase 2a Margin/Padding family 3: UniformGrid cell Padding default stays 2 (intentional; not WPF Panel).
Public Function Phase2aBench_UniformGridPaddingDefault() As Boolean
    Dim Ug As UniformGrid
    Dim Pad As Thickness
    Dim Custom As Thickness

    On Error GoTo Fail

    Set Ug = New UniformGrid
    If Not Ug.DependencyProperties.Exists("Padding") Then Err.Raise vbObjectError, , "UniformGrid missing Padding DP"

    Set Pad = Ug.Padding
    If Pad Is Nothing Then Err.Raise vbObjectError, , "UniformGrid.Padding is Nothing"
    If Pad.Left <> 2 Or Pad.Top <> 2 Or Pad.Right <> 2 Or Pad.Bottom <> 2 Then
        Err.Raise vbObjectError, , "UniformGrid.Padding default expected 2,2,2,2 got " & _
            Pad.Left & "," & Pad.Top & "," & Pad.Right & "," & Pad.Bottom
    End If

    Set Custom = New Thickness
    Custom.Left = 4
    Custom.Top = 4
    Custom.Right = 4
    Custom.Bottom = 4
    Set Ug.Padding = Custom
    Set Pad = Ug.Padding
    If Pad.Left <> 4 Or Pad.Top <> 4 Or Pad.Right <> 4 Or Pad.Bottom <> 4 Then
        Err.Raise vbObjectError, , "UniformGrid.Padding set expected 4,4,4,4"
    End If

    LogResult "P2a-PAD-UG", 0, "OK UniformGrid Padding default=2 + setter"
    Debug.Print "PASS  P2a-PAD-UG UniformGrid Padding default"
    Phase2aBench_UniformGridPaddingDefault = True
    Exit Function

Fail:
    LogResult "P2a-PAD-UG", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2a-PAD-UG ? " & Err.Description
    Phase2aBench_UniformGridPaddingDefault = False
End Function

Public Function Phase2aBench_UniformGridMeasure() As Boolean
    Dim Ug As UniformGrid
    Dim P1 As Panel
    Dim P2 As Panel

    On Error GoTo Fail

    Set Ug = New UniformGrid
    Ug.Rows = 2
    Ug.Columns = 2
    Ug.Width = 0
    Ug.Height = 0
    Ug.Widget.Move 0, 0, 200, 200

    Set P1 = New Panel
    P1.Width = 40
    P1.Height = 40
    Set P2 = New Panel
    P2.Width = 40
    P2.Height = 50
    Ug.Children.Add P1
    Ug.Children.Add P2

    Ug.MeasureLayout 200, 200
    ' Max cell 40x50 * 2x2 => Desired 80x100
    If Abs(Ug.DesiredWidth - 80#) > 0.5 Then Err.Raise vbObjectError, , "DesiredWidth expected 80, got " & Ug.DesiredWidth
    If Abs(Ug.DesiredHeight - 100#) > 0.5 Then Err.Raise vbObjectError, , "DesiredHeight expected 100, got " & Ug.DesiredHeight

    Ug.RelayoutChildren
    If Abs(Ug.ActualWidth - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualWidth expected 200, got " & Ug.ActualWidth
    If Abs(Ug.ActualHeight - 200#) > 0.5 Then Err.Raise vbObjectError, , "ActualHeight expected 200, got " & Ug.ActualHeight

    KeepAlive Ug
    LogResult "P2a-UG-MEAS", 0, "OK Desired=" & Ug.DesiredWidth & "x" & Ug.DesiredHeight
    Debug.Print "PASS  P2a-UG-MEAS UniformGrid Measure/Actual"
    Phase2aBench_UniformGridMeasure = True
    Exit Function

Fail:
    LogResult "P2a-UG-MEAS", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P2a-UG-MEAS - " & Err.Description
    Phase2aBench_UniformGridMeasure = False
    On Error Resume Next
    KeepAlive Ug
    Err.Clear
End Function
' Phase 7c-dialog: ItemsControl + DataTemplate root=Button (binding-only; no @ substitution).
Public Function Phase7cBench_DialogDataTemplate() As Boolean
    Dim IC As ItemsControl
    Dim Coll As ObservableCollection
    Dim Tmpl As DataTemplate
    Dim BtnTmpl As Button
    Dim OkItem As Phase0DialogButtonItem
    Dim CancelItem As Phase0DialogButtonItem
    Dim SharedCmd As Phase0DialogCommand
    Dim Btn0 As Button
    Dim Btn1 As Button
    Dim Cmd As ICommand

    On Error GoTo Fail

    Set SharedCmd = New Phase0DialogCommand

    Set OkItem = New Phase0DialogButtonItem
    OkItem.Text = "OK"
    OkItem.Value = "ok"
    Set OkItem.Command = SharedCmd

    Set CancelItem = New Phase0DialogButtonItem
    CancelItem.Text = "Cancel"
    CancelItem.Value = "cancel"
    Set CancelItem.Command = SharedCmd

    Set Coll = New ObservableCollection
    Coll.Add OkItem
    Coll.Add CancelItem

    Set BtnTmpl = New Button
    BtnTmpl.Width = 80
    BtnTmpl.Height = 28
    AttachDialogButtonTemplateBindings BtnTmpl

    Set Tmpl = New DataTemplate
    Tmpl.Children.Add BtnTmpl

    Set IC = New ItemsControl
    IC.ItemsHost.Orientation = OrientationHorizontal
    Set IC.ItemTemplate = Tmpl
    Set IC.ItemsSource = Coll

    If IC.ItemCount <> 2 Then Err.Raise vbObjectError, , "Expected ItemCount=2"
    If IC.ItemsHost.Children.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 host children"

    Set Btn0 = IC.ItemsHost.Children(0)
    Set Btn1 = IC.ItemsHost.Children(1)
    If Not TypeOf Btn0 Is Button Then Err.Raise vbObjectError, , "Child 0 is not Button"
    If Not TypeOf Btn1 Is Button Then Err.Raise vbObjectError, , "Child 1 is not Button"

    If CStr(Btn0.Content) <> "OK" Then Err.Raise vbObjectError, , "Btn0 Content expected OK got " & CStr(Btn0.Content)
    If CStr(Btn1.Content) <> "Cancel" Then Err.Raise vbObjectError, , "Btn1 Content expected Cancel got " & CStr(Btn1.Content)

    OkItem.Text = "Accept"
    If CStr(Btn0.Content) <> "Accept" Then Err.Raise vbObjectError, , "Btn0 Content expected Accept after INPC"

    Set Cmd = Btn0.Command
    If Cmd Is Nothing Then Err.Raise vbObjectError, , "Btn0.Command not bound"
    If CStr(Btn0.CommandParameter) <> "ok" Then Err.Raise vbObjectError, , "Btn0.CommandParameter expected ok"

    SharedCmd.Reset
    Cmd.Execute Btn0.CommandParameter
    If SharedCmd.ExecuteCount <> 1 Then Err.Raise vbObjectError, , "Command ExecuteCount expected 1"
    If CStr(SharedCmd.LastParameter) <> "ok" Then Err.Raise vbObjectError, , "Command LastParameter expected ok"

    LogResult "P7c-DLG", 0, "OK ItemsControl Button DataTemplate Content+Command"
    Debug.Print "PASS  P7c-DLG dialog DataTemplate (no @)"
    Phase7cBench_DialogDataTemplate = True

    ' Keepalive ? releasing Button ItemsHost mid-suite disconnects widgets (RPC_E_DISCONNECTED).
    KeepAlive IC
    Set Btn0 = Nothing
    Set Btn1 = Nothing
    Set Cmd = Nothing
    Set IC = Nothing
    Set Coll = Nothing
    Set Tmpl = Nothing
    Set BtnTmpl = Nothing
    Set OkItem = Nothing
    Set CancelItem = Nothing
    Set SharedCmd = Nothing
    Debug.Print "P7c-DLG keepalive OK"
    Exit Function

Fail:
    LogResult "P7c-DLG", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7c-DLG ? " & Err.Description
    Phase7cBench_DialogDataTemplate = False
    On Error Resume Next
    KeepAlive IC
    Err.Clear
End Function

' ItemsPanelTemplate gate ? UniformGrid shell + ItemsSource inflate (TextBlock then Button).
Public Function Phase7cBench_ItemsPanelUniformGrid() As Boolean
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
    Debug.Print "P7c-PANEL enter"

    ' --- C0: code ItemsPanel = UniformGrid shell (no ItemsSource) ---
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
    Debug.Print "P7c-PANEL C0 code UniformGrid host OK"
    KeepAlive IC
    Set IC = Nothing
    Set UgHost = Nothing
    Set UgProto = Nothing
    Set PanelTmpl = Nothing
    Set ZeroPad = Nothing

    ' --- C: XAML ItemsPanel UniformGrid shell (no ItemsSource) ---
    Set Reader = New XAMLReader
    Set Root = Reader.Load( _
        "<ItemsControl Width=""200"" Height=""40"">" & _
        "<ItemsControl.ItemsPanel><ItemsPanelTemplate>" & _
        "<UniformGrid Rows=""1"" Columns=""3"" Padding=""0""/>" & _
        "</ItemsPanelTemplate></ItemsControl.ItemsPanel>" & _
        "</ItemsControl>")
    If Root Is Nothing Then Err.Raise vbObjectError, , "XAML ItemsControl returned Nothing"
    Root.Widget.Move 0, 0, 200, 40
    If Root.ItemsPanel Is Nothing Then Err.Raise vbObjectError, , "XAML ItemsPanel is Nothing"
    If Not TypeOf Root.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "XAML ItemsHost expected UniformGrid"
    Set UgXaml = Root.ItemsHost
    If UgXaml.Rows <> 1 Or UgXaml.Columns <> 3 Then Err.Raise vbObjectError, , "XAML host expected 1x3"
    Debug.Print "P7c-PANEL C XAML UniformGrid host OK"
    KeepAlive Root
    Set UgXaml = Nothing
    Set Root = Nothing
    Set Reader = Nothing

    ' --- T: UniformGrid ItemsHost + TextBlock ItemTemplate ---
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
    Debug.Print "P7c-PANEL T ItemsSource set"

    If Not TypeOf IC.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "T ItemsHost expected UniformGrid"
    Set UgHost = IC.ItemsHost
    If UgHost.Children.Count <> 2 Then Err.Raise vbObjectError, , "T expected 2 UniformGrid children"
    Set Tb0 = UgHost.Children(0)
    Set Tb1 = UgHost.Children(1)
    If Tb0.Text <> "OK" Then Err.Raise vbObjectError, , "T Tb0 expected OK got " & Tb0.Text
    If Tb1.Text <> "Cancel" Then Err.Raise vbObjectError, , "T Tb1 expected Cancel got " & Tb1.Text
    Debug.Print "P7c-PANEL T UniformGrid+TextBlock OK"
    KeepAlive IC
    Set Tb0 = Nothing
    Set Tb1 = Nothing
    Set UgHost = Nothing
    Set IC = Nothing
    Set Tmpl = Nothing
    Set TbTmpl = Nothing
    Set PanelTmpl = Nothing
    Set UgProto = Nothing
    Set B = Nothing

    ' --- B: UniformGrid ItemsHost + Button ItemTemplate ---
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
    Debug.Print "P7c-PANEL B ItemsSource set"

    If Not TypeOf IC.ItemsHost Is UniformGrid Then Err.Raise vbObjectError, , "B ItemsHost expected UniformGrid"
    Set UgHost = IC.ItemsHost
    If UgHost.Children.Count <> 2 Then Err.Raise vbObjectError, , "B expected 2 UniformGrid children"
    Set Btn0 = UgHost.Children(0)
    Set Btn1 = UgHost.Children(1)
    If CStr(Btn0.Content) <> "OK" Then Err.Raise vbObjectError, , "B Btn0 Content expected OK"
    If CStr(Btn1.Content) <> "Cancel" Then Err.Raise vbObjectError, , "B Btn1 Content expected Cancel"
    Debug.Print "P7c-PANEL B UniformGrid+Button OK"
    KeepAlive IC
    Set Btn0 = Nothing
    Set Btn1 = Nothing
    Set UgHost = Nothing
    Set IC = Nothing
    Set Coll = Nothing
    Set OkItem = Nothing
    Set CancelItem = Nothing

    LogResult "P7c-PANEL", 0, "OK ItemsPanel UG shell + TextBlock/Button inflate"
    Debug.Print "PASS  P7c-PANEL ItemsPanelTemplate (UniformGrid items hardened)"
    Phase7cBench_ItemsPanelUniformGrid = True
    Exit Function

Fail:
    LogResult "P7c-PANEL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7c-PANEL ? " & Err.Description
    Phase7cBench_ItemsPanelUniformGrid = False
    On Error Resume Next
    KeepAlive IC
    KeepAlive Root
    Err.Clear
End Function

Private Sub AttachDialogButtonTemplateBindings(ByVal Btn As Button)
    Dim B As Binding

    Set B = New Binding
    Set B.TargetProperty = Btn.DependencyProperties.GetProperty("Content")
    Set B.Source = Btn.DependencyProperties.GetProperty("DataContext")
    B.Path = "Text"
    B.Mode = OneWay
    Set B.Target = Btn
    Btn.Bindings.Add B

    Set B = New Binding
    Set B.TargetProperty = Btn.DependencyProperties.GetProperty("Command")
    Set B.Source = Btn.DependencyProperties.GetProperty("DataContext")
    B.Path = "Command"
    B.Mode = OneWay
    Set B.Target = Btn
    Btn.Bindings.Add B

    Set B = New Binding
    Set B.TargetProperty = Btn.DependencyProperties.GetProperty("CommandParameter")
    Set B.Source = Btn.DependencyProperties.GetProperty("DataContext")
    B.Path = "Value"
    B.Mode = OneWay
    Set B.Target = Btn
    Btn.Bindings.Add B
End Sub

' Markup-equivalent: Source = DataContext DP, Path = Title (same as BindingsManager default).
Private Sub AttachDataContextTitleBinding(ByVal Tb As TextBlock)
    Dim B As Binding

    Set B = New Binding
    Set B.TargetProperty = Tb.DependencyProperties.GetProperty("Text")
    Set B.Source = Tb.DependencyProperties.GetProperty("DataContext")
    B.Path = "Title"
    Set B.Target = Tb
    Tb.Bindings.Add B
End Sub

Private Function LoadTextFile(ByVal Path As String) As String
    Dim Fn As Integer
    Fn = FreeFile
    Open Path For Input As #Fn
    LoadTextFile = Input$(LOF(Fn), #Fn)
    Close #Fn
End Function

Private Sub LogResult(ByVal Id As String, ByVal ElapsedMs As Long, ByVal Detail As String)
    LogLine Id & vbTab & CStr(ElapsedMs) & " ms" & vbTab & Detail
End Sub

Private Sub LogLine(ByVal Text As String)
    Dim Fn As Integer
    Fn = FreeFile
    Open App.Path & "\" & LOG_FILE For Append As #Fn
    Print #Fn, Format$(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Text
    Close #Fn
End Sub

Private Sub ClearLog()
    Dim Fn As Integer
    Fn = FreeFile
    Open App.Path & "\" & LOG_FILE For Output As #Fn
    Print #Fn, "Phase 0 benchmark log ? " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #Fn
End Sub
