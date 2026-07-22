Attribute VB_Name = "modPhase0Bench"
Option Explicit

Private Const LOG_FILE As String = "Phase0_bench.log"

' Hold Button/ItemsControl trees for the IDE session — releasing them (Terminate /
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
    If Not Phase1Bench_BorderWidthXaml() Then Failed = Failed + 1
    If Not Phase2Bench_StackPanelXaml() Then Failed = Failed + 1
    If Not Phase2Bench_StackPanelLayout() Then Failed = Failed + 1
    If Not Phase2Bench_GridRowDefinitionsXaml() Then Failed = Failed + 1
    If Not Phase3Bench_MergedDictionaryLookup() Then Failed = Failed + 1
    If Not Phase3Bench_ResourceSourceLoad() Then Failed = Failed + 1
    If Not Phase3Bench_DynamicResourceExtension() Then Failed = Failed + 1
    If Not Phase3Bench_StrictUnknownProperty() Then Failed = Failed + 1
    If Not Phase4Bench_BindingOneWay() Then Failed = Failed + 1
    If Not Phase4Bench_DataContextRebind() Then Failed = Failed + 1
    If Not Phase4Bench_BindingDetach() Then Failed = Failed + 1
    If Not Phase4bBench_BeginUpdateDefer() Then Failed = Failed + 1
    If Not Phase4bBench_Move() Then Failed = Failed + 1
    If Not Phase4bBench_ItemsControl() Then Failed = Failed + 1
    If Not Phase4dBench_Selector() Then Failed = Failed + 1
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
    If Not Phase6fBench_TemplateBindingSlot() Then Failed = Failed + 1
    If Not Phase7aBench_PosSalesOrderShell() Then Failed = Failed + 1
    If Not Phase7cBench_LegacyLayoutShim() Then Failed = Failed + 1
    If Not Phase7dBench_BorderDesignResize() Then Failed = Failed + 1
    If Not Phase8Bench_InheritanceBatch() Then Failed = Failed + 1
    If Not Phase2aBench_NestedUniformGridResize() Then Failed = Failed + 1
    If Not Phase2aBench_ViewNavLeak() Then Failed = Failed + 1
    If Not Phase2aBench_ListViewBindHotspot() Then Failed = Failed + 1
    If Not Phase2aBench_ListViewPaddingDefaults() Then Failed = Failed + 1
    If Not Phase2aBench_TextBoxButtonPaddingDefaults() Then Failed = Failed + 1
    If Not Phase2aBench_UniformGridPaddingDefault() Then Failed = Failed + 1
    If Not Phase7cBench_DialogDataTemplate() Then Failed = Failed + 1
    If Not Phase7cBench_ItemsPanelUniformGrid() Then Failed = Failed + 1

    ' Report only — do not RemoveAll / release KeepAlive here (Button ItemsHost
    ' Terminate after MsgBox silently crashes the IDE).
    Debug.Print "=== Done: " & (43 - Failed) & " passed, " & Failed & " failed ==="
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
    Debug.Print "FAIL  B-GOLD — " & Err.Description
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
    Debug.Print "FAIL  B-COLL — " & Err.Description
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
    Debug.Print "FAIL  B-LCV — " & Err.Description
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
        Debug.Print "FAIL  B-STRICT Malformed — " & Err.Description
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
        Debug.Print "FAIL  B-STRICT Unknown — " & Err.Description
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
    Debug.Print "FAIL  P1-WIDTH — " & Err.Description
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
    Debug.Print "FAIL  P1-VIS — " & Err.Description
    Phase1Bench_PanelVisibilityCollapsed = False
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
    Debug.Print "FAIL  P1-BORDER — " & Err.Description
    Phase1Bench_BorderWidthXaml = False
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
    Debug.Print "FAIL  P2-STACK — " & Err.Description
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
    Debug.Print "FAIL  P2-STACK-LAY — " & Err.Description
    Phase2Bench_StackPanelLayout = False
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
    Debug.Print "FAIL  P2-GRID — " & Err.Description
    Phase2Bench_GridRowDefinitionsXaml = False
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
    Debug.Print "FAIL  P3-MERGE — " & Err.Description
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
    Debug.Print "FAIL  P3-SOURCE — " & Err.Description
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
    Debug.Print "FAIL  P3-DYNAMIC — " & Err.Description
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
        Debug.Print "FAIL  P3-STRICT — " & Err.Description
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
    Debug.Print "FAIL  P4-BIND — " & Err.Description
    Phase4Bench_BindingOneWay = False
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
    Debug.Print "FAIL  P4-DCTX — " & Err.Description
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
    If Tb.Text <> "Before" Then Err.Raise vbObjectError, , "Expected text frozen at Before, got " & Tb.Text

    LogResult "P4-DETACH", 0, "OK Detach stops updates"
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
    Debug.Print "FAIL  P4-DETACH — " & Err.Description
    Phase4Bench_BindingDetach = False
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
    Debug.Print "FAIL  P4b-DEFER — " & Err.Description
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
    Debug.Print "FAIL  P4b-MOVE — " & Err.Description
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
    Debug.Print "FAIL  P4b-ICtrl — " & Err.Description
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
    Debug.Print "FAIL  P4d-SEL — New ListView: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailSource:
    LogResult "P4d-SEL", 0, "FAIL at ListView ItemsSource: " & Err.Description
    Debug.Print "FAIL  P4d-SEL — ListView ItemsSource: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailIndex1:
    LogResult "P4d-SEL", 0, "FAIL at ListView SelectedIndex=1: " & Err.Description
    Debug.Print "FAIL  P4d-SEL — ListView SelectedIndex=1: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailIndex2:
    LogResult "P4d-SEL", 0, "FAIL at ListView SelectedIndex=2: " & Err.Description
    Debug.Print "FAIL  P4d-SEL — ListView SelectedIndex=2: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

FailSelector:
    LogResult "P4d-SEL", 0, "FAIL at Selector: " & Err.Description
    Debug.Print "FAIL  P4d-SEL — Selector: " & Err.Description
    Phase4dBench_Selector = False
    Exit Function

Fail:
    LogResult "P4d-SEL", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4d-SEL — " & Err.Description
    Phase4dBench_Selector = False
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
    Debug.Print "FAIL  P5a-OWN — " & Err.Description
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
    Debug.Print "FAIL  P5b-MSR — " & Err.Description
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
    Debug.Print "FAIL  P5c-HIER — " & Err.Description
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
    Debug.Print "FAIL  P6a-CONTENT — " & Err.Description
    Phase6aBench_ButtonContent = False
End Function

Public Function Phase6bBench_PropertyTrigger() As Boolean
    Dim St As Style
    Dim Btn As Button
    Dim Trig As PropertyTrigger

    On Error GoTo Fail

    Set St = NewStyle("Button")
    St.SetSetter "BackColor", CLng(16777215)
    St.SetSetter "HoverColor", CLng(-1)

    Set Trig = New PropertyTrigger
    Trig.Initialize "IsMouseOver", "True"
    Trig.SetSetter "BackColor", CLng(255)
    St.AddTrigger Trig

    Set Btn = New Button
    Set Btn.Style = St

    If Btn.Widget.BackColor <> 16777215 Then Err.Raise vbObjectError, , "Expected base BackColor 16777215, got " & Btn.Widget.BackColor

    Btn.IsMouseOver = True
    If Btn.Widget.BackColor <> 255 Then Err.Raise vbObjectError, , "Expected hover BackColor 255, got " & Btn.Widget.BackColor

    Btn.IsMouseOver = False
    If Btn.Widget.BackColor <> 16777215 Then Err.Raise vbObjectError, , "Expected restored BackColor 16777215, got " & Btn.Widget.BackColor

    LogResult "P6b-TRIG", 0, "OK IsMouseOver PropertyTrigger BackColor"
    Debug.Print "PASS  P6b-TRIG Style PropertyTrigger IsMouseOver"
    Phase6bBench_PropertyTrigger = True
    Exit Function

Fail:
    LogResult "P6b-TRIG", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6b-TRIG — " & Err.Description
    Phase6bBench_PropertyTrigger = False
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
    Debug.Print "FAIL  P6c-TMPL — " & Err.Description
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
    Debug.Print "FAIL  P6d-COAL — " & Err.Description
    Phase6dBench_RenderCoalesce = False
End Function

' §2.11 ContentPresenter paint-only path (Button caption delegates here).
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
    Debug.Print "FAIL  P6e-PRES — " & Err.Description
    Phase6eBench_ContentPresenter = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
End Function

' §2.11 ContentPresenter / Button HorizontalContentAlignment + VerticalContentAlignment.
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
    Debug.Print "FAIL  P6e-ALIGN — " & Err.Description
    Phase6eBench_ContentAlignment = False
    On Error Resume Next
    KeepAlive Btn
    KeepAlive Root
    Err.Clear
End Function

' §2.11 ContentControl shares Button Content model (string presenter + IUIElement child).
Public Function Phase6eBench_ContentControlContent() As Boolean
    Dim CC As ContentControl
    Dim Tb As TextBlock
    Dim Reader As XAMLReader
    Dim Root As ContentControl
    Dim Child As Object

    On Error GoTo Fail

    Set CC = New ContentControl
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
    Debug.Print "FAIL  P6e-CC — " & Err.Description
    Phase6eBench_ContentControlContent = False
    On Error Resume Next
    KeepAlive CC
    KeepAlive Root
    Err.Clear
End Function

' §2.11 lookless-prep: ControlTemplate Border chrome + ContentPresenter marker (no live widgets).
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
    ' Explicit slot — TypeOf ContentPresenter is unreliable across EXE/DLL boundary.
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
    If Btn.Children.Count <> 0 Then Err.Raise vbObjectError, , "Template must not add live Button children"
    If CStr(Btn.Content) <> "OK" Then Err.Raise vbObjectError, , "Content expected OK (TemplateBinding from parent)"

    KeepAlive Btn
    Set Btn = Nothing
    Set St = Nothing
    Set Tmpl = Nothing
    Set B = Nothing
    Set CP = Nothing

    LogResult "P6f-TBIND", 0, "OK ControlTemplate Border+ContentPresenter marker (no live tree)"
    Debug.Print "PASS  P6f-TBIND template ContentPresenter slot"
    Phase6fBench_TemplateBindingSlot = True
    Exit Function

Fail:
    LogResult "P6f-TBIND", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P6f-TBIND — " & Err.Description
    Phase6fBench_TemplateBindingSlot = False
    On Error Resume Next
    KeepAlive Btn
    Err.Clear
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
    Debug.Print "FAIL  P7a-SMOKE — " & Err.Description
    Phase7aBench_PosSalesOrderShell = False
End Function

Public Function Phase7cBench_LegacyLayoutShim() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Panel
    Dim Tb As TextBlock

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\PosMigratedTextBlockLayout.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "POS migrated layout returned Nothing"
    If Root.Children.Count <> 1 Then Err.Raise vbObjectError, , "Expected 1 child, got " & Root.Children.Count

    Set Tb = Root.Children(0)
    If TypeName(Tb) <> "TextBlock" Then Err.Raise vbObjectError, , "Expected TextBlock, got " & TypeName(Tb)
    If Tb.DesignLeft <> 10# Then Err.Raise vbObjectError, , "Expected DesignLeft=10, got " & Tb.DesignLeft
    If Tb.DesignTop <> 20# Then Err.Raise vbObjectError, , "Expected DesignTop=20, got " & Tb.DesignTop
    If Tb.DesignWidth <> 200# Then Err.Raise vbObjectError, , "Expected DesignWidth=200, got " & Tb.DesignWidth
    If Tb.DesignHeight <> 30# Then Err.Raise vbObjectError, , "Expected DesignHeight=30, got " & Tb.DesignHeight

    LogResult "P7c-LAY", 0, "OK Margin/Width/Height on TextBlock -> Design*"
    Debug.Print "PASS  P7c-LAY legacy layout shim (migrated TextBlock)"
    Phase7cBench_LegacyLayoutShim = True
    Exit Function

Fail:
    LogResult "P7c-LAY", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7c-LAY — " & Err.Description
    Phase7cBench_LegacyLayoutShim = False
End Function

Public Function Phase7dBench_BorderDesignResize() As Boolean
    Dim Reader As XAMLReader
    Dim Root As Border
    Dim A As TextBlock
    Dim B As TextBlock

    On Error GoTo Fail

    Set Reader = New XAMLReader
    Set Root = Reader.Load(LoadTextFile(App.Path & "\Resources\BorderDesignChildren.xml"))

    If Root Is Nothing Then Err.Raise vbObjectError, , "BorderDesignChildren returned Nothing"
    If Root.Children.Count <> 2 Then Err.Raise vbObjectError, , "Expected 2 children, got " & Root.Children.Count

    Set A = Root.Children(0)
    Set B = Root.Children(1)

    ' Border.Move requires Parent; Widget.Move raises W_Resize -> ArrangeBorderChild (Option B).
    Root.Widget.Move 0, 0, 400, 300
    If Abs(A.Widget.Left - 40!) > 2! Then Err.Raise vbObjectError, , "Full A.Left expected ~40, got " & A.Widget.Left
    If Abs(A.Widget.Width - 200!) > 2! Then Err.Raise vbObjectError, , "Full A.Width expected ~200, got " & A.Widget.Width
    If Abs(B.Widget.Left - 100!) > 2! Then Err.Raise vbObjectError, , "Full B.Left expected ~100, got " & B.Widget.Left

    Root.Widget.Move 0, 0, 200, 150
    If Abs(A.Widget.Left - 20!) > 2! Then Err.Raise vbObjectError, , "Half A.Left expected ~20, got " & A.Widget.Left
    If Abs(A.Widget.Width - 100!) > 2! Then Err.Raise vbObjectError, , "Half A.Width expected ~100, got " & A.Widget.Width
    If Abs(B.Widget.Left - 50!) > 2! Then Err.Raise vbObjectError, , "Half B.Left expected ~50, got " & B.Widget.Left
    If Abs(B.Widget.Width - 40!) > 2! Then Err.Raise vbObjectError, , "Half B.Width expected ~40, got " & B.Widget.Width

    LogResult "P7d-LAY-RESIZE", 0, "OK Design* scale 400x300 -> 200x150"
    Debug.Print "PASS  P7d-LAY-RESIZE Border Design* children scale with host"
    Phase7dBench_BorderDesignResize = True
    Exit Function

Fail:
    LogResult "P7d-LAY-RESIZE", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P7d-LAY-RESIZE — " & Err.Description
    Phase7dBench_BorderDesignResize = False
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
    Debug.Print "FAIL  P8-INHERIT — " & Err.Description
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

    ' Loop ends on odd i=49 → half size; i=50 even → full. Assert full cell size.
    If Abs(Inner.Widget.Width - 200!) > 3! Then Err.Raise vbObjectError, , "After 50× resize Inner.Width expected ~200, got " & Inner.Widget.Width
    If Abs(Inner.Widget.Height - 150!) > 3! Then Err.Raise vbObjectError, , "After 50× resize Inner.Height expected ~150, got " & Inner.Widget.Height

    LogResult "B-RESZ", ElapsedMs, "OK nested UniformGrid 50x resize"
    Debug.Print "PASS  B-RESZ nested UniformGrid resize x50 (" & ElapsedMs & " ms)"
    Phase2aBench_NestedUniformGridResize = True
    Exit Function

Fail:
    LogResult "B-RESZ", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  B-RESZ — " & Err.Description
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

    Win.Dispose
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
    If Not Win Is Nothing Then Win.Dispose
    Cairo.WidgetForms.RemoveAll
    Set Shell = Nothing
    Set AppHost = Nothing
    VCF.ClearApplication
    LogResult "B-NAV", 0, "FAIL: " & FailDesc
    Debug.Print "FAIL  B-NAV — " & FailDesc
    Phase2aBench_ViewNavLeak = False
End Function

' ListView bind hotspot (framework-first): menu-like density = 21 rows × 6 DataContext bindings/cell.
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
    Debug.Print "FAIL  B-BIND-DENSE — (" & FailNum & ") " & FailDesc
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
    Debug.Print "FAIL  P2a-PAD — " & Err.Description
    Phase2aBench_ListViewPaddingDefaults = False
End Function

' Phase 2a Margin/Padding family 2: TextBox (Margin=0, Padding=1) + Button (Padding=1 Aero2).
' Qualify VCF.TextBox — bare TextBox resolves to VB.TextBox (higher in project refs).
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
    Debug.Print "FAIL  P2a-PAD-TB — " & Err.Description
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
    Debug.Print "FAIL  P2a-PAD-UG — " & Err.Description
    Phase2aBench_UniformGridPaddingDefault = False
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
    BtnTmpl.DesignWidth = 80
    BtnTmpl.DesignHeight = 28
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

    ' Keepalive — releasing Button ItemsHost mid-suite disconnects widgets (RPC_E_DISCONNECTED).
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
    Debug.Print "FAIL  P7c-DLG — " & Err.Description
    Phase7cBench_DialogDataTemplate = False
    On Error Resume Next
    KeepAlive IC
    Err.Clear
End Function

' ItemsPanelTemplate gate — UniformGrid shell + ItemsSource inflate (TextBlock then Button).
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
    BtnTmpl.DesignWidth = 80
    BtnTmpl.DesignHeight = 28
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
    Debug.Print "FAIL  P7c-PANEL — " & Err.Description
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
    Print #Fn, "Phase 0 benchmark log — " & Format$(Now, "yyyy-mm-dd hh:nn:ss")
    Close #Fn
End Sub
