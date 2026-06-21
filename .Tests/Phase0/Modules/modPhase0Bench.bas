Attribute VB_Name = "modPhase0Bench"
Option Explicit

Private Const LOG_FILE As String = "Phase0_bench.log"

Public Sub RunAll()
    Dim Failed As Long
    Failed = 0

    Debug.Print "=== Demac.VCF Phase 0 benchmarks ==="
    ClearLog

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

    Debug.Print "=== Done: " & (25 - Failed) & " passed, " & Failed & " failed ==="
    If Failed > 0 Then
        MsgBox Failed & " Phase 0/1/2/3/4/5 test(s) failed. See Immediate window and " & LOG_FILE, vbExclamation, "Phase0"
    Else
        MsgBox "All Phase 0/1/2/3/4/5 tests passed.", vbInformation, "Phase0"
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
