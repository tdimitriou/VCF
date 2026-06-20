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

    Debug.Print "=== Done: " & (18 - Failed) & " passed, " & Failed & " failed ==="
    If Failed > 0 Then
        MsgBox Failed & " Phase 0/1/2/3/4 test(s) failed. See Immediate window and " & LOG_FILE, vbExclamation, "Phase0"
    Else
        MsgBox "All Phase 0/1/2/3/4 tests passed.", vbInformation, "Phase0"
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
    Phase4Bench_BindingOneWay = True
    Exit Function

Fail:
    LogResult "P4-BIND", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-BIND — " & Err.Description
    Phase4Bench_BindingOneWay = False
End Function

Public Function Phase4Bench_DataContextRebind() As Boolean
    Dim Vm1 As Phase0ViewModel
    Dim Vm2 As Phase0ViewModel
    Dim P As Panel
    Dim Tb As TextBlock
    Dim Expr As BindingExpression

    On Error GoTo Fail

    Set Vm1 = New Phase0ViewModel
    Vm1.Title = "One"
    Set Vm2 = New Phase0ViewModel
    Vm2.Title = "Two"

    Set P = New Panel
    Set Tb = New TextBlock
    P.Children.Add Tb

    Set P.DataContext = Vm1
    Set Expr = New BindingExpression
    Expr.Attach Tb, "Text", Tb.DependencyProperties.GetProperty("DataContext"), "Title", OneWay

    If Tb.Text <> "One" Then Err.Raise vbObjectError, , "Expected One, got " & Tb.Text

    Set P.DataContext = Vm2
    If Tb.Text <> "Two" Then Err.Raise vbObjectError, , "Expected Two after DataContext swap, got " & Tb.Text

    LogResult "P4-DCTX", 0, "OK DataContext rebind"
    Debug.Print "PASS  P4-DCTX DataContext rebind"
    Phase4Bench_DataContextRebind = True
    Exit Function

Fail:
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
    Phase4Bench_BindingDetach = True
    Exit Function

Fail:
    LogResult "P4-DETACH", 0, "FAIL: " & Err.Description
    Debug.Print "FAIL  P4-DETACH — " & Err.Description
    Phase4Bench_BindingDetach = False
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
