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

    Debug.Print "=== Done: " & (8 - Failed) & " passed, " & Failed & " failed ==="
    If Failed > 0 Then
        MsgBox Failed & " Phase 0/1 test(s) failed. See Immediate window and " & LOG_FILE, vbExclamation, "Phase0"
    Else
        MsgBox "All Phase 0/1 tests passed.", vbInformation, "Phase0"
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
