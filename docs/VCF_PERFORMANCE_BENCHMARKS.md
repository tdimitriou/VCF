# Demac.VCF — performance benchmarks (Phase 0 baseline)

**Status:** Phase 4d — **v2.8.0-wpf-alignment-p4d** (22 tests, validated)  
**Runner:** `.Tests/Phase0` (`modPhase0Bench`)  
**Threshold policy:** Regressions > 10% vs previous tag require explanation in release notes.

---

## Environment

| Field | Value |
|-------|-------|
| DLL version | 2.8.0-wpf-alignment-p4d |
| OS | Windows 10/11 x64 (build 26200) |
| vbRichClient5 | v5 (path in test `.vbp`) |
| Process bitness | 32-bit (VB6) |
| Recorded | 2026-06-21 |
| Phase0/1/2/3/4 result | **22/22 pass** |

Record on future runs: machine model, CPU, `Demac.VCF.dll` file date, and whether POS video secondary display is active (exclude from UI benchmarks).

---

## Phase 0 benchmarks

| ID | Scenario | Method | Baseline (ms) | Threshold |
|----|----------|--------|---------------|-----------|
| B-GOLD | Golden XAML load (minimal tree) | `Phase0Bench_GoldenXamlLoad` | **23** | ≤ 24 (10% margin) |
| B-COLL | 1000× `ObservableCollection.Add` | `Phase0Bench_CollectionAdd1000` | **9** | ≤ 23 (10% margin) |
| B-LCV | Two simultaneous `ListCollectionView` init | `Phase0Bench_DualListCollectionView` | **pass** | must not raise |
| B-STRICT-MALFORM | Malformed XAML raises `XamlLoadException` | `Phase0Bench_StrictMalformedXaml` | **pass** | must raise |
| B-STRICT-UNKNOWN | Unknown type raises `XamlLoadException` | `Phase0Bench_StrictUnknownType` | **pass** | must raise |
| P1-WIDTH | Panel `Width`/`Height` from XAML | `Phase1Bench_LayoutWidthXaml` | **pass** | — |
| P1-VIS | Panel `Visibility=Collapsed` DP | `Phase1Bench_PanelVisibilityCollapsed` | **pass** | must store Collapsed |
| P1-BORDER | Border `Width` from XAML | `Phase1Bench_BorderWidthXaml` | **pass** | — |
| P2-STACK | StackPanel Width/Orientation XAML | `Phase2Bench_StackPanelXaml` | **pass** | — |
| P2-STACK-LAY | Vertical stack child positions | `Phase2Bench_StackPanelLayout` | **pass** | P2.Top ≈ 50 |
| P2-GRID | Grid RowDefinitions XAML | `Phase2Bench_GridRowDefinitionsXaml` | **pass** | 2 rows, 2 cols |
| P3-MERGE | Merged ResourceDictionary lookup | `Phase3Bench_MergedDictionaryLookup` | **pass** | TryGetResource |
| P3-SOURCE | `Source=` file load | `Phase3Bench_ResourceSourceLoad` | **pass** | Greeting=Phase3 |
| P3-DYNAMIC | Element TryFindResource path | `Phase3Bench_DynamicResourceExtension` | **pass** | BgColor=12345 |
| P3-STRICT-PROP | Unknown property raises | `Phase3Bench_StrictUnknownProperty` | **pass** | must raise |
| P4-BIND | OneWay binding + INPC | `Phase4Bench_BindingOneWay` | **pass** | Title sync |
| P4-DCTX | DataContext swap rebind | `Phase4Bench_DataContextRebind` | **pass** | One→Two |
| P4-DETACH | Detach stops updates | `Phase4Bench_BindingDetach` | **pass** | Text stays Before |
| P4b-DEFER | BeginUpdate coalesces 100 adds | `Phase4bBench_BeginUpdateDefer` | **pass** | 1 Reset notify |
| P4b-MOVE | Move(0,2) reorder | `Phase4bBench_Move` | **pass** | b,c,a |
| P4b-ICtrl | ItemsControl item generation | `Phase4bBench_ItemsControl` | **pass** | 3 items |
| P4d-SEL | Selector SelectedIndex/Value | `Phase4dBench_Selector` | **pass** | ListView + Selector |
| B-RESZ | Window resize nested UniformGrid 50× | *Phase 1+* | — | deferred |
| B-NAV | 50× view navigation binding leak | *Phase 4+* | — | deferred |

---

## How to run

1. Build `Demac.VCF.dll` (Release).
2. Open `.Tests/Phase0/Phase0.vbp` in VB6 IDE.
3. Run (F5). Results print to Immediate window and log file `Phase0_bench.log`.

---

## POS telemetry context

Normal POS process **< 100 MB** without secondary customer-display video. Framework targets (Phases 1–4):

| Metric | Issue | Target |
|--------|-------|--------|
| DP registration | Per-instance Register × N | Shared registry (~10–25 MB savings est.) |
| Collection Add | `New List` per notification | Single-item scratch buffers + batch Reset |
| Binding graph | 3× WithEvents per binding | BindingExpression + Detach |
| Layout resize | Design* cascade | Measure/Arrange |

---

## Changelog

| Date | Change |
|------|--------|
| 2026-06-20 | Initial scaffold for Phase 0 |
| 2026-06-20 | Baselines recorded: B-GOLD 14 ms, B-COLL 16 ms; all Phase0 tests pass |
| 2026-06-20 | Validated build: B-COLL 19 ms; P1-WIDTH/P1-VIS pass (7/7) |
| 2026-06-20 | **v2.4.0 Phase 3 validated:** 15/15 pass; B-GOLD 22 ms, B-COLL 21 ms; P3-MERGE, P3-SOURCE, P3-DYNAMIC, P3-STRICT-PROP pass |
| 2026-06-20 | **v2.4.0 Phase 3:** P3-MERGE, P3-SOURCE, P3-DYNAMIC, P3-STRICT-PROP added (15 total tests) |
| 2026-06-20 | **v2.3.0 Phase 2 validated:** 11/11 pass; B-GOLD 19 ms, B-COLL 16 ms; P2-STACK, P2-STACK-LAY, P2-GRID pass |
| 2026-06-21 | **v2.8.0 Phase 4d validated:** 22/22 pass; B-GOLD **23 ms**, B-COLL **9 ms**; P4d-SEL pass |
| 2026-06-20 | **v2.7.0 Phase 4c validated:** 21/21 pass; B-GOLD **20 ms**, B-COLL **7 ms**; P4b-ICtrl pass |
| 2026-06-20 | **v2.6.0 Phase 4b validated:** 20/20 pass; B-COLL **3 ms** (scratch buffers); P4b-DEFER, P4b-MOVE pass |
| 2026-06-20 | **v2.5.0 Phase 4 validated:** 18/18 pass; P4-BIND, P4-DCTX, P4-DETACH pass; binding detach no longer hangs on target read |
