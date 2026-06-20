# Demac.VCF — performance benchmarks (Phase 0 baseline)

**Status:** Baselines recorded — **v2.1.0-wpf-alignment-p1** validated  
**Runner:** `.Tests/Phase0` (`modPhase0Bench`)  
**Threshold policy:** Regressions > 10% vs previous tag require explanation in release notes.

---

## Environment

| Field | Value |
|-------|-------|
| DLL version | 2.2.0-wpf-alignment-p1b |
| OS | Windows 10/11 x64 (build 26200) |
| vbRichClient5 | v5 (path in test `.vbp`) |
| Process bitness | 32-bit (VB6) |
| Recorded | 2026-06-20 |
| Phase0/1 result | **8 passed, 0 failed** (target after rebuild) |

Record on future runs: machine model, CPU, `Demac.VCF.dll` file date, and whether POS video secondary display is active (exclude from UI benchmarks).

---

## Phase 0 benchmarks

| ID | Scenario | Method | Baseline (ms) | Threshold |
|----|----------|--------|---------------|-----------|
| B-GOLD | Golden XAML load (minimal tree) | `Phase0Bench_GoldenXamlLoad` | **14** | ≤ 16 (10% margin) |
| B-COLL | 1000× `ObservableCollection.Add` | `Phase0Bench_CollectionAdd1000` | **19** | ≤ 21 (10% margin) |
| B-LCV | Two simultaneous `ListCollectionView` init | `Phase0Bench_DualListCollectionView` | **pass** | must not raise |
| B-STRICT-MALFORM | Malformed XAML raises `XamlLoadException` | `Phase0Bench_StrictMalformedXaml` | **pass** | must raise |
| B-STRICT-UNKNOWN | Unknown type raises `XamlLoadException` | `Phase0Bench_StrictUnknownType` | **pass** | must raise |
| P1-WIDTH | Panel `Width`/`Height` from XAML | `Phase1Bench_LayoutWidthXaml` | **pass** | — |
| P1-VIS | Panel `Visibility=Collapsed` DP | `Phase1Bench_PanelVisibilityCollapsed` | **pass** | must store Collapsed |
| P1-BORDER | Border `Width` from XAML | `Phase1Bench_BorderWidthXaml` | *TBD* | — |
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
| Collection Add | `New List` per notification | Single-item args |
| Binding graph | 3× WithEvents per binding | BindingExpression + Detach |
| Layout resize | Design* cascade | Measure/Arrange |

---

## Changelog

| Date | Change |
|------|--------|
| 2026-06-20 | Initial scaffold for Phase 0 |
| 2026-06-20 | Baselines recorded: B-GOLD 14 ms, B-COLL 16 ms; all Phase0 tests pass |
| 2026-06-20 | Validated build: B-COLL 19 ms; P1-WIDTH/P1-VIS pass (7/7) |
