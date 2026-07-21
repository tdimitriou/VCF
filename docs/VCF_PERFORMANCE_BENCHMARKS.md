# Demac.VCF — performance benchmarks (Phase 0 baseline)

**Status:** Phase 7f — **P7d-LOAD** timed gates in DeNovoSmoke Immediate window (`ENABLE_LOAD_BENCH`)
**Runner:** `.Tests/Phase0` (`modPhase0Bench`)  
**Threshold policy:** Regressions > 10% vs previous tag require explanation in release notes.

> **B-GOLD disclaimer:** **B-GOLD** (~21 ms) loads a **minimal** golden panel only. It does **not** represent DeNovo **LoginView** / **SplashView** load time. Real POS screens include nested borders, ListView, numpad bindings, and `res:` fragments — use **P7d-LOAD-*** (planned) for regression gates.

---

## Environment

| Field | Value |
|-------|-------|
| DLL version | 2.15.0 (framework) · 2.16.0 (7a docs/tests) |
| OS | Windows 10/11 x64 (build 26200) |
| vbRichClient5 | v5 (path in test `.vbp`) |
| Process bitness | 32-bit (VB6) |
| Recorded | 2026-06-21 |
| Phase0/1/2/3/4/5/6/7 result | **30/30 pass** |

Record on future runs: machine model, CPU, `Demac.VCF.dll` file date, and whether POS video secondary display is active (exclude from UI benchmarks).

---

## Phase 0 benchmarks

| ID | Scenario | Method | Baseline (ms) | Threshold |
|----|----------|--------|---------------|-----------|
| B-GOLD | Golden XAML load (minimal tree) | `Phase0Bench_GoldenXamlLoad` | **21** | ≤ 24 (10% margin) |
| B-COLL | 1000× `ObservableCollection.Add` | `Phase0Bench_CollectionAdd1000` | **7** | ≤ 23 (10% margin) |
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
| P5a-OWN | Owner-draw ListView + XAML alias | `Phase5aBench_OwnerDrawListView` | **pass** | No ItemsSource |
| P5b-MSR | MeasureRow 40/20px rows | `Phase5bBench_MeasureRow` | **pass** | InvoiceGrid prep |
| P5c-HIER | QueryRowLevel parent/child indent | `Phase5cBench_RowLevel` | **pass** | InvoiceGrid prep |
| P6a-CONTENT | Button Content DP + Text alias + bind | `Phase6aBench_ButtonContent` | **pass** | WPF caption |
| P6b-TRIG | Style PropertyTrigger IsMouseOver | `Phase6bBench_PropertyTrigger` | **pass** | hover BackColor |
| P6c-TMPL | ControlTemplate Border chrome | `Phase6cBench_ControlTemplate` | **pass** | Button CornerRadius |
| P6d-COAL | Render refresh coalescing | `Phase6dBench_RenderCoalesce` | **pass** | dedupe + Style batch |
| P7a-SMOKE | POS SalesOrder shell XAML | `Phase7aBench_PosSalesOrderShell` | **pass** | Scene + UniformGrid |
| B-RESZ | Window resize nested UniformGrid 50× | `Phase2aBench_NestedUniformGridResize` | *record* | Phase 2a — Widget.Move ×50 |
| B-NAV | 50× view navigation binding leak | `Phase2aBench_ViewNavLeak` | *record* | Phase 2a — Visibility nav + Windows=0 |
| B-BIND-DENSE | 21×6 ItemTemplate clone + bindings | `Phase2aBench_ListViewBindHotspot` | *record* | Phase 2a — clone bind fidelity + INPC + detach |
| P2a-PAD | ListView Margin/Padding defaults | `Phase2aBench_ListViewPaddingDefaults` | pass | Margin=0; Padding=4,1,4,1 |
| P2a-PAD-TB | TextBox/Button Margin/Padding defaults | `Phase2aBench_TextBoxButtonPaddingDefaults` | pass | TextBox 0/1; Button Padding=1 |
| P2a-PAD-UG | UniformGrid Padding default | `Phase2aBench_UniformGridPaddingDefault` | pass | Padding=2 locked |
| P7c-DLG | Dialog Button DataTemplate | `Phase7cBench_DialogDataTemplate` | pass | ItemsControl+Button Content/Command |
| P7c-PANEL | ItemsPanelTemplate | `Phase7cBench_ItemsPanelUniformGrid` | pass | Code+XAML UniformGrid ItemsPanel shell (no item inflate; Button covered by P7c-DLG) |
| P6e-PRES | ContentPresenter paint-only | `Phase6eBench_ContentPresenter` | pass | Caption path + SuppressContent with children |

---

## Phase 7f load benchmarks (DeNovoSmoke)

Timed gates for **real POS fixtures**. Logs to Immediate window when `ENABLE_LOAD_BENCH = True` in `modHarnessConfig.bas`.

| ID | Scenario | Fixture | Baseline | Threshold |
|----|----------|---------|----------|-----------|
| P7d-LOAD-SPLASH | Splash `LoadView` | `Screens\Splash\SplashView` | **24–41 ms** (8a: **24 ms** 2026-07-20) | Soft: &lt; 100 ms |
| P7d-LOAD-LOGIN | Login `LoadView` (lazy, first ShowLogin) | `LoginViewWpf` + LoginPad | **403–670 ms** (8a: **403 ms** 2026-07-20) | Soft: &lt; 1000 ms |
| P7d-LOAD-BORDERED | Bordered window XAML load | `BorderTestWindow.xml` (Shift+B) | *not recorded this run* | TBD |
| P7d-LAY-RESIZE | Border Design* children scale with host | Phase0 `BorderDesignChildren.xml` | **pass** (Option B) | Half size → ~0.5 child geometry (±2 px) |
| P8-INHERIT | Lazy GetValue inherit + DataContext | Phase0 `InheritanceNestedBorder.xml` | **pass** | PassPropertyValueCalls during load = 0; Mid/Inner DataContext pull from root |
| P7d-SHUTDOWN | Full harness session → `RemoveAll` | DeNovoSmoke E2E | **pass** (2026-07-18) | No error 91; IDE second F5 |
| B-CHROME | First-frame window chrome | Manual / screenshot | — | Borderless shell + bordered dialog, no flash |

**Notes (2026-07-18 first 7f run):** Splash-only startup is cheap; Login dominates (~13× Splash) because of nested Border/Grid/ListView/LoginPad + bindings. Lazy policy confirmed — `[P7d-LOAD-LOGIN]` appears only after ShowLogin. Shutdown tears down Binding/`UIElementBase` terminations orderly (HarnessScreen → StubLoginViewModel last among views). **`VCF_SHUTDOWN_DIAG` instrumentation has been removed.**

**Notes (2026-07-20 Phase 8b):** Lazy `GetValue` parent-walk; attach notify deferred inside `InheritanceBatch` (single End wave). IDE DeNovoSmoke (4 runs, both in IDE): Splash **27–41 ms**, Login **465–543 ms** (avg ~503). Prior: broken 8b ~831 → double-notify fix ~659 → coalesce **this band**; 8a IDE baseline Login **~403 ms**. Treat as relative; compiled VCF should tighten further.

**XAML load batching (framework):** `XAMLReader.LoadSuperclassData` wraps tree build in `LockRefresh` + `BeginRenderUpdate` / `EndRenderUpdate` ([alignment §2.7.4 P0 #1](./VCF_WPF_ALIGNMENT_NOTES.md)).

**Profiling split (record in bench log):**

1. `InitializeApplication` — MyApp + resource dictionary  
2. Per-screen `LoadView` / `LoadSuperclassData`  
3. First `Show` + root `Refresh`

**Harness policy (7f):** lazy **Login** load on first `ShowLogin` (`EAGER_LOGIN_LOAD = False`). Set `EAGER_LOGIN_LOAD = True` only for m1 A/B comparison — do not compare startup times across policies without noting it.

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
| XAML tree load | Per-node refresh during parse | **7f:** LockRefresh + BeginRenderUpdate in LoadSuperclassData |
| DataContext push | O(n) on large trees | **8:** lazy inheritance §2.8 |

---

## Changelog

| Date | Change |
|------|--------|
| 2026-07-21 | **B-BIND-DENSE** — ItemTemplate clone preserves TextBlock bindings; Phase0 21×6 + ItemsControl (34 tests) |
| 2026-07-21 | **B-RESZ** / **B-NAV** Phase0 benches (33 tests); Option C deferred; 2a.1 Visibility auto Relayout done |
| 2026-07-18 | First 7f baselines — Splash **35 ms**, Login **448 ms**; shutdown pass |
| 2026-07-20 | Phase 8a measured — Splash **24 ms**, Login **403 ms** (vs prior Login ~448–670 ms) |
| 2026-07-20 | Phase 8a — inheritance batch; **P8-INHERIT**; soft Login ms compare after rebuild |
| 2026-07-20 | Phase 8b IDE×4 — Splash **27–41 ms**, Login **465–543 ms** (vs 8a Login ~403; vs broken 8b ~831) |
| 2026-07-18 | Layout Option B — Border Design* LegacyScaleLayout; **P7d-LAY-RESIZE** Phase0 gate |
| 2026-07-18 | Removed `VCF_SHUTDOWN_DIAG` instrumentation (modShutdownDiag + harness/framework Trace*) |
| 2026-07-18 | Phase 7f implemented — XAML LoadSuperclassData batching; DeNovoSmoke lazy Login + P7d-LOAD-* Immediate logs |
| 2026-06-20 | Phase 7f planned — P7d-LOAD-* table, B-GOLD disclaimer, profiling split, harness lazy-Login note |
| 2026-06-20 | Initial scaffold for Phase 0 |
| 2026-06-20 | Baselines recorded: B-GOLD 14 ms, B-COLL 16 ms; all Phase0 tests pass |
| 2026-06-20 | Validated build: B-COLL 19 ms; P1-WIDTH/P1-VIS pass (7/7) |
| 2026-06-20 | **v2.4.0 Phase 3 validated:** 15/15 pass; B-GOLD 22 ms, B-COLL 21 ms; P3-MERGE, P3-SOURCE, P3-DYNAMIC, P3-STRICT-PROP pass |
| 2026-06-20 | **v2.4.0 Phase 3:** P3-MERGE, P3-SOURCE, P3-DYNAMIC, P3-STRICT-PROP added (15 total tests) |
| 2026-06-20 | **v2.3.0 Phase 2 validated:** 11/11 pass; B-GOLD 19 ms, B-COLL 16 ms; P2-STACK, P2-STACK-LAY, P2-GRID pass |
| 2026-06-21 | **v2.17.0 Phase 7b:** Invoke-VcfXamlMigration.ps1 + XAML_MIGRATION_PROMPTS (no new Phase0 test) |
| 2026-06-21 | **v2.16.0 Phase 7a validated:** 30/30 pass; B-GOLD **19 ms**, B-COLL **6 ms**; P7a-SMOKE pass |
| 2026-06-21 | **v2.16.0 Phase 7a started:** P7a-SMOKE, POS_INTEGRATION_SMOKE, MIGRATION 2.15 pin guide |
| 2026-06-21 | **v2.15.0 Phase 6d validated:** 29/29 pass; B-GOLD **21 ms**, B-COLL **7 ms**; P6d-COAL pass |
| 2026-06-21 | **v2.15.0 Phase 6d started:** P6d-COAL added (29 tests); render refresh coalescing |
| 2026-06-21 | **v2.14.0 Phase 6c validated:** 28/28 pass; B-GOLD **17 ms**, B-COLL **3 ms**; P6c-TMPL pass |
| 2026-06-21 | **v2.14.0 Phase 6c started:** P6c-TMPL added (28 tests); ControlTemplate + Style.Template |
| 2026-06-21 | **v2.13.0 Phase 6b validated:** 27/27 pass; B-GOLD **63 ms**, B-COLL **4 ms**; P6b-TRIG pass |
| 2026-06-21 | **v2.13.0 Phase 6b started:** P6b-TRIG added (27 tests); PropertyTrigger + Style.Triggers |
| 2026-06-21 | **v2.12.0 Phase 6a validated:** 26/26 pass; B-GOLD **19 ms**, B-COLL **6 ms**; P6a-CONTENT pass |
| 2026-06-21 | **v2.12.0 Phase 6a started:** P6a-CONTENT added (26 tests); Button Content DP + Text alias |
| 2026-06-21 | **v2.11.0 Phase 5c validated:** 25/25 pass; B-GOLD **19 ms**, B-COLL **7 ms**; P5c-HIER pass |
| 2026-06-21 | **v2.10.0 Phase 5b validated:** 24/24 pass; B-GOLD **18 ms**, B-COLL **5 ms**; P5b-MSR pass |
| 2026-06-21 | **v2.9.0 Phase 5a validated:** 23/23 pass; B-GOLD **19 ms**, B-COLL **5 ms**; P5a-OWN pass |
| 2026-06-21 | **v2.8.0 Phase 4d validated:** 22/22 pass; B-GOLD **23 ms**, B-COLL **9 ms**; P4d-SEL pass |
| 2026-06-20 | **v2.7.0 Phase 4c validated:** 21/21 pass; B-GOLD **20 ms**, B-COLL **7 ms**; P4b-ICtrl pass |
| 2026-06-20 | **v2.6.0 Phase 4b validated:** 20/20 pass; B-COLL **3 ms** (scratch buffers); P4b-DEFER, P4b-MOVE pass |
| 2026-06-20 | **v2.5.0 Phase 4 validated:** 18/18 pass; P4-BIND, P4-DCTX, P4-DETACH pass; binding detach no longer hangs on target read |
