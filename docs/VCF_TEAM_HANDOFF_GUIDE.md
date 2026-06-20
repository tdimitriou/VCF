# Demac.VCF — complete team handoff guide

**Status:** Ready for VCF team review  
**Audience:** Demac VCF maintainers (primary) · POS / DeNovo (requirements, migration)  
**VCF source repo:** `Projects\Demac\Framework\Demac.VCF` → ships as `Demac.VCF.dll`  
**Last updated:** 2026-06-19  

---

## How to use this package

This is the **single entry point** for the VCF rewrite program. Read in order for onboarding; jump by role for execution.

| # | Document | Purpose | Read when |
|---|----------|---------|-----------|
| 1 | **This guide** | Scope, roles, release contract, doc map | First |
| 2 | [VCF_WPF_ALIGNMENT_NOTES.md](./VCF_WPF_ALIGNMENT_NOTES.md) | Decisions, discussion log, §2 deep dives, §9 agreed direction | Understanding *why* |
| 3 | [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md) | Per-type rewrite actions, phases, bugs, tests | Planning sprints |
| 4 | [VCF_CLASS_REFERENCE.md](./VCF_CLASS_REFERENCE.md) | **Every class**: public API, line count, private behavior, issues | Implementation |
| 5 | [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md) | **Every dependency property** today + target registry | DP / binding work |
| 6 | [VCF_XAML_WPF_SUBSET.md](./VCF_XAML_WPF_SUBSET.md) | XAML grammar: today vs target, POS patterns, migration | Parser / XAML |
| 7 | [VCF_INFRASTRUCTURE.md](./VCF_INFRASTRUCTURE.md) | Collections, views, events, resources, modules | Non-visual core |
| 8 | [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md) | ListView stack, engine, binding, InvoiceGrid target | ListView phase |
| 9 | [VCF_BREAKING_CHANGES_TEMPLATE.md](./VCF_BREAKING_CHANGES_TEMPLATE.md) | Breaking changes log (copy to VCF repo per release) | Release manager |
| 10 | [VCF_MIGRATION_TEMPLATE.md](./VCF_MIGRATION_TEMPLATE.md) | Step-by-step consumer migration guide | Release manager |
| 11 | [VCF_KICKOFF_AGENDA.md](./VCF_KICKOFF_AGENDA.md) | Kickoff meeting — open items, decisions, Phase 1 | VCF lead |

**Denovo context:** [UI_AND_PARTITIONING_BASELINE.md](./UI_AND_PARTITIONING_BASELINE.md) · [pos-v1/docs/DOCUMENTATION_INDEX.md](../../pos-v1/docs/DOCUMENTATION_INDEX.md)

---

## 1. Executive summary

### 1.1 Mission

Evolve **Demac.VCF** from a WPF-*inspired* VB6/Cairo MVVM stack into a **WPF-semantics LOB UI framework** that POS developers can use with minimal surprise — while staying **as light and fast as possible without removing features**.

### 1.2 Consumers

| Consumer | Role |
|----------|------|
| **POS (DeNovo)** | Primary; ~95% of usage; large XAML tree in `pos-v1/UI/Resources/XAML/` |
| **Internal Demac apps** | Small; same DLL patterns |
| **Public / third-party** | Not a goal |

**Implication:** Breaking changes are acceptable when they improve design and WPF parity, **if** each major release ships migration docs. DeNovo will use **AI-assisted bulk migration** in Cursor.

### 1.3 What VCF is today

```text
~98 registered classes · 9 modules · 19 interfaces · ~15,500 LOC (Classes/)
vbRichClient5 / Cairo · AsyncKit merged (BackgroundWorker)
OLE DLL · POS sets IObjectConstructor for app types + res: fragments
```

**Architecture today:**

```text
MyApp (IApplication) → Application.Create → XAMLReader.LoadApp
  → VCF.SetCustomConstructor(POS ObjectConstructor)
  → Run → Cairo message loop

View/Window/UserControl → x:Class shell → LoadSuperclassData
  → visual tree + BindingsManager + Binding (listener-pull on GetValue)
  → DataContext on ViewModel; ICommand via Function delegate
```

### 1.4 What VCF must become

```text
FrameworkElement (shared DP registry, Measure/Arrange, Visibility)
  → Control / ContentControl / Selector / ItemsControl
  → Panels (Grid, StackPanel, UniformGrid compat)
  → BindingExpression (push, detach, DataContext rebind)
  → ResourceDictionary + MergedDictionaries + DynamicResource
  → XamlServices + TypeRegistry + fail-loud XamlLoadException
  → ListView on Selector + ListViewBase engine rewrite
Optional: Demac.VCF.Core (Mail, INIParser, AsyncKit)
```

### 1.5 Non-negotiable outcomes (from POS coordination)

1. **WPF semantic parity** for LOB UI — not full WPF feature parity (no 3D, no FlowDocument, etc.).
2. **Fail-loud XAML** — no silent `Nothing`, no `CallByName`/widget fallback for unknown attributes.
3. **`res:` → ResourceDictionary** — merged dictionaries, `{StaticResource}` / `{DynamicResource}`, built-in resolver.
4. **Unified ListView** — merge `ListView` + `UnboundListView`; rewrite `ListViewBase` for variable row height / hierarchy (InvoiceGrid).
5. **DP-only XAML setters** — layout as DPs (`Width`, `Margin`, …); remove public `Design*`.
6. **BindingExpression** + **DataContext rebind** + **Detach** on navigation.
7. **Lightweight collections** — fix `ListCollectionView` static bug; no `New List` per `Add`.
8. **Release contract** — semver, `BREAKING_CHANGES.md`, `MIGRATION.md`, benchmark baselines.

---

## 2. Inventory at a glance

### 2.1 Scale

| Area | Count | Largest files |
|------|------:|---------------|
| Registered classes | 98 | `TextBoxBase` (1573), `ListViewBase` (1178), `Button` (712), `ListView` (755), `XAMLReader` (581), `TextBlock` (630) |
| Modules | 9 | `modConstructors`, `modStaticClasses`, … |
| Interfaces | 19 | See [VCF_CLASS_REFERENCE.md § Interfaces](./VCF_CLASS_REFERENCE.md) |
| Orphan files (delete) | 5 | `_Image`, `_TextBlock`, duplicate `MarkupExtensions`, stub `IDependencyPropertyCallbackListener`, `API.bas` |
| Built-in styles | 12 | `Styles/*.xml` in VCF repo |
| Test apps | 5+ | `.Tests/Test0`–`Test4`, `SampleApp` |

### 2.2 Rewrite action summary

| Action | Count (approx) | Examples |
|--------|----------------|----------|
| **Refactor** | ~45 | Button, Border, XAMLReader, ObservableCollection |
| **Replace / New** | ~12 types | FrameworkElement, BindingExpression, ResourceDictionary, Selector |
| **Merge** | ~5 | ListView+UnboundListView, DependencyObjectBase→FrameworkElement |
| **Remove** | ~8 | NestedProperty, OverlayWidget, orphans |
| **Keep** | ~25 | Thickness, ICommand, Color, ScrollBar engine |
| **Split (optional)** | ~8 | Mail, INIParser, BackgroundWorker → VCF.Core |

Full table: [VCF_FRAMEWORK_REWRITE_SPEC.md §3 & §19](./VCF_FRAMEWORK_REWRITE_SPEC.md).

---

## 3. Known defects — fix in rewrite (mandatory)

| ID | Severity | Location | Issue | Target fix |
|----|----------|----------|-------|------------|
| **B1** | Critical | `ListCollectionView.Initialize` | **Static** `bIsInitialized` — only first view works | Per-instance init |
| **B2** | High | `ListView` / `ListCollectionView` | `CollectionChangedActionMove` not handled | Implement Move |
| **B3** | High | `ListView` | `IItemsControl_ItemTemplate` stubs | Wire ItemTemplate DP |
| **B4** | High | `ListView` | ItemsSource must be `ObservableCollection`; else silent fail | `IEnumerable` + clear error |
| **B5** | Critical | `XAMLReader` | Silent failures; `On Error Resume Next`; widget/`CallByName` fallback | `XamlLoadException` |
| **B6** | Critical | All bound controls | DataContext change — `' TO-DO: Recreate the Bindings!!!` | BindingExpression rebind |
| **B7** | Low | `ThemesManager`, `Style` | Event typo **`ThemeCkanged`** | Rename `ThemeChanged` |
| **B8** | Medium | `Style` | `BasedOn` edge cases | Test + document |
| **B9** | Medium | Visibility | Hidden treated as Collapsed at widget; Collapsed doesn't affect UniformGrid layout | WPF Visibility DP |
| **B10** | Low | Duplicates | `modAPI` vs `API.bas`; orphan markup files | Delete duplicates |
| **B11** | Medium | `XAMLImagePropertyManager` | `LoadImageFromResource` **empty stub** | Implement or fail loud |
| **B12** | Low | Multiple | Active `Debug.Print` on binding/DP errors | Structured logging or remove |
| **B13** | Medium | `IUserControl.Move` vs `IUIElement.Move` | **ByRef vs by-value** signature mismatch | Unify in rewrite |

---

## 4. Phased program (execution order)

Aligned with [VCF_WPF_ALIGNMENT_NOTES.md §4](./VCF_WPF_ALIGNMENT_NOTES.md) and [VCF_FRAMEWORK_REWRITE_SPEC.md §16](./VCF_FRAMEWORK_REWRITE_SPEC.md).

```text
Phase 0 — Spec, tests, migration contract
  Delete orphans · TypeRegistry · XamlLoadException · golden XAML tests · bench scaffold
  Runner: .Tests/Phase0 · Release docs: docs/BREAKING_CHANGES.md, docs/MIGRATION.md

Phase 1 — Layout core
  DependencyPropertyRegistry · FrameworkElement · Visibility DP · Measure/Arrange
  Remove public Design* (breaking) · optional one-release XAML alias with deprecation

Phase 2 — Panels
  Grid · StackPanel · Border-as-decorator · ContentControl · UniformGrid compat

Phase 3 — Resources & strict XAML
  ResourceDictionary · MergedDictionaries · DynamicResource · built-in res: resolver
  DP-only property setters · delete CallByName fallback

Phase 4 — Bindings & collections
  BindingExpression · DataContext rebind · Detach
  ObservableCollection lightweight args · ListCollectionView fixes · ItemsControl · Selector start

Phase 5 — ListView
  Merge ListView/UnboundListView · ListViewBase rewrite · variable row height · hierarchy

Phase 6 — Templates & polish
  ControlTemplate · Style.Triggers · Button Content DP · render coalescing

Phase 7 — POS migration support (denovo)
  AI migration scripts · pin DLL · integration smoke
```

### 4.1 Acceptance criteria (every phase)

- All **§3 bugs** in scope for that phase are closed or explicitly deferred with ticket.
- **`.Tests`** green; new tests for changed behavior.
- **`BREAKING_CHANGES.md`** updated for any API/XAML surface change.
- **Benchmarks** (layout resize, 1000× collection Add, binding attach/detach) within thresholds — see §6.

---

## 5. Release & migration contract (VCF team owns)

Each **major** or **breaking minor** release MUST include:

| Artifact | Content |
|----------|---------|
| **`BREAKING_CHANGES.md`** | Removed/renamed APIs, XAML attribute renames (`DesignWidth`→`Width`), behavior changes |
| **`MIGRATION.md`** | Step-by-step: XAML transforms, VB6 code patterns, ObjectConstructor changes |
| **`CHANGELOG.md`** | Full change list |
| **Semver tag** | e.g. `v3.0.0-wpf-alignment-p1` |
| **Benchmark report** | vs previous tag; regressions explained |

**DeNovo coordination:** After each tag, denovo pins `Demac.VCF.dll` in `DeNovo.vbp` and runs integration smoke (login → sales → grid → dialog).

---

## 6. Performance & memory expectations

**Observed POS (user telemetry):** Process **< 100 MB** in normal operation; **> 100 MB** only when secondary customer display plays video (decode buffers — not VCF UI).

**Framework targets (not crisis-driven, health/margin):**

| Metric | Current issue | Target (Phase 1–4) |
|--------|---------------|---------------------|
| DP registration | Per-instance Register × N controls | Shared registry — **~10–25 MB** process savings estimated |
| Collection Add | `New List` per notification | Single-item args |
| Binding graph | Binding + NestedProperty + 3× WithEvents | BindingExpression, detach on navigate |
| Layout resize | Design* MoveChild cascade + Widgets.RemoveAll | Measure/Arrange, stable widget tree |
| ListView bind | Template clone + 6 bindings/cell (POS menu grid) | Framework-first; POS VM deferred |

**Phase 0 benchmarks** (add to `.Tests`):

- Golden XAML load (POS Sales subset)
- 1000× `ObservableCollection.Add`
- Two simultaneous `ListCollectionView` instances (B1)
- Window resize nested UniformGrid 50×
- 50× view navigation binding leak test

---

## 7. POS integration points (do not break without migration doc)

| Integration | Location | Rewrite impact |
|-------------|----------|----------------|
| **Custom constructor** | `pos-v1/UI/Source/Classes/ObjectConstructor.cls` | Move `res:` into VCF; shrink Select Case |
| **App bootstrap** | `MyApp.cls` — `LoadXAMLResources`, `SetCustomConstructor` | ResourceDictionary merge |
| **XAML tree** | `pos-v1/UI/Resources/XAML/` | `Design*` → `Width`/`Margin`; `res:` → StaticResource |
| **ViewModels** | `*ViewModel.cls` — INPC | Unaffected if binding rebind fixed |
| **Menu grid hotspot** | `MenuItemsGridButton.xml` — 6 bindings/cell | Deferred POS fix; framework Selector helps |
| **Invoice grid** | Codejock on `FormMain` — out of VCF | Target: VCF ListView Phase 5 |
| **Themes** | `MyApp.xml` styles | `{ThemeResource}` → `{DynamicResource}` |

---

## 8. VB6 constraints & recommended patterns

| Constraint | Rewrite pattern |
|------------|-----------------|
| No implementation inheritance | **Composed `FrameworkElement`** + codegen for `Implements` stubs |
| No true static class fields | **`DependencyPropertyRegistry`** module keyed by `TypeName` |
| Cairo widgets required | **Internal** widget sync from DP **PropertyChangedCallback** only |
| Single-threaded UI | BackgroundWorker stays; UI updates via `RaiseEventAsync` |
| COM visibility | Keep `VB_Exposed` policy; document Friend vs Public |

**Boilerplate reduction:** Generate ~200 lines of `Implements IDependencyObject_*` per control from registry metadata.

### 8.1 OLE DLL compatibility mode (Demac.VCF.vbp)

Use **Project Compatibility** during Phases 0–7 while the public typelib is still evolving.

| VB6 setting | Typelib GUID | Class / member layout | Caller `.vbp` reference | Caller EXE after VCF rebuild |
|-------------|--------------|------------------------|-------------------------|------------------------------|
| **No Compatibility** | New each build | Unrelated | **Lost** — re-add reference in IDE every open | Must recompile |
| **Project Compatibility** *(current)* | **Stable** | May add / change / remove | **Preserved** | **Must recompile** to use new API |
| **Binary Compatibility** | Stable | **Frozen** — same dispids | Preserved | Old EXE may run without recompile |

**Rewrite workflow (Project Compatibility):**

1. Make `Demac.VCF.dll` → `regsvr32 bin\Demac.VCF.dll`
2. Open caller (e.g. `.Tests/Phase0/Phase0.vbp`) — reference should still resolve
3. **Make** the caller EXE (required after any VCF interface change)
4. Run tests

Do **not** switch to No Compatibility during development (reference churn). Do **not** use Binary Compatibility until the public `VCF.*` surface is frozen for a POS pin (post–Phase 7 or hotfix-only releases).

**Typelib warnings** on compile are expected when adding coclasses (`StackPanel`, `Grid`, …) or new `Constructor` members — append new public members at the **end** of existing classes to avoid unnecessary dispatch churn. Warnings about **changed** existing members mean caller recompile is mandatory (same as any Project Compatibility rebuild).

---

## 9. Open questions (resolve at kickoff or defer with ticket)

From [VCF_WPF_ALIGNMENT_NOTES.md §8](./VCF_WPF_ALIGNMENT_NOTES.md):

| Question | Options | Recommendation |
|----------|---------|----------------|
| Grid vs UniformGrid default | Grid primary; UniformGrid compat one release | **Grid primary** after Phase 2 |
| Strict `x:Class` failure | Fail vs fallback to root tag | **Fail loud** (remove dangerous fallback) |
| Listener-pull for non-INPC sources | Keep vs BindingExpression only | **Keep opt-in** for legacy POS objects |
| Utilities DLL split | Monolith vs VCF.Core | **Split in Phase 0** if build allows; else Phase 6 |
| `ObservableCollection.BeginUpdate` | Phase 4 vs later | **Phase 4** with ListView batch updates |
| ViewBase / codegen | Manual vs tool | **Optional ViewBase** in Phase 1; codegen P2 |

---

## 10. Role-based reading paths

### VCF lead / architect

1. This guide §1–§4  
2. [VCF_WPF_ALIGNMENT_NOTES.md §9](./VCF_WPF_ALIGNMENT_NOTES.md)  
3. [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md)  
4. Resolve §9 open questions → update alignment §10  

### Core platform developer (DP, binding, XAML)

1. [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md)  
2. [VCF_INFRASTRUCTURE.md](./VCF_INFRASTRUCTURE.md)  
3. [VCF_XAML_WPF_SUBSET.md](./VCF_XAML_WPF_SUBSET.md)  
4. [VCF_CLASS_REFERENCE.md](./VCF_CLASS_REFERENCE.md) — Core/Binding/XAML sections  

### Controls / layout developer

1. [VCF_FRAMEWORK_REWRITE_SPEC.md §6](./VCF_FRAMEWORK_REWRITE_SPEC.md)  
2. [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md)  
3. [VCF_CLASS_REFERENCE.md](./VCF_CLASS_REFERENCE.md) — Controls section  

### ListView developer

1. [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md)  
2. Alignment §2.14, §6.1 InvoiceGrid  
3. [VCF_CLASS_REFERENCE.md](./VCF_CLASS_REFERENCE.md) — ListViewBase, ListView, UnboundListView  

### QA / release

1. This guide §3 bugs, §4 acceptance, §6 benchmarks  
2. [VCF_FRAMEWORK_REWRITE_SPEC.md §14](./VCF_FRAMEWORK_REWRITE_SPEC.md) test matrix  
3. `.Tests` + POS smoke checklist §7  

---

## 11. Files in this repo (`docs/`)

The handoff package lives under **`docs/`** (not legacy `doc/` CHM help):

- `docs/VCF_TEAM_HANDOFF_GUIDE.md` (this file)
- `docs/VCF_CLASS_REFERENCE.md`
- `docs/VCF_PROPERTY_REGISTRY.md`
- `docs/VCF_XAML_WPF_SUBSET.md`
- `docs/VCF_INFRASTRUCTURE.md`
- `docs/VCF_LISTVIEW_ARCHITECTURE.md`
- `docs/VCF_FRAMEWORK_REWRITE_SPEC.md`
- `docs/MIGRATION.md` (from [VCF_MIGRATION_TEMPLATE.md](./VCF_MIGRATION_TEMPLATE.md))
- `docs/BREAKING_CHANGES.md` (from [VCF_BREAKING_CHANGES_TEMPLATE.md](./VCF_BREAKING_CHANGES_TEMPLATE.md))
- `docs/VCF_PERFORMANCE_BENCHMARKS.md` (Phase 0)

---

## 12. Changelog (this package)

| Date | Change |
|------|--------|
| 2026-06-19 | Initial complete handoff package — 8 documents, full class/property inventory |
| 2026-06-20 | Phase 0 complete — tag v2.0.0-wpf-alignment-p0; baselines recorded; kickoff agenda added |
| 2026-06-20 | Phase 4 validated — tag v2.5.0-wpf-alignment-p4; Phase0 18/18 (P4-BIND, P4-DCTX, P4-DETACH) |
| 2026-06-20 | Phase 4b validated — tag v2.6.0-wpf-alignment-p4b; Phase0 20/20; B-COLL 3 ms |

---

*Questions from VCF team: open issue in VCF repo or coordinate via denovo `VCF_WPF_ALIGNMENT_NOTES.md` §11 discussion log.*
