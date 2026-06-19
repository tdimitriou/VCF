# VCF → WPF alignment — living notes (POS coordination)

**Status:** Draft — discussion in progress  
**Audience:** Demac VCF team (final suggestion TBD) · POS / DeNovo maintainers  
**VCF source repo:** `Demac.VCF` (separate from this monorepo; referenced as `Demac.VCF.dll` from `DeNovo.vbp`)  
**Last updated:** 2026-06-19  

**Purpose:** Capture agreed direction, open questions, and evidence from POS usage while we discuss. When complete, distill this into a **final handoff** to the VCF team.

**Related (existing):**

- **[VCF team handoff guide](./VCF_TEAM_HANDOFF_GUIDE.md)** — **master entry point** for complete rewrite package
- **[VCF framework rewrite spec](./VCF_FRAMEWORK_REWRITE_SPEC.md)** — per-type rewrite blueprint (~98 classes, modules, interfaces, bugs, phases)
- [UI and partitioning baseline](./UI_AND_PARTITIONING_BASELINE.md) — Front/Back Office UI stance, VCF control gaps (ComboBox, ListView bound mode)
- [Application layers](./APPLICATION_LAYERS.md) — MVVM + VCF bindings
- [POS rebuild roadmap](./POS_REBUILD_ROADMAP.md) — Phase E POS.UI/UX (VCF)
- [pos-v1 UI source README](../../pos-v1/UI/Source/README.md) — DeNovo XAML tree

---

## How to use this document

| Section | Use |
|---------|-----|
| **Discussion log** | Dated bullets as we talk (newest first) |
| **Open questions** | Unresolved items — need decision before handoff |
| **Agreed direction** | Stable decisions (move from log when settled) |
| **Final handoff draft** | Polished list for VCF team — fill in last |

---

## Executive summary (current)

VCF is a **WPF-inspired** VB6/Cairo MVVM stack, not a WPF subset today. POS **Front** sales UI is largely on VCF + XAML already; the **shell** (`FormMain`), **Codejock `InvoiceGrid`**, modals, and **Back Office** remain classic VB6.

**Strategic goal (agreed in discussion):** Make VCF **as close to WPF as practical** — semantics, XAML dialect, layout, bindings, controls, **templates**, **selection (`Selector`)** — with **clean design over backward compatibility** where WPF alignment requires it. **Build complex controls by composing simple elements** (WPF-style visual tree), not monolithic per-control Cairo — see §2.11.

**Performance goal (agreed):** Framework **as light and fast as possible** — **without removing or dumbing down any targeted WPF feature** (§2.20). Wins come from **better implementation** (registry, push bindings, Measure/Arrange, coalesced paint), not from a “lite” subset or POS-only shortcuts in the framework.

**Consumer scope:** VCF is used primarily by **POS (DeNovo)** and **a few small internal Demac apps** — not a public or multi-tenant framework. **Breaking changes are acceptable** when they improve design, efficiency, and WPF/XAML parity, **provided the VCF team ships clear documentation and step-by-step migration instructions** with each major release.

**Migration plan (denovo):** After VCF releases + docs, use **AI-assisted migration in Cursor** to bulk-update existing POS XAML and VB6 source (ViewModels, code-behind, `ObjectConstructor`). This lowers the cost of breaking API/XAML changes and favors **one decisive redesign** over long dual-support periods.

**VCF stays in its own repo;** this document lives in **denovo** to coordinate POS requirements and migration.

---

## 1. VCF today (verified from source scan)

**Location on disk (dev):** `Projects\Demac\Framework\Demac.VCF`  
**Shipped:** ~95 classes, OLE DLL, vbRichClient5/Cairo, AsyncKit merged (`BackgroundWorker` in main DLL).

### 1.1 Architecture (short)

```text
MyApp (IApplication) → Application.Create → XAMLReader.LoadApp
  → SetCustomConstructor (POS ObjectConstructor)
  → Run → Cairo message loop

View/Window/UserControl → NewWindow/NewUserControl → LoadSuperclassData
  → Bindings via BindingsManager + Binding
  → DataContext on ViewModel; commands via NewFunction + ICommand
```

### 1.2 Controls in vbp

`Window`, `UserControl`, `Panel`, `Border`, `UniformGrid`, `Button`, `TextBox`, `TextBlock`, `Image`, `ListView`, `UnboundListView`, `ScrollBar`, `WindowsFormsHost`, `Scene`, …

**Missing (vs WPF / POS needs):** ComboBox, TabControl, CheckBox, RadioButton, TreeView, DatePicker, ProgressBar, real **Grid**, ScrollViewer, Menu/ContextMenu, …

### 1.3 Known framework gaps (code-level)

| Issue | Where |
|-------|--------|
| **`DataContext` change does not rebind** | `' TO-DO: Recreate the Bindings!!!` in UserControl, Window, Button, Border, … |
| **Layout: design coords × parent scale** | Window/Border `MoveChild`: `xFactor = actualWidth / DesignWidth` |
| **No Measure/Arrange** | `IUIElement` has commented `ActualWidth/Height`; layout via `Move()` |
| **ListView bound mode** | Slow; `CollectionChangedActionMove` stub; `ItemsSource` must be `ObservableCollection` |
| **`Grid.Row` / `Grid.Column`** | Not implemented; only `Grid.ColumnSpan` / `RowSpan` on **UniformGrid** (child order) |
| **Parser silent failures** | `On Error Resume Next` in parts of XAMLReader |

### 1.4 POS consumption today

| Area | State |
|------|--------|
| XAML | `pos-v1/UI/Resources/XAML/` — large tree, `MyApp.xml` themes/styles |
| Includes | `res:Path\To\Fragment` → app `ObjectConstructor` (**legacy**); target **`MergedDictionaries` + `{StaticResource}`** — §2.16 |
| Sales UI | Mostly VCF views + ViewModels; XAML layouts use **UniformGrid** + many **`DesignLeft/Top/Width/Height`** |
| Invoice lines | Still **Codejock ReportControl** on `FormMain`; stub `OrderItemsView.xml` → `<UnboundListView/>` |
| Shell | `FormMain.frm` hybrid; ~68 `.frm` dialogs, many Codejock |

---

## 2. XAML dialect vs WPF (verified)

**Parser:** `XAMLReader` + vbRichClient `SimpleDOM` — not `System.Xaml`.

| Topic | WPF | VCF today |
|-------|-----|-----------|
| Namespaces | `xmlns`, CLR mapping | Default **`VCF`**; no xmlns handling; custom via `x:Class` + `IObjectConstructor` |
| Root load | Partial class + generated | `x:Class` → construct instance; **`LoadSuperclassData`** for body |
| Markup `{Binding}` | Full grammar | **Subset:** Path, Mode, Converter, StringFormat, Source (nested `{StaticResource}` via literal escape) |
| `{StaticResource}` | Yes | Yes; POS uses `ResourceKey=` form |
| `{ThemeResource}` | N/A | **Deprecated alias** for **`{DynamicResource}`** (§2.17) |
| `{Self}` | RelativeSource Self | **SelfBinding** |
| `{DynamicResource}` | Yes | **Phase 3** — theme merged dict + invalidation (§2.17) |
| RelativeSource / ElementName | Yes | **No** |
| Layout attrs | Width, Height, Margin, Grid.* | **`DesignLeft/Top/Width/Height`**; alignment often numeric |
| Grid panel | Row/Column definitions, `*` | **No Grid**; UniformGrid + span only |
| Triggers, ControlTemplate | Yes | **No** / partial DataTemplate on ListView only |

**POS implication:** Existing XAML is **VCF dialect**, not portable WPF. Alignment = evolve VCF + migrate XAML over time.

---

## 2.5 Dependency properties — evaluation vs WPF

**Sources reviewed:** `DependencyProperty.cls`, `DependencyProperties.cls`, `DependencyPropertiesStatic.cls`, `DependencyPropertyMetadata.cls`, `Binding.cls`, `StyleManager.cls`, `Button.cls` (representative control), `IDependencyPropertyCallbackListener.cls`.

### 2.5.1 WPF model (reference)

| Concept | WPF |
|---------|-----|
| Registration | **`DependencyProperty.Register`** — static identifier per property **per type**; shared metadata |
| Storage | **`DependencyObject`** value store with **precedence stack** |
| Precedence (high → low) | Active animation → local value → template triggers → style triggers → template setters → style setters → default → inherited |
| API | `GetValue` / `SetValue` / **`ClearValue`** / `ReadLocalValue` |
| Bindings | **`BindingExpression`** stored on the property; updates when source or **DataContext** changes |
| Metadata | Default value, **CoerceValue**, **PropertyChanged**, **ValidateValue** callbacks |
| Attached | **`RegisterAttached`** + `GetValue`/`SetValue` on any `DependencyObject` |
| Layout DPs | **`Width`, `Height`, `Margin`, `Alignment`** are dependency properties → drive **Measure/Arrange** |

### 2.5.2 VCF model (today)

| Concept | VCF |
|---------|-----|
| Registration | **Per-instance** `DependencyProperties.Register` in each control’s `Class_Initialize` — **not** static type-level identifiers |
| Storage | Each property = **`DependencyProperty`** object with **`m_Value`** (local / `SetValue`) and **`m_CurrentValue`** (style / inherit / `SetCurrentValue`) + **`m_UnsetValue`** sentinel |
| Effective value | **`GetValue`:** ask **Listeners** (binding pull via `OnValueRequested`) → if local unset use **`m_CurrentValue`** → **`Conversion.TryCast`** |
| Precedence (simplified) | **Local `SetValue`** beats **`SetCurrentValue`** (style/inherit) — **no** trigger/template layers |
| Style | **`StyleManager`** applies setters via **`SetCurrentValue`** (WPF-like intent) |
| Inheritance | **`DependencyPropertiesStatic.InheritPropertyValues`** on parent set; **`IsInheritable`** in metadata (e.g. `DataContext`) |
| Bindings | **`Binding`** implements **`IDependencyPropertyCallbackListener`**; **`AddListener`** on target DP; **`SetValue(ProvideValue)`** on source change; **DataContext** change via **Callback** on `DataContext` DP |
| Metadata | **`IsInheritable`, `AffectsMeasure`, `AffectsRender`, `BindingMode`** — **no** coerce/validate/changed callbacks |
| Attached | Partial: **`AttachedProperties`** dictionary (e.g. `Grid.ColumnSpan` on **`UniformGrid`**) — not full `RegisterAttached` |
| Layout | **`DesignLeft/Top/Width/Height`** are **ordinary fields**, **not** dependency properties — set via `CallByName` from XAML, **not** in binding/style precedence system |

### 2.5.3 What works well (keep)

- **`SetValue` vs `SetCurrentValue`** split matches WPF’s local-vs-style distinction (styles use `SetCurrentValue` — see `StyleManager` comment).
- **`BindingMode` on metadata** + **`Default`** mode in `Binding.GetEffectiveMode`.
- **Inheritance** for `DataContext` and visual properties (`BackColor`, etc.).
- **`AffectsRender`** → `W.Refresh` in several controls (TextBlock, UniformGrid, Image).
- **Listener pull** on `GetValue` documented in `IDependencyPropertyCallbackListener` — clever for sources **without** INPC (legacy objects).
- **Binding ↔ DataContext DP callback** path exists (`OnValueChanged` on `SrcDepObj`) — foundation for context changes.

### 2.5.4 Gaps vs WPF (issues for POS / MVVM)

| Gap | Impact |
|-----|--------|
| **Dual property storage** | Layout (`Design*`) and some CLR props (`CornerRadius`, `GradientBackground`) live **outside** DPs → styles/bindings/layout engine cannot treat them uniformly |
| **No static `DependencyProperty` fields** | No attached-property registry by type; harder WPF API surface (`Button.CommandProperty`); more memory per instance |
| **No full value precedence** | Triggers, template setters, animated values — missing; limits ControlTemplate/triggers roadmap |
| **DataContext change doesn’t rebind** | Controls TODO; binding registry not invalidated when context changes |
| **Binding uses `SetValue` (local)** | Correct for many cases; but no **`BindingExpression`** object — hard to detach, update, or debug |
| **`GetValue` listener pull** | Non-WPF; potential **perf** cost if called frequently; WPF pushes updates via expressions |
| **`AffectsMeasure` unused** | Metadata flag exists but **no `InvalidateMeasure`** pipeline yet |
| **`ClearValue` not public** | WPF uses this to remove local value and fall back to inherited/style |
| **No coerce/validate/changed callbacks** | WPF metadata hooks for validation and side effects |
| **Per-control Register boilerplate** | Every control repeats Register list — error-prone vs WPF static constructor |

### 2.5.5 Suggested improvements (for VCF team — WPF alignment)

**P0 — Fix MVVM correctness**

1. **`BindingExpression` (or equivalent)** — one object per target DP; registered on **`DependencyProperties`** collection; **`Attach` / `Detach` / `UpdateTarget` / `UpdateSource`**.
2. **On `DataContext` change** — walk all **`BindingExpression`** on subtree (or listen at root) and **re-resolve Path**; remove per-control `' TO-DO: Recreate the Bindings!!!`. **Do not push `DataContext` via `PassPropertyValue`** (see §2.8).
3. **Public `ClearValue(property)`** on `DependencyObject` — maps to WPF.

**P1 — Unify property system**

4. Register **`Width`, `Height`, `Margin`, `HorizontalAlignment`, `VerticalAlignment`** as DPs on **`FrameworkElement`-like base** (replace `Design*` fields).
5. **`DependencyProperty.Register` pattern** — static registration helper per class (even if VB6 can’t do true static fields, use **class module singleton registry** keyed by `TypeName`).
6. **Default values in metadata** — replace long `SetCurrentValue` blocks in `Class_Initialize`.
7. **PropertyChangedCallback** in metadata — centralize `W.Refresh`, overlay updates, **`InvalidateMeasure`** (when layout exists).

**P2 — WPF precedence & attached**

8. **Value precedence stack** (minimum): Local → Style → Inherited → Default (document; implement in `GetValue`).
9. **`RegisterAttached`** for `Grid.Row`, `Grid.Column`, `Grid.ColumnSpan`, etc. — replace ad-hoc `AttachedProperties("Grid")` dictionary where possible.
10. **Read-only DPs** — `ActualWidth`, `ActualHeight` (set only by arrange pass).

**P3 — Advanced (later)**

11. **CoerceValueCallback / ValidateValueCallback**.
12. **Trigger/setter precedence** (when ControlTemplate/triggers added).
13. **Replace listener-pull** with push-only for INPC sources; keep pull as opt-in for legacy non-INPC sources only.

### 2.5.6 POS implication

- Today, **only properties registered on `DependencyProperties`** participate in binding/style/inheritance. Layout and many visual knobs **do not**.
- **WPF-aligned layout (§3.2)** depends on **P1 #4–7** — Width/Margin as DPs with **`AffectsMeasure`** driving invalidation.
- **Invoice grid / list perf** less affected; **screen switching and `{Binding}` reliability** depends on **P0**.

### 2.5.7 Open DP questions (see also §8)

- [ ] Adopt **`BindingExpression`** name/API from WPF or VCF-specific name?
- [ ] Keep **listener pull** on `GetValue` for non-INPC legacy POS objects?
- ~~Migrate **`Design*`** to DPs in one breaking release or dual-register?~~ → **One breaking release** (§3.3, §9).

---

## 2.6 Constructor & custom constructor — evaluation vs WPF

**Sources reviewed:** `Constructor.cls`, `modConstructors.bas`, `IObjectConstructor.cls`, `XAMLReader.cls`, `MarkupExtensions.cls`, `XAMLDependencyPropertyManager.cls`, `BindingsManager.cls`, `Application.cls` (`GetStartupObject`), POS `ObjectConstructor.cls`, `MyApp.cls` (`LoadXAMLResources`, `SetCustomConstructor`), sample/test `ObjectConstructor` classes.

### 2.6.1 WPF model (reference)

| Concept | WPF |
|---------|-----|
| Type resolution | **Compile-time** + **`xmlns`** → assembly/CLR type; `XamlTypeMapper` at load |
| `x:Class` | Partial class; **`InitializeComponent()`** generated — loads same XAML/BAML into `this` |
| Resources / includes | **`ResourceDictionary`**, **`MergedDictionaries`**, **`Source=`** URI |
| Markup extensions | Resolved via **`IServiceProvider`** / `ProvideValue` |
| App hook | No global “custom constructor”; types are **project references** |
| Failure mode | Parse/load errors with **line info** (debug) |

### 2.6.2 VCF model (today)

**Two roles merged into `Constructor` / `modConstructors`:**

1. **Internal factory** — `NewWindow`, `NewUserControl`, `NewBinding`, `NewThickness`, … (framework object graph).
2. **App type resolver** — global **`IObjectConstructor`** via `VCF.SetCustomConstructor` → module singleton `modConstructors.CustomConstructor`.

**Object creation flow (XAML tree):**

```text
XAMLReader.NewObject(Node)
  → CreateInstance(Node.Prefix, Node.BaseName)   [modConstructors]
       1. If CustomConstructor set and Namespace=res → CustomConstructor.CreateInstance("res.Path\Fragment")
       2. Else CustomConstructor.CreateInstance(Class)  [only when CC set; see code]
       3. Else CreateObject(Namespace.Class)  e.g. VCF.Button
  → SetObjectProperties / children recurse

XAMLReader.Load(XML) with x:Class on root:
  → NewCustomObject(x:Class) ONLY → returns app view/window shell
  → Rest of document IGNORED until view calls LoadSuperclassData(Me, xamlString)
```

**POS `ObjectConstructor` pattern:**

- Large **`Select Case Classname`** for views, view models, converters (`SalesOrderView`, `LoginViewModel`, …).
- **`Case Else` → `TryCreateObject`:** if classname starts with `res.`, load fragment from **`MyApp.XAMLResources`** via `XAMLReader.Load` (app-specific; not in VCF).
- **`res:Screens\...`** in XAML → prefix `res` + path → `res.Screens\...` key in resource dictionary (backslashes, no `.xml`).

**Other call sites:**

| Caller | Uses CustomConstructor for |
|--------|----------------------------|
| `MarkupExtensions` (unknown `{Extension}`) | Converters / custom markup extensions |
| `XAMLDependencyPropertyManager` | String object property values (e.g. `DataContext="SomeType"`) |
| `BindingsManager` | `{Binding Converter=MyConverter}` via `CreateInstance("", name)` |
| `XAMLStyleReader` | `BasedOn` style type |
| `Application.GetStartupObject` | `StartupURI` short name (e.g. view key without `.xml`) |

### 2.6.3 What works well (keep)

- **`VCF.SetCustomConstructor`** — simple app bootstrap hook; matches VB6 “no assemblies” reality.
- **Two-phase `x:Class` + `LoadSuperclassData`** — analogous to WPF code-behind + generated `InitializeComponent`, once the view loads its XAML key.
- **`CreateInstance(Namespace, Class)`** with default namespace **`VCF`** — reasonable for built-in controls.
- **POS resource dictionary** — preload all XAML into `Application.Resources("XAML")` keyed by relative path; fast lookup for fragments.
- **Internal `NewBinding` / `NewWindow`** — central place for framework wiring (even if API surface should be split).

### 2.6.4 Gaps & risks

| Gap | Impact |
|-----|--------|
| **`Constructor` god object** | Factory helpers + global resolver + `CreateInstance` in one type; hard to document and evolve |
| **`IObjectConstructor` too narrow** | Single `CreateInstance(Classname As String)` — namespace folded into string (`res.Path`); no failure reason, no context |
| **Global singleton** | One `CustomConstructor` per process — OK for POS, but implicit and untyped |
| **Silent failures** | `On Error Resume Next` in `NewObject`, `CreateInstance`, `GetResource` → **Nothing** with no XAML line/element |
| **`x:Class` fallback is dangerous** | If custom create fails, loader may **`CreateObject` root tag** and **ignore `x:Class`** (documented in `XAMLReader` comments) |
| **`res:` is app code, not framework** | Fragment includes only work because POS `TryCreateObject` knows `MyApp.XAMLResources` — every app reimplements |
| **Manual `Select Case` registry** | Every new view/converter requires editing `ObjectConstructor` — does not scale; AI migration helps but framework should reduce boilerplate |
| **Duplicate modules** | `Constructor.cls` duplicates many `modConstructors` functions |
| **Timing / ordering** | App must call `SetCustomConstructor` **before** any XAML that needs app types; easy to get wrong in tests/tools |
| **No WPF `MergedDictionaries`** | `res:` is a parallel include mechanism |
| **View boilerplate** | Each `x:Class` view implements IUserControl/IUIElement/IDependencyObject delegates manually (`SalesOrderView` ~350+ lines) — constructor story incomplete without **ViewBase** |

### 2.6.5 Suggested improvements (VCF team — aligned with breaking-change policy §9)

**P0 — Correctness & debuggability**

1. **Fail loud on create** — replace silent `Nothing` with **`XamlLoadException`** (element name, attribute, resource key, inner error). Optional strict mode for `.Tests` and POS CI.
2. **Remove or restrict `x:Class` fallback** — if `x:Class` present and custom create fails, **error** (do not silently instantiate root tag type).
3. **`IApplication.RegisterTypes`** (or extend `IObjectConstructor`) — called once from `Application.Create` **before** `InitializeComponent`, so apps cannot forget `SetCustomConstructor`.

**P1 — WPF-like type resolution (breaking OK)**

4. **`IXamlTypeResolver`** (name TBD): `Resolve(Namespace As String, TypeName As String) As Object` with built-in precedence:
   - Registered app types (dictionary)
   - **`res:`** → resolve from **`ResourceDictionary`** ( **`DataTemplate` → clone** ); move POS `TryCreateObject` into VCF — §2.16
   - VCF internal registry (`Button`, `ThemeResource`, …)
   - Optional **`CreateObject`** only for explicitly whitelisted prog IDs (dev/test)
5. **`TypeRegistry` class** — `Register "SalesOrderView", SalesOrderView` or **`RegisterConvention "*.View"`, "*.ViewModel", "*.Converter"`** to eliminate giant `Select Case` (POS: one module + AI can regenerate registry).
6. **Split `Constructor`** — public **`VCF`** factory (`NewWindow`, …) vs **`XamlServices`** / resolver (load, create, resources).

**P2 — Resource & x:Class ergonomics**

7. **First-class fragment syntax** — **`ResourceDictionary.MergedDictionaries` + `Source=`** (decided §2.16); **`res:`** transitional shim only, then deprecate in `MIGRATION.md`.
8. **Optional `x:Class` + `x:Key` / `StartupUri` auto-load** — if view declares **`x:ResourceKey="Screens\..."`**, VCF base **`ViewBase`** calls `LoadSuperclassData` automatically (removes copy-paste `InitializeComponent` in every view).
9. **`StartupURI`** — resolve via same resolver as XAML (`res:Key` or type name), not ad hoc string chop.

**P3 — Tooling & migration**

10. **Codegen script** (VCF or denovo) — scan `.cls` for `Implements IUserControl` / `IValueConverter` → generate registry entries for AI/human review.
11. **Document migration:** `ObjectConstructor Select Case` → `TypeRegistry`; `TryCreateObject` deleted when VCF owns `res:`; **`MIGRATION.md`** with before/after XAML include syntax.

### 2.6.6 POS implication

- **Today:** ~20+ explicit cases in `ObjectConstructor.cls` + **`res:`** path in `TryCreateObject` tied to **`MyApp.XAMLResources`** — any VCF change to creation order affects all screens.
- **AI migration angle:** Mechanical transforms are feasible once VCF documents **one** include model and **one** registration API (registry table easier to merge than `Select Case`).
- **Coupling with §2.5:** Converters and `DataContext` types resolved via constructor — must align with **`BindingExpression`** and DataContext rebind.

### 2.6.7 Open constructor questions (see §8)

- ~~**`res:` vs `MergedDictionaries`**~~ — **resolved:** **`MergedDictionaries` + `Source=`**; inline **`res:`** → transitional shim then **`StaticResource`** — §2.16
- [ ] **`ViewBase` in VCF** — framework provides base class vs POS stays with IUserControl boilerplate?
- [ ] **Strict create failures** — default in release builds or debug-only?

---

## 2.7 Rendering & layout performance (nested grids)

**Symptom (POS):** UI feels slow with **deep `UniformGrid` nesting** and many controls (sales screens, menu panels, command grids).

**Hypothesis (product owner):** Re-implementing **render logic** and dropping **`Design*`** will help a lot.

**Verdict:** **Mostly yes** — but **`Design*` removal fixes the layout/resize path**; **render/paint and refresh coalescing** need separate work. All three belong in the WPF alignment program.

### 2.7.1 What happens today (code-level)

**Layout (resize cascade — scales with tree depth × control count):**

| Pattern | Where | Cost |
|---------|--------|------|
| **`xFactor = W.Width / DesignWidth`** on every child | `Window`, `UserControl`, `Panel`, `Border`, `Button` `MoveChild` | O(children) per container **per resize** |
| **`W.Widgets.RemoveAll` then re-add all children** | `UniformGrid.MoveChildren`, `Border.W_Resize` | Rebuilds Cairo widget subtree **every layout pass** |
| **`Child.Move` → `W.Move` → `W_Resize`** | All `IUIElement.Move` | Resize **propagates down** the whole branch |
| **Widget Remove + Add on every `Move`** | `TextBlock.Move`, grid `MoveChild` | Extra dictionary churn per control |
| **UniformGrid on `CollectionChanged`** | `MoveChildren` + `W.Refresh` | Full grid relayout on any child add/remove |

**POS tree shape (example):** `SalesOrderView` → root `UniformGrid` → `res:` columns → inner `Panel` / `Border` stacks → nested `UniformGrid` (8–11 columns of buttons) → `clsMenuPanel` (680×648 design surface). **4+ layout container levels** × **dozens of buttons** × **Design* math each resize**.

**Render (paint — scales with visible controls × paint cost):**

| Pattern | Where | Cost |
|---------|--------|------|
| **Border `W_Paint`** | `Border.cls` | **4× clip + fill + 4× separate rounded-rect strokes** per border per frame |
| **Button gradient** | `Button.DrawBackgroundGradient` | Linear gradient + rounded rect **every paint** (many POS buttons set `GradientBackground="0"` — good) |
| **TextBlock `DrawOn`** | `TextBlock.cls` | `CalcTextRowsInfo` + per-row `GetTextExtents` **every paint** (no layout cache) |
| **`W.Refresh` storms** | `Button` (14 call sites), DP `AffectsRender`, collection changes | **Full widget invalidation**; `Button` children change always refreshes (no `LockRefresh` guard) |
| **`GetScaleFactor` walks to visual root** | `TextBlock` when `ScaleFont=True` | Extra tree walk per text draw (POS often uses `ScaleFont="0"`) |

**UniformGrid vs Design*:** Grids **cell-size in pixels** (`W.Width / Columns`) — good — but **children inside** `Panel`/`Border`/`res:` fragments still use **Design* scaling** (`LeftColumn.xml`: `Panel DesignWidth="512"` with many `DesignLeft/Top/Width/Height` children).

### 2.7.2 Why dropping `Design*` + Measure/Arrange helps

1. **Single arrange in device pixels** — parent passes **final bounds** once; no per-container `actual/design` ratio.
2. **Optional one scale at root** (`Viewbox` / legacy mode) — not repeated at every `Border`/`Panel`.
3. **`AffectsMeasure` → `InvalidateArrange`** — relayout only when needed, not on every property tweak.
4. **Stable visual tree** — arrange **updates positions/sizes** without `Widgets.RemoveAll` on every pass (target behavior).
5. **Canvas-only absolute coords** — WPF model; POS migrates free-floating `DesignLeft/Top` into Grid cells or shared layouts.

**Expected impact:** Large improvement on **first paint after resize**, **screen switches**, and **window resize** — the scenarios where nested grids hurt most today.

### 2.7.3 What Design* removal alone does *not* fix

- **Border/Button Cairo overdraw** — still expensive until render pass is simplified.
- **Uncoalesced `W.Refresh`** — binding/INPC updates can still repaints entire subtrees.
- **Deep nesting** — still many controls; Measure/Arrange reduces **work per arrange**, not **control count** (POS XAML flattening still helps interim).

### 2.7.4 Suggested improvements (VCF team)

**P0 — Quick wins (can ship before full Grid)**

1. **`LockRefresh` during batch layout** — wrap `LoadSuperclassData` / `MoveChildren` tree updates; single refresh at root.
2. **Border paint** — one rounded-rect path + fill + stroke (replace 4× clip + 4× stroke).
3. **Stop `Widgets.RemoveAll` on resize** when child set unchanged — update geometry in place (`UniformGrid`, `Border`).
4. **Text layout cache** — invalidate only when `Text`, `FontSize`, or **arranged width** changes.

**P1 — With Measure/Arrange (same release as Design* removal)**

5. **Arrange pass sets `Actual*`**; panels call `Child.Arrange(rect)` — **delete `MoveChild` Design* math** from `Panel`, `Border`, `UserControl`, `Button`.
6. **`InvalidateArrange` coalescing** — one pass per frame/message pump tick (WPF-style).
7. **Retained visuals** — child widgets stay in `Widgets` collection; update Left/Top/Width/Height only.

**P2 — Render pipeline**

8. **Optional retained layer / bitmap cache** for static subtrees (e.g. menu panel background).
9. **Dirty-region invalidation** — refresh affected widget only, not entire root (if Cairo API allows).
10. **Style default `GradientBackground=False`** for POS-like flat buttons.

**P3 — POS XAML (denovo / AI migration)**

11. **Flatten** where possible: fewer nested `UniformGrid` + `Border` wrappers; merge command grids.
12. **Prefer grid cells** over absolute `DesignLeft/Top` within fragments.

### 2.7.5 POS implication

- Slowness on sales screens is **consistent with measured architecture** — not imagined.
- **Your guess is directionally correct:** Measure/Arrange + dropping Design* is the **main structural fix**.
- Plan **Border/render optimization** and **refresh coalescing** in the same VCF program — otherwise some jank remains after layout migration.
- **Profiling ask for VCF team:** time `MoveChildren` vs `W_Paint` vs `W.Refresh` on a POS golden screen (QuickService2 LeftColumn) before/after changes.

### 2.7.6 Open perf questions (see §8)

- [ ] Ship **P0 render fixes** before Measure/Arrange, or one release?
- [ ] **Bitmap cache** for static panels — worth complexity on POS hardware?

---

## 2.8 Dependency property inheritance — performance review

**Symptom (POS):** UI feels slow on large trees; **inheritance** suspected as a contributor alongside layout/render (§2.7).

**Verdict:** **Valid.** VCF uses **push inheritance** (copy parent → every child on change) instead of WPF **lazy pull** (walk parent chain on `GetValue` only). On sales screens with **hundreds of elements**, **`DataContext` propagation alone** triggers O(n) `SetCurrentValue` + event/callback work per assignment.

**Sources reviewed:** `DependencyPropertiesStatic.cls`, `DependencyProperty.cls` (`SetCurrentValue`, `GetValue`), control `DependencyPropertyChanged` handlers, `Binding.cls` (`SrcDepObj` callback), metadata registration across controls.

### 2.8.1 WPF vs VCF inheritance

| | WPF | VCF today |
|---|-----|-----------|
| Model | **Lazy** — `GetValue` walks **visual parent** while value unset | **Push** — `SetCurrentValue` copied to **each direct child** |
| On parent change | Children read on next `GetValue`; no mass update | **`PassPropertyValue`** → every child updated **immediately** |
| On child attached | Implicit tree link | **`InheritPropertyValues`** — loops **all** registered DPs on child |
| Depth | O(depth) per read, not O(n) per change | O(n) nodes × **event wave** per inheritable change |

### 2.8.2 Current implementation (hot paths)

**`InheritPropertyValues(Target)`** — called when **`Parent` is set** (every control attach):

```text
For Each Prop In Target.DependencyProperties.RegisteredProperties   ' ALL props
    InheritPropertyValue Prop, Target.Parent   ' skip if not IsInheritable
Next
```

- Scans **every registered property** (TextBlock ~15, Button ~10) even though only **`DataContext`** (+ few Button visuals) are inheritable.
- Calls **`Parent.DependencyProperties.GetProperty` → `Source.GetValue`** (runs **binding listeners** on `GetValue`).

**`PassPropertyValue(Children, Source)`** — called from **`DependencyPropertyChanged`** on containers:

| Control | Calls `PassPropertyValue` on every DP change? |
|---------|-----------------------------------------------|
| `Window`, `UserControl`, `Panel`, `Border`, `Button`, `UniformGrid`, `Scene` | **Yes** — all changes (early exit if not `IsInheritable`) |
| `TextBlock`, `Image`, … | **No** — leaf controls do not push to descendants |

**Cascade on `DataContext` change at root:**

```text
Root SetValue/SetCurrentValue DataContext
  → PassPropertyValue(direct children)     [each child SetCurrentValue]
    → each container's DependencyPropertyChanged
      → PassPropertyValue(its children)    [repeat depth times]
        → leaf: SetCurrentValue + Binding SrcDepObj callbacks + INPC hooks
```

**Inheritable properties in practice (metadata `IsInheritable=True`):**

- **`DataContext`** — all major controls (**primary cost** on screen switch / VM swap).
- **`Button`:** `Selected`, `BackColor`, `BorderColor` (also `AffectsRender` — can trigger **`W.Refresh`** on change).
- **`TextBlock` font/visual DPs** — `AffectsRender` but **`IsInheritable=False`** (not pushed; good).

**Interaction with bindings:** `Binding` registers **callback** on target’s **`DataContext` DP** (`AddCallback`). Each pushed `SetCurrentValue` on a bound element runs **`OnDependencyPropertyChanged` → callbacks → `Set Me.Source`** even when the **same object reference** is re-pushed (partially mitigated by `Object.Equals` on `m_CurrentValue`, but **`GetValue` still runs twice** per `SetCurrentValue`).

**Interaction with styles:** `StyleManager` applies setters via **`SetCurrentValue`** → each changed inheritable DP fires **`PassPropertyValue`** again during bulk style apply at load.

### 2.8.3 Why this hurts POS nested grids

- Sales screens: **deep tree** (§2.7) × **many `{Binding}`** elements × **`DataContext` on root view model**.
- **Screen switch** (`SalesOrderView.DataContext = …`) → full-tree push + binding re-init (and TODO rebind not implemented — still pays inheritance cost).
- **Tree build** (XAML load): each **`Parent = …`** → `InheritPropertyValues` × control count × registered prop count.
- **Not O(n²)** on single DataContext change (shallow `PassPropertyValue` per level), but **high constant factor**: VB events, `GetValue`, `Object.Equals`, binding callbacks × **node count**.

### 2.8.4 Suggested optimizations (VCF team)

**P0 — Align with WPF (breaking OK — see §9)**

1. **Lazy inheritance** — remove **`PassPropertyValue`** from default path; **`GetValue`** resolves inheritable DPs by walking **`VisualParent`** until local/binding value found (WPF precedence).
2. **`DataContext` special-case** — store on element; **`BindingExpression`** resolves from **nearest ancestor** with non-null context (no push to descendants).
3. **`InheritPropertyValues`** — maintain **`InheritableProperties` list at `Register`**; do not scan all registered DPs.

**P1 — Until lazy model ships (interim)**

4. **`PassPropertyValue` DFS in one pass** — single recursive walk; **suppress `DependencyPropertyChanged` propagation** during batch (`InheritanceBatch.Begin/End`).
5. **Coalesce `DataContext` updates** — one flag per layout/load; push once at end (or switch to lazy).
6. **Remove `IsInheritable` from `BackColor`/`BorderColor`/`Selected` on `Button`** — use **styles/setters** instead; reduces spurious push + refresh (document in `MIGRATION.md`).

**P2 — With `BindingExpression` (§2.5)**

7. On **`DataContext` change**, **`UpdateTarget` on expressions** in subtree — **not** `SetCurrentValue` on every node’s DP store.
8. **`PropertyChangedCallback` in metadata** — optional **`Inherits`** flag handled centrally, not duplicated in every control’s `DependencyPropertyChanged`.

**P3 — Diagnostics**

9. **Counter/timing** in `.Tests`: `PassPropertyValue` / `InheritPropertyValues` call counts during `LoadSuperclassData` on POS golden XAML.

### 2.8.5 POS implication

- Inheritance slowness is **real and architectural**, not misconfiguration.
- **`DataContext` push** is the dominant cost; fixing **§2.5 P0 (`BindingExpression` + rebind)** and **§2.8 lazy inheritance** should be **one coordinated redesign**.
- Until then, avoid redundant **`DataContext` assigns** in code-behind and prefer **one context at root** (already POS pattern).

### 2.8.6 Open inheritance questions (see §8)

- [ ] **Lazy-only** vs one-release **batched push** shim during migration?
- [ ] Keep **Button `BackColor` inheritance** or style-only?

---

## 2.9 TextBox TwoWay binding & barcode scanner responsiveness

**Symptom (POS):** `TextBox` with `Text="{Binding InputData}"` (TwoWay default on `Text` DP) feels **unresponsive** when a **barcode scanner** (keyboard wedge) sends many characters quickly + Enter.

**Example:** `InputField` in `QuickService/LeftColumn.xml`, `FormMain.txtInput_KeyPress` reads `m_DataContext.InputData` on Enter → `bt_numpad_enter_Click` / barcode lookup.

**Proposed fix (product owner):** **Debounce** source updates — timer (few ms) **reset on each character** so VM/`InputData` updates once after typing pauses.

**Verdict:** **Yes, this approach should fix most of the issue**, if implemented as **deferred UpdateSource** (not deferred display). **Must flush immediately on Enter** (and LostFocus) so barcode submit still sees the full string.

### 2.9.1 Current path (per keystroke)

```text
Scanner char → TextBoxBase.W_KeyPress → InsertText
  → mText updated, CalcRows, caret moved
  → RaiseEvent Change
  → TextBox.m_Base_Change → Me.Text = m_Base.Text
  → DependencyProperties.SetValue("Text")          [local DP]
  → Binding.TargetPropertyChangedEvent (TwoWay)
  → NestedProperty.SetValue → VM.InputData = …
  → VM OnPropertyChanged "InputData"
  → Binding.SetTargetPropertyValue (if value differs)
  → optionally m_Base.Text = … via DP changed
```

**Per character (typical 8–20 char barcode):**

| Step | Cost |
|------|------|
| `InsertText` | **`CalcRows`** (Cairo text metrics) |
| `SetValue` + binding | VM **`OnPropertyChanged`**, nested property **`CallByName`** |
| Round-trip target update | If binding writes back: **`TextBoxBase.Text` Let** → **`mSelStart = 0`**, **`CalcRows`**, **`W.Refresh`** |

**Caret reset (aggravates “unresponsive” feel):** `TextBoxBase.Text` Let clears selection on any text change (`mSelStart = 0`, `mSelLength = 0`, `CalcRows`, `W.Refresh`). If a **stale** value round-trips from the VM during the burst, the caret jumps to the start and wedge input lands in the wrong place.

**VCF has no `UpdateSourceTrigger` / delay** — TwoWay pushes to source on **every** target change (`Binding.cls` `TargetPropertyChangedEvent_PropertyChanged`).

**Text DP metadata:** default **`BindingMode.TwoWay`** on register (`TextBox.cls`).

### 2.9.2 Will debounce fix it?

| If debounce… | Effect |
|--------------|--------|
| **Delays UpdateSource only** (local `mText` still updates each char) | **Yes** — UI stays responsive; VM updated once after pause |
| **Delays InsertText / display** | **No** — wrong layer |
| **No flush on Enter** | **Breaks barcode** — Enter handler reads `InputData` before timer fires |
| **+ suppress target write-back while user is editing** | **Yes** — avoids caret reset / double `CalcRows` |

**Barcode timing:** wedge sends chars + Enter in one rapid burst (often under 10 ms between keys). A 20–50 ms debounce is fine **if Enter calls `FlushUpdateSource`** before `FormMain` reads `InputData` (`FormMain.frm` Enter handler ~line 1686).

### 2.9.3 Recommended implementation (VCF team)

**P0 — Binding (preferred)**

1. **`UpdateSourceDelay`** (ms) on `Binding` — `0` = today; `>0` = debounce **Target → Source** only.
2. **`FlushUpdateSource()`** — push pending value immediately.
3. **TextBox** / **TextBoxBase:** on **Enter** and **LostFocus**, flush all TwoWay bindings on `Text` before bubbling to app.
4. **`IsUpdatingSource` / `IsUpdatingTarget` flags** — skip round-trip when change originated from the other side.

**P1 — TextBox hardening**

5. **`Text` Let:** preserve caret when new value is a **prefix extension** at caret (defensive).
6. Optional: **`m_Base_Change`** path that schedules source update via binding when delay &gt; 0.

**P2 — Defaults**

7. Metadata default for **`Text`:** `UpdateSourceDelay=50` (breaking — document in `MIGRATION.md`).
8. XAML: `{Binding InputData, UpdateSourceDelay=50}` when parser supports extra binding props.

**Interim POS workaround:** On Enter in `txtInput_KeyPress`, read **`txtInput.Text`** instead of **`InputData`** — fixes submit only, not mid-scan jank.

### 2.9.4 Suggested delay values

| Scenario | Delay |
|----------|-------|
| Barcode / scanner wedge | **30–80 ms** + **flush on Enter** |
| Hand typing | **100–300 ms** or LostFocus-only |

### 2.9.5 Open questions (see §8)

- [ ] Default **`UpdateSourceDelay`** for all Text TwoWay vs opt-in only?
- [ ] **`LostFocus` flush** — always, or only when delay &gt; 0?

---

## 2.10 Button content / caption (string without nested TextBlock)

**Symptom (POS):** Every labeled button requires a **nested `TextBlock`** to show text — verbose XAML, extra child widget, extra `MoveChild` / layout work (§2.7).

**Example today:**

```xml
<Button Command="..." GradientBackground="0">
  <TextBlock HorizontalAlignment="2" Text="REPEAT" ForeColor="&HFFFFFF" FontName="Segoe UI" FontSize="12" FontBold="1" ScaleFont="0"/>
</Button>
```

**Desired (WPF-like):**

```xml
<Button Content="REPEAT" Command="..." />
<!-- or -->
<Button Content="{Binding Caption}" Command="..." />
```

**Verdict:** **Correct direction.** VCF `Button` today is a **chrome-only** control (background, border, overlay); **text is never drawn on the button itself** — only **child** elements (usually `TextBlock` or `Image` + `TextBlock`) are positioned via `MoveChild` and rendered as **separate widgets**.

### 2.10.1 WPF model (reference)

- `Button` inherits **`ContentControl`**.
- **`Content`** property (`Object`) — string, `TextBlock`, `Image`, panel, etc.
- **`Content="Save"`** or inner text → framework shows caption via **ContentPresenter** (default template creates text display for strings).
- **`Content="{Binding ...}"`** supported.
- Separate: **`ContentTemplate`** for complex content (later for VCF).

### 2.10.2 VCF Button today

| Aspect | VCF |
|--------|-----|
| **`Content` / `Text` DP** | **None** on `Button` |
| **`W_Paint`** | Draws background, border, overlay only — **no text** |
| **Caption** | **Child `TextBlock`** in `m_Children`, scaled via `MoveChild` (Design* math) |
| **Font on button** | Base style sets `FontName`, `ForeColor`, … on **`W` (Widget)** — **not used for caption text** |
| **Designer** | Property editor lists `Content`, `Caption` as **future** names — not implemented |

**POS impact:** Hundreds of command fragments (`Commands/*.xml`, grid buttons) repeat the same **Button + TextBlock** pattern — good **AI migration** target once `Content` exists.

### 2.10.3 Recommended design (VCF team)

**P0 — String content (covers most POS buttons)**

1. Add **`Content` dependency property** on `Button` (`Object` / string; bindable).
2. When **`Content` is a string** (or resolves to string from binding): **draw caption in `W_Paint`** using existing widget font properties (`W.FontName`, `W.FontSize`, `W.ForeColor`, `W.FontBold`, … from style) — **no extra child widget** (better perf than nested `TextBlock`).
3. Add **`HorizontalContentAlignment` / `VerticalContentAlignment`** DPs (default **Center**) — maps to today’s `TextBlock HorizontalAlignment="2"`.
4. **XAML:** `Content="REPEAT"` and `Content="{Binding Caption}"` via existing markup / DP pipeline.
5. **Precedence:** If **`Content` is set** (non-empty string), **do not require** child `TextBlock`; if **both** `Content` and visual children exist, **document:** children win (WPF: explicit content replaces template) or **`Content` wins** — pick one, document in `MIGRATION.md`.

**P1 — WPF parity & mixed content**

6. **`Content` as object** — when set to **`Image`** or **`TextBlock`**, treat as today’s child model (single content child, auto `MoveChild` / widget host).
7. Optional **`ContentControl`** base extracted from `Button` for reuse (`UserControl` host patterns later).
8. **XAML shorthand:** single text-only element child could map to `Content` at load (optional reader sugar for migration).

**P2 — Later**

9. **`ContentTemplate`** / **`ContentTemplateSelector`** — not needed for POS v1.

### 2.10.4 Migration (denovo / AI)

| Before | After |
|--------|-------|
| `<Button …><TextBlock Text="X" HorizontalAlignment="2" FontSize="12" …/></Button>` | `<Button Content="X" FontSize="12" …/>` (font attrs on Button — already in style setters) |
| `Text="{Binding …}"` on inner TextBlock | `Content="{Binding …}"` on Button |
| Button with **Image + TextBlock** | Keep explicit children **or** future `Content` as panel |

Mechanical transform for **simple** buttons (single TextBlock, text only); manual review for image+text combos.

### 2.10.5 Open questions (see §8)

- [ ] **`Content` vs legacy `Text` alias** on Button — support both names during migration?
- [ ] Image + string — require children until template support?

---

## 2.11 Compositional architecture (simple elements → complex controls)

**Question:** WPF builds complex controls from **simple visual elements** composed in a tree. Can VCF follow the same logic?

**Verdict:** **Yes — and it should**, as the **structural theme** of the WPF alignment program. VCF already has **fragments** of composition; the gap is **inconsistent models** (monolithic `W_Paint` vs child widgets vs `IVisualChild`) and **no shared element base** or **ControlTemplate** story.

### 2.11.1 WPF composition model (reference)

```text
Visual primitives     TextBlock, Border, Rectangle, Image
        ↓
Panels                Grid, StackPanel, DockPanel — arrange children
        ↓
Content controls      ContentControl / Button — single Content + ContentPresenter
        ↓
Items controls        ItemsControl / ListView — ItemsSource + DataTemplate
        ↓
Lookless controls     Control.Template — entire visual tree replaced in XAML
```

One **layout/render pipeline** walks the tree. **Styles** set properties; **ControlTemplates** define structure.

### 2.11.2 VCF today — mixed models

| Layer | VCF state |
|-------|-----------|
| **Child collection** | **`UIElementCollection`** on Panel, Border, Button, Window, … — compositional |
| **Panels** | UniformGrid, Panel, Border — host children (layout via Design* / MoveChild — §2.7) |
| **DataTemplate** | **Exists** — `DataTemplate.Children` + XAML; **ListView** builds item UI from template trees — **best existing compositional pattern** |
| **Style** | **Property setters only** — not a visual tree (no ControlTemplate) |
| **Button** | **Monolithic** Cairo in `W_Paint` + **`OverlayWidget`** hack + manual child `MoveChild`; caption = nested TextBlock (§2.10) |
| **Border** | **Monolithic** multi-clip `W_Paint` — not a Border wrapping a child Border element |
| **TextBlock / Image** | Can be **standalone widgets** or **`IVisualChild.DrawOn`** — two render paths |
| **Element base** | **`UIElementBase`** = resources/attached props only; **no** shared `FrameworkElement` with layout DPs, Measure/Arrange |
| **Per-control boilerplate** | Each control **re-implements** `IControl` + `IUIElement` + `IDependencyObject` (~300+ lines in POS views) |

**Summary:** VCF **composes screens** from Border/Grid/Button in XAML, but **complex controls are mostly hand-drawn**, not built from smaller **lookless** pieces. ListView **DataTemplate** proves the approach works in VB6.

### 2.11.3 Target VCF composition stack (recommended)

Align with phased roadmap (§4); **breaking changes OK** (§9).

**Layer 0 — Element core (with Phase 1 layout)**

- **`FrameworkElement`** (name TBD): shared DPs (`Width`, `Height`, `Margin`, `Visibility`, …), **`Measure` / `Arrange`**, **`VisualParent`**, single invalidation path.
- **`UIElement`** leaf contract: participate in tree walk (render or delegate to Cairo widget).

**Layer 1 — Visual primitives (reuse/refactor, don’t duplicate)**

| Primitive | Role |
|-----------|------|
| **TextBlock** | Text layout + draw (one code path) |
| **Border** | Background + border + **single child** (child fills interior — WPF model) |
| **Image** | Surface draw |
| **Rectangle / Shape** | Optional thin fills for templates |

**Layer 2 — Panels**

- **Grid, StackPanel, DockPanel, Canvas** — **only** arrange children; **no** custom Cairo chrome in panel `W_Paint` except optional background.
- UniformGrid behavior folded into **Grid** or kept as simplified panel.

**Layer 3 — Content & items**

- **`ContentControl`**: `Content` + **`ContentPresenter`** (internal) — hosts string (auto text), or one visual child.
- **`Button` : ContentControl** — drop monolithic chrome over time; default template = Border + presenter (Phase 6) or interim styled Border child in code.
- **`ItemsControl` / ListView`**: generalize existing **DataTemplate** + `ItemsPresenter` pattern (already partially there).

**Layer 4 — Templates (Phase 6)**

- **`ControlTemplate`** for Button, TextBox chrome, ScrollViewer, …
- **`DataTemplate`** (existing) — keep; unify API with items/content templates.
- Default templates **in XAML** in VCF repo (WPF `generic.xaml` equivalent).

### 2.11.4 Phased adoption (avoid big-bang rewrite)

| Phase | Composition win |
|-------|-----------------|
| **Now → P1** | `FrameworkElement` + Measure/Arrange; **Border** = decorator with one child; **Button.Content** string (§2.10) |
| **P2** | **ContentControl** base; Button inherits; Grid/StackPanel |
| **P3–P4** | ItemsControl pattern; ListView on shared template engine |
| **P6** | **ControlTemplate** — Button default visual tree in XAML; POS restyle via templates not forked controls |

**Do not** rewrite TextBoxBase in one release — wrap incrementally (chrome via template, editor core stays).

### 2.11.5 POS / performance implications

- **Fewer ad hoc widgets** (OverlayWidget, Button+TextBlock pairs) → shallower Cairo trees (§2.7).
- **One layout pass** instead of per-control Design* `MoveChild`.
- **AI migration:** compositional XAML closer to WPF → easier transforms; Button `Content`, Grid layouts, templates.

### 2.11.6 What not to copy literally from WPF

- Full **visual/logical tree split** on day one — start with **one tree** + optional template expansion.
- **RetainedDirect3D**-style retained mode — Cairo immediate mode is fine; cache static subtrees if needed (§2.7).
- **Every** WPF panel/control — POS LOB subset first.

### 2.11.7 Open composition questions (see §8)

- [ ] **`FrameworkElement` in VB6** — one base class module + thin wrappers vs codegen for interfaces?
- [ ] **Button first step:** `Content` string only (§2.10) **before** full ContentControl/template?
- [ ] **Border refactor** — breaking visual change acceptable in same release as layout engine?

---

## 2.12 Template mechanisms — VCF `DataTemplate` vs POS `@`-fragments (MessageBox)

**Question:** VCF uses a form of XAML templates; MessageBox is an example. Evaluate.

**Verdict:** POS MessageBox uses **two different mechanisms** — only one is framework-native. Both support §2.11, but should **not be conflated** in the VCF handoff.

### 2.12.1 MessageBox anatomy (POS)

| Piece | File / mechanism | Role |
|-------|------------------|------|
| **View shell** | `Screens\MessageBox\MessageBoxView.xml` | Full `Window` via `LoadSuperclassData` — title, message `{Binding}`, icon slot, empty `Buttons` grid |
| **Button fragment** | `Templates\MessageBoxButton.xml` | **Not** a VCF `DataTemplate` — snippet with `@Text`, `@CommandParameter`, … |
| **Instantiation** | `MessageBox.cls` → `GetButtonXML` + `Replace$` + `XAMLReader.Load` | One `Button` per `MessageBoxButton` def; added to `UniformGrid` at runtime |
| **Resources** | `MyApp.XAMLResources` | All `Resources\XAML\**\*.xml` keyed by relative path |

**Same POS pattern:** `DialogWindow.cls` + `Templates\DialogWindow.xml` and `Templates\DialogButton.xml`.

**MessageBoxView.xml** is compositional XAML. **MessageBoxButton.xml** is a **string-substitution micro-template** — closer to mail-merge than WPF `DataTemplate`.

### 2.12.2 VCF native `DataTemplate` (framework)

| Aspect | Behavior |
|--------|----------|
| **Class** | `DataTemplate` — `Children` collection, optional `DataType` |
| **XAML** | `ListView.Resources` / `ListView.ItemTemplate` — parsed via `ListView.*` dot-properties in `XAMLReader` |
| **Selection** | Per item: `TryFindResource("DataTemplate_" & TypeName(Item))` else `ItemTemplate` |
| **Per row** | `ICloneable.Clone` each child; `DataContext = Item`; cache in `ItemTemplates` |
| **Render** | Owner-draw: `IVisualChild.DrawOn` — **not** full widget subtree per row |

**Scope today:** **ListView items only.** No `ControlTemplate`, no general `ItemsControl`, no framework `@`-fragment support.

### 2.12.3 Comparison

| | VCF `DataTemplate` | POS `@`-fragment |
|--|-------------------|------------------|
| **WPF equivalent** | `DataTemplate` (partial) | Ad hoc |
| **Parameterization** | `{Binding}` on item `DataContext` | `Replace$("@Text", …)` before parse |
| **Reuse** | Resource dict + `DataType` | Per-screen `GetButtonXML` helper |
| **Parse cost** | Clone once per row (cached) | Re-parse XML per instance |
| **Framework** | Yes | App convention only |

### 2.12.4 What works well

- Chrome vs repeated cell separated (shell XAML + button fragment).
- Proves XAML fragments + `XAMLReader.Load` work for dynamic UI.
- VCF `DataTemplate` clone + per-item `DataContext` is the right list-row model.

### 2.12.5 Gaps and risks

- **`@` substitution fragile** — no XML escaping; duplicated helpers (MessageBox vs DialogWindow).
- **No binding in fragments** — countdown timer mutates `Btn.Children(0).Text` instead of VM.
- **`@Style` in `GetButtonXML` unused** in `MessageBoxButton.xml`.
- **DataTemplate cannot template** a dynamic `UniformGrid` of buttons today.

### 2.12.6 Recommendations

| Priority | Action |
|----------|--------|
| **P1** | **`ItemsControl` + `DataTemplate`** for MessageBox/dialog buttons — `ItemsSource="{Binding DialogButtons}"` |
| **P1** | Prefer **binding-only** templates over formalizing `@` syntax |
| **P2** | **`ControlTemplate`** for Button; built-in **`res:`** loader (§2.6) |

**Short term:** MessageBox pattern is acceptable legacy; not a framework template API.

### 2.12.7 Relation to §2.11

Extend **ListView `DataTemplate`** to general **ItemsControl**; absorb MessageBox `@`-fragments into that — do not add `@` to VCF long-term.

### 2.12.8 WPF-aligned template target (agreed direction)

**Decision:** Template support in VCF should follow **WPF semantics and XAML shape** where VB6/Cairo constraints allow. **Do not** formalize POS `@`-fragment substitution in the framework. POS `@` templates are **legacy**; migrate to WPF-style **`DataTemplate` + `{Binding}`** when `ItemsControl` ships.

#### WPF template stack (reference)

```text
FrameworkTemplate (conceptual base)
├── DataTemplate          — visual tree for a *data item* (DataType / x:Key)
├── ControlTemplate       — visual tree for a *control's chrome* (TargetType)
├── ItemsPanelTemplate    — panel that lays out generated items
└── (HierarchicalDataTemplate — out of scope v1)

ItemsControl
├── ItemsSource
├── ItemTemplate          — DataTemplate (default)
├── ItemTemplateSelector  — optional; defer v1 unless InvoiceGrid needs it
├── ItemsPanel            — ItemsPanelTemplate (default: StackPanel)
└── ItemContainerStyle    — optional; defer unless ListViewItem-like chrome needed

ContentControl
├── Content
└── ContentTemplate       — DataTemplate for Content (Phase 2b)

Control (Button, …)
├── Template              — ControlTemplate (Phase 6)
└── default template in generic resources (WPF generic.xaml equivalent)
```

#### VCF today → WPF target mapping

| WPF | VCF today | VCF target |
|-----|-----------|------------|
| `DataTemplate` | `DataTemplate` class; ListView only; owner-draw | Same class; **shared engine** for ListView + ItemsControl; typed lookup via `DataType` / `x:Key` |
| `ItemsControl` | `IItemsControl` interface; **only ListView implements** | **`ItemsControl` control** — generates **visual/logical children** (or panel slots), not `@` loops |
| `ItemsPanelTemplate` | Manual `UniformGrid.Children.AddRange` | `<ItemsControl.ItemsPanel><ItemsPanelTemplate><UniformGrid …/></ItemsPanelTemplate>` |
| `ControlTemplate` | None | Phase 6; `TargetType="Button"` etc. |
| `ContentTemplate` | None | Phase 2b on `ContentControl` |
| Implicit template (`DataType`) | `DataTemplate_<TypeName>` resource key hack | Keep key convention **or** document as WPF-compatible implicit lookup in `ResourceDictionary` |
| Template inflation | Clone + `DataContext = item` (ListView) | Same: **parse once**, **clone per instance**, bind with item as `DataContext` |
| POS `@`-fragments | MessageBox, DialogWindow | **Remove** — replace with resources + bindings (see below) |

#### Implementation notes (VCF team)

1. **Unify template engine** — extract from `ListView.CreateDataTemplate` into shared module used by `ItemsControl` and (later) `ContentControl.ContentTemplate`.
2. **Two render modes (pragmatic):**
   - **`ItemsControl`** — full element tree per item (MessageBox buttons, toolbars) — WPF-default behavior.
   - **`ListView`** — may keep **owner-draw** path for large lists (perf); same `DataTemplate` definition, different presenter — document as intentional deviation unless/until virtualization exists.
3. **ResourceDictionary** — templates live in `Window.Resources` / `Application.Resources` / merged dicts (Phase 3), referenced as `{StaticResource MessageBoxButtonTemplate}` — not flat `XAMLResources` string keys in app code long-term (framework `res:` can load files into dict).
4. **No `@` token API** — if a one-release shim maps `@Text` → binding for AI migration, treat as **deprecated** in `MIGRATION.md` only.

#### MessageBox — WPF-aligned target XAML (migration example)

**Resources** (in `MessageBoxView.xml` or merged dictionary):

```xml
<DataTemplate x:Key="MessageBoxButtonTemplate" DataType="MessageBoxButton">
  <Button Grid.ColumnSpan="2"
          CommandParameter="{Binding Value}"
          Content="{Binding Text}"
          BackColor="{Binding BackColor}"
          IsDefault="{Binding IsDefault}"
          IsCancel="{Binding IsCancel}"/>
</DataTemplate>
```

**View** (replace empty `Buttons` grid + `SetButtons` loop):

```xml
<ItemsControl ItemsSource="{Binding DialogButtons}"
              ItemTemplate="{StaticResource MessageBoxButtonTemplate}">
  <ItemsControl.ItemsPanel>
    <ItemsPanelTemplate>
      <UniformGrid Rows="1" Columns="8" Padding="8"/>
    </ItemsPanelTemplate>
  </ItemsControl.ItemsPanel>
</ItemsControl>
```

**Code-behind changes:**

- Delete `GetButtonXML`, `Replace$`, per-button `XAMLReader.Load` loop.
- Set `Button.Command` via **style setter**, **ItemContainerStyle**, or `{Binding}` to command on view / window `DataContext` once **RelativeSource** or **ElementName** exists (Phase 4 minimal); interim: attach command in `ItemsControl` item-generated callback if needed — **not** `@CommandParameter` substitution.
- Countdown: bind display string on `MessageBoxButton` VM (`Text` + `DelaySuffix`) — no `Btn.Children(0).Text` mutation.

**DialogWindow** — same pattern: one `DataTemplate` for dialog buttons, `ItemsSource` from DB-driven VM collection.

#### Phasing (ties to §4)

| Phase | Template deliverable |
|-------|---------------------|
| **3** | Templates in `ResourceDictionary`; `{StaticResource}` |
| **4** | **`ItemsControl`** + **`ItemsPanelTemplate`**; unify `DataTemplate` engine; ListView uses same parser |
| **2b** | `ContentControl` + `ContentTemplate` |
| **6** | **`ControlTemplate`** + default Button template (Border + ContentPresenter) |
| **7 / POS** | Migrate MessageBox, DialogWindow, `@` templates → bindings; delete `GetButtonXML` helpers |

#### Out of scope (templates v1)

- `HierarchicalDataTemplate`, `DataTemplateSelector` — **InvoiceGrid may need hierarchical rows in Phase 5** (§6); else flat list + row level in engine
- `ControlTemplate` triggers beyond PropertyTrigger subset
- Compiled/BAML templates; visual tree serialization identical to WPF internal format

### 2.12.9 Open questions (see §8)

- [ ] **ListView owner-draw vs full tree** — keep dual presenter permanently or converge when perf allows?
- [ ] **Command binding in DataTemplate** — require minimal RelativeSource in Phase 4, or temporary ItemsControl hook?

---

## 2.13 Selector support (WPF-aligned selection in XAML)

**Request:** Add **Selector** to VCF XAML capabilities — WPF-aligned selection API for list-like controls.

**Decision:** Introduce a **`Selector`** base (conceptual; VB6 class hierarchy) extending **`ItemsControl`**, with **bindable selection DPs** exposed in XAML. Refactor **`ListView`** onto this stack; future **`ListBox`**, **`ComboBox`**, **`TabControl`** inherit the same contract.

### 2.13.1 WPF Selector model (reference)

```text
ItemsControl
    └── Selector                    — adds selection state
            ├── ListBox             — single/multiple selection
            ├── ComboBox            — editable dropdown + selection
            ├── ListView            — ListBox + optional grid view (VCF: one control)
            └── TabControl            — selected tab as selection
```

**Core selection API (WPF — target for VCF XAML):**

| Member | Role |
|--------|------|
| **`SelectedItem`** | The selected object from `ItemsSource` (TwoWay bindable) |
| **`SelectedIndex`** | Index in view (-1 = none) |
| **`SelectedValue`** | Property of selected item named by `SelectedValuePath` |
| **`SelectedValuePath`** | e.g. `"Value"`, `"ID"` — path on item type |
| **`IsSynchronizedWithCurrentItem`** | Sync with `CollectionView.CurrentItem` (default true where applicable) |
| **`SelectionChanged`** | Routed event / callback when selection moves |

**Not in v1 (defer):** `Selector.SelectionMode` (Multiple/Extended), `SelectedItems` collection — unless InvoiceGrid needs multi-select early.

### 2.13.2 VCF today — partial, non-WPF surface

| WPF | VCF today | Gap |
|-----|-----------|-----|
| `SelectedIndex` | **`ListViewBase.ListIndex`** (Long, not a DP on ListView XAML) | Wrong name; not bindable from XAML on bound ListView |
| `SelectedItem` | **`ListCollectionView.CurrentItem`** (internal sync) | Not exposed as ListView DP; no `{Binding SelectedItem}` |
| `SelectedValue` / `SelectedValuePath` | **DialogWindow:** `@Selected` string compare in `GetButtonXML` | Ad hoc, not framework |
| `SelectionChanged` | **`ListIndexChanged`** event on ListViewBase | Not WPF name; not a DP-driven pattern |
| `Selector` type | **None** | No shared base for ComboBox/TabControl later |
| Button toggle | **`Button.Selected`** DP | Visual state only — not container selection |

**Existing foundation (reuse, don't rewrite):**

- `ListCollectionView` — `CurrentItem`, `CurrentPosition`, `MoveCurrentTo*` (partial **`ICollectionView`**).
- `ListView` already syncs `ListIndex` ↔ `ListCollectionView.CurrentPosition` in code.
- `CollectionViewSource.GetDefaultView(ObservableCollection)` — default view per collection.

**Work:** Expose WPF names as **dependency properties** on **`Selector`**, wire to existing `ListCollectionView`, deprecate public **`ListIndex`** on bound controls (breaking; document in `MIGRATION.md`).

### 2.13.3 VCF target — class stack & XAML

```text
ItemsControl          — §2.12.8 (no selection)
Selector              — SelectedItem, SelectedIndex, SelectedValue, SelectedValuePath
├── ListBox           — optional alias / simpler default panel (Phase 5)
├── ListView          — refactor: Selector + owner-draw presenter (POS primary)
├── ComboBox          — Phase 5
└── TabControl        — Phase 5
```

**Dependency properties (register on Selector, inherit):**

```text
SelectedItem          Object      TwoWay default
SelectedIndex         Long        TwoWay; -1 when empty
SelectedValue         Variant     TwoWay
SelectedValuePath     String
IsSynchronizedWithCurrentItem  Boolean  default True
```

**Behavior (match WPF semantics):**

1. Setting **`SelectedIndex`** updates **`SelectedItem`**, **`SelectedValue`**, and **`ListCollectionView`** current position.
2. Setting **`SelectedItem`** (or **`MoveCurrentTo`**) updates index and value; scrolls into view on ListView.
3. **`SelectedValuePath`** — simple property name on item (v1); no full property-path parser until needed.
4. **`ItemsSource`** change clears selection when current item absent; re-sync if `IsSynchronizedWithCurrentItem`.
5. **`SelectionChanged`** — raise after batch updates; bindings use same precedence as other DPs (§2.5).

### 2.13.4 XAML examples (target)

**Login user list** (replace code setting `ListIndex`):

```xml
<ListView Name="MyList"
          ItemsSource="{Binding UserList}"
          SelectedItem="{Binding SelectedUser, Mode=TwoWay}"
          ItemTemplate="{StaticResource UserRowTemplate}"/>
```

**Dialog / pick-one grid** (replaces DialogWindow `@Selected` + `SelectedValue` param):

```xml
<ListBox ItemsSource="{Binding DialogButtons}"
         SelectedValue="{Binding Result, Mode=TwoWay}"
         SelectedValuePath="Value"
         ItemTemplate="{StaticResource DialogButtonTemplate}">
  <ListBox.ItemsPanel>
    <ItemsPanelTemplate>
      <UniformGrid Rows="3" Columns="4" Padding="8"/>
    </ItemsPanelTemplate>
  </ListBox.ItemsPanel>
</ListBox>
```

**TabControl (future):**

```xml
<TabControl SelectedItem="{Binding ActiveTab, Mode=TwoWay}"
            ItemsSource="{Binding Tabs}"/>
```

### 2.13.5 Item containers & visual selection

WPF uses **`ListBoxItem`** / **`ListViewItem`** with **`IsSelected`**. VCF v1 options:

| Approach | When |
|----------|------|
| **Owner-draw highlight** (ListView today) | Keep for perf — `SelectedItem` drives row chrome in `DrawOn` |
| **`ItemContainerStyle`** + triggers | Phase 6 with `ControlTemplate` / styles — selected button border |
| **`Button.IsSelected` attached/state** | Avoid for list selection — use **`Selector`**, not per-button `@Selected` |

**DialogWindow migration:** delete `Selected="@Selected"` token; selection comes from **`ListBox.SelectedValue`** or **`SelectedItem`** binding.

### 2.13.6 Implementation phasing

| Phase | Selector deliverable |
|-------|---------------------|
| **4** | **`Selector` base** + **`ItemsControl`**; **`ListView : Selector`** — expose `SelectedItem` / `SelectedIndex` / `SelectedValue` / `SelectedValuePath` in XAML; sync with `ListCollectionView` |
| **4** | **`SelectionChanged`** event; **`MIGRATION.md`:** `ListIndex` → `SelectedIndex` |
| **5** | **`ListBox`**, **`ComboBox`**, **`TabControl`** on `Selector` |
| **6** | **`ItemContainerStyle`**, selected triggers in templates |
| **7 / POS** | LoginView, DialogWindow, any `ListIndex` code-behind → bindings |

### 2.13.7 POS impact

| Screen / pattern | Today | WPF-aligned |
|------------------|-------|-------------|
| **LoginView** | `ItemsSource` only; selection in code? | `SelectedItem="{Binding SelectedUser}"` |
| **DialogWindow** | `@Selected` vs `SelectedValue` string | `ListBox` + `SelectedValue` + `SelectedValuePath="Value"` |
| **Invoice grid (future)** | Codejock read-only hierarchical columns | Merged **`ListView`** + **`MeasureRow`** + parent/child rows — §6 |
| **Tabbed settings** | Codejock / .frm | `TabControl` + `SelectedItem` |

### 2.13.8 Out of scope (Selector v1)

- Multi-select (`SelectionMode`, `SelectedItems`)
- Full **`PropertyPath`** for `SelectedValuePath` (nested paths)
- Keyboard navigation parity (Home/End/Ctrl+click) — add incrementally
- **`Selector.IsSelectionActive`** and focus visuals — later polish

### 2.13.9 Open questions (see §8)

- [ ] **`ListBox` vs `ListView` only** — ship both (WPF parity) or `ListView` only with ItemsPanel for dialog grids?
- [ ] **`ListIndex` deprecation** — one release alias property or hard break with `SelectedIndex`?
- [ ] **UnboundListView** — expose same Selector DPs when `ItemsSource` is owner-draw indices only?

---

## 2.14 ListView stack review — ItemsSource, DataTemplate, bound/unbound, vbWidgets

**Request:** Review the complete **ItemsSource / DataTemplate / ListView (bound & unbound)** implementation. **`ListViewBase` and `TextBoxBase` are the vbWidgets-derived Cairo engines** embedded in VCF; public **`ListView`**, **`UnboundListView`**, **`TextBox`** wrap them with dependency properties, XAML, and MVVM. Runtime depends on **vbRichClient5** (`cWidgetBase`); **not** on vbWidgets.dll.

**Verdict:** The stack is **three layers**: VCF **`ListView` / `UnboundListView`** shells over **`ListViewBase`** (~1.2k lines) — which **is** the adapted **vbWidgets/cwWidget** list engine (same relationship as **`TextBox`** → **`TextBoxBase`**). Bound mode adds a partial WPF items pipeline with known bugs. **Agreed:** merge the two public list types; evolve **`ListViewBase`** (refactor or **new implementation**) for InvoiceGrid — not a separate parallel stack (§2.14.10).

### 2.14.1 Architecture today

```text
┌─────────────────────────────────────────────────────────────┐
│  ListView (bound)              UnboundListView (unbound)     │
│  · ItemsSource DP              · No ItemsSource              │
│  · ItemTemplate / DataTemplate   · Re-raises OwnerDrawItem     │
│  · ListCollectionView sync       · Host sets Base.ListCount    │
│  · IItemsControl (stubs!)        · Full event surface          │
└──────────────────────────┬──────────────────────────────────┘
                           │ composes m_Base
┌──────────────────────────▼──────────────────────────────────┐
│  ListViewBase (Friend, ~1179 lines)                          │
│  · cWidgetBase W — W_Paint → Draw → RaiseEvent OwnerDrawItem │
│  · VCF.ScrollBar V/H (InnerWidget) — was cwScrollBar         │
│  · ListCount, ListIndex, columns, headers, multi-select, …   │
│  · Image grid mode, row selector, hover bar, keyboard/mouse  │
└──────────────────────────┬──────────────────────────────────┘
                           │
┌──────────────────────────▼──────────────────────────────────┐
│  vbRichClient5: cWidgetBase, cCairoContext, cWidgets         │
│  (ScrollBar.cls / TextBoxBase.cls use same pattern)         │
└─────────────────────────────────────────────────────────────┘
```

**Parallel wrappers (same widget pattern):**

| VCF control | Base engine | Root widget | Scrollbars |
|-------------|-------------|-------------|------------|
| **ListView** | `ListViewBase` | `m_Base.Widget` | `ListViewBase` → `VCF.ScrollBar` |
| **TextBox** | `TextBoxBase` (~1.5k lines) | `W` = `cWidgetBase` | Optional `VCF.ScrollBar` on `W.Widgets` |
| **ScrollBar** | (none — self-contained) | `W` = `Cairo.WidgetBase` | N/A — also vbWidgets/cw lineage |

**Pattern:** `*Base` = **widget engine** (vbWidgets source, in-repo); **VCF control** = WPF-shaped façade (DPs, XAML, binding). **New implementations** of `ListViewBase` / `TextBoxBase` are allowed when refactor cost exceeds rewrite — same public **`ListView`** / **`TextBox`** API.

**POS usage:**

| Control | Where | Mode |
|---------|-------|------|
| **ListView** | `LoginView.xml` — `ItemsSource="{Binding UserList}"` | Bound (minimal XAML — no inline ItemTemplate in file) |
| **UnboundListView** | `OrderItemsView.xml` — stub for invoice lines | Unbound / owner-draw (intended InvoiceGrid replacement) |

### 2.14.2 ItemsSource pipeline (bound ListView)

**Flow:**

1. XAML/code sets **`ItemsSource`** dependency property.
2. **`DependencyPropertyChanged`** → **`SetItemSource`** only if `TypeOf Value Is ObservableCollection` — **arrays, `List`, ADO fail silently**.
3. **`CollectionViewSource.GetDefaultView`** → **`ListCollectionView`** (one cached view per collection pointer).
4. **`ListCount = collection.Count`** on `ListViewBase`; selection sync **`ListIndex` ↔ `CurrentPosition`**.
5. On **`CollectionChanged`**, parallel **`ItemTemplates`** list resized (slots = `Empty` or cached `DataTemplate`).

**Pain points:**

| Issue | Detail |
|-------|--------|
| **ObservableCollection-only** | No `IEnumerable`, no adapters — Login works; generic binding does not |
| **`IItemsControl.ItemTemplate` stubs** | `ListView.cls` 319–325 — interface methods **empty** |
| **`ListCollectionView.Initialize` static flag** | Only **first** view instance can initialize — **multi-list bug** |
| **Template cache** | `Empty` + `On Error` to detect state — fragile |
| **`CollectionChangedActionMove`** | **Unimplemented** — reorder desyncs cache indices |
| **Non-object items** | `DrawDefaultItemTemplate` called **inside child loop** (once per template child) |
| **Clone surface** | Only **`ICloneable`** types (chiefly **TextBlock**); **Image** has no `Clone` |

### 2.14.3 DataTemplate pipeline

**Definition:** `DataTemplate.cls` — **`Children`** only (`UIElementCollection`); optional `DataType`, `Key`.

**XAML loading:** `ListView.Resources` / `ListView.ItemTemplate` via `XAMLReader` dot-properties; typed resources registered as **`DataTemplate_<DataType>`** (`XAMLReader.cls` ~406).

**Per-row inflation (`CreateDataTemplate`):**

1. Resolve template: `TryFindResource("DataTemplate_" & TypeName(Item))` → else **`m_ItemTemplate`**.
2. **`Clone`** each child (`ICloneable`).
3. **`DataContext = Item`** on each `IUIElement` child.
4. Cache in **`ItemTemplates(Index)`**.

**Draw (`OwnerDrawItem`):** For each visible row index, **`FindDataTemplate`** → foreach child **`MoveChild`** (Design* scale) → **`IVisualChild.DrawOn`** on Cairo context. **No widget subtree per row** — immediate-mode draw only.

**Invalidation:** `ListViewPropertyChangedHandler` on items implementing **`INotifyPropertyChanged`** clears cache slot + refresh.

**WPF gap:** No **`ItemContainerGenerator`**, no visual tree per row, no **`ContentPresenter`**, templates not usable on **`ItemsControl`** (§2.12.8).

### 2.14.4 UnboundListView (unbound)

- Same **`ListViewBase`** engine; **forwards** `OwnerDrawItem`, `Click`, `HeaderClick`, scroll events to host.
- Host responsibility: set **`Base.ListCount`**, handle **`OwnerDrawItem`** (InvoiceGrid-style).
- **`MoveChild` copied from ListView but never used** — dead code.
- **`Refresh()`** → `m_Base.Widget.Refresh`.

**Why two public classes today:** Historical split — bound MVVM vs owner-draw POS grid. **Agreed (§2.14.10): merge** into one **`ListView`** on **`Selector`** — bound when `ItemsSource` set; owner-draw presenter when not (or explicit presentation mode).

### 2.14.5 ListViewBase engine (vbWidgets / cwWidget lineage)

**Origin:** Comments reference **cwWidget convention** and **cwScrollBar**; engine is **custom list UX on Cairo**, not a thin alias of a vbWidgets List class.

**Capabilities (beyond WPF ListView defaults):**

- Column headers, resize, sort state events
- Row selector column, multi-select (`mSelBits`)
- Image grid mode (`mImageView`, `mVisibleCols`)
- Alternate row colors, hover bar, drag selection, edge auto-scroll (`TScroll` timer)

**Render loop:** `W_Paint` → `AdjustDimensions` (scrollbar layout) → `Draw` → for each visible cell **`RaiseEvent OwnerDrawItem`**.

**Scrollbar integration:** Child **`VCF.ScrollBar`** widgets on `W.Widgets`; `ScrollIndex` = `VScrollBar.Value`; wheel routed in `W_MouseWheel`.

**Coupling risk — POS InvoiceGrid (§6):** Needs **read-only multi-column list** with **parent/child rows** and **per-row height** — **not** in-cell editing or DataGrid semantics. Engine must support **columns + `MeasureRow` + hierarchical flat index**; bound `DataTemplate` alone is insufficient without **`HierarchicalDataTemplate`** or owner-draw row paint.

### 2.14.6 Engine layer — ListViewBase / TextBoxBase (= vbWidgets source in VCF)

**Clarification (agreed):**

- **`ListViewBase`** and **`TextBoxBase`** **are** the vbWidgets/cwWidget-derived implementations — already **in the VCF repo**, not an external DLL list control.
- **`ListView`**, **`UnboundListView`**, **`TextBox`** are **thin VCF wrappers** (DPs, styles, XAML, binding, DataTemplate).
- **`ScrollBar.cls`** is the same lineage (Colin Edwards / cwWidget); ListViewBase still has commented **`cwScrollBar`** references from migration to **`VCF.ScrollBar`**.
- vbWidgets source remains **public reference** for fixes/porting; **no separate “vendor snippets” fork** — evolve **`ListViewBase`** in place or **replace with a new engine class** behind the same wrapper.

**Implication for InvoiceGrid (§6):** **Read-only hierarchical multi-column list** — extend **`ListViewBase`** (columns, **`MeasureRow`**, parent/child row levels). **Not** a DataGrid / editable grid control. Merged **`ListView`** + **`Selector`** for selection.

**Attribution:** Document cwWidget/vbWidgets lineage in VCF **`THIRD_PARTY.md`** / architecture doc.

### 2.14.7 WPF-aligned target architecture (refactor proposal)

**Unify public surface (agreed — merge ListView + UnboundListView):**

```text
Selector (§2.13)
└── ListView                    — single public control (breaking: drop UnboundListView type)
    ├── ItemsSource set         → bound: DataTemplate + ItemsSource pipeline
    └── ItemsSource null /      → owner-draw: OwnerDrawItem (InvoiceGrid path)
        OwnerDraw mode
    └── composes
        ListViewBase            — current engine, refactored OR replaced (§2.14.6)
```

**XAML migration:**

| Today | Target |
|-------|--------|
| `<ListView ItemsSource="…"/>` | unchanged |
| `<UnboundListView …/>` | `<ListView …/>` — owner-draw via events or `ItemsSource` unset; document in **`MIGRATION.md`** |

**Engine vs presentation split:**

| Layer | Responsibility |
|-------|----------------|
| **`ListViewBase`** (evolved or new impl) | Scroll, hit-test, selection chrome, columns, **`MeasureRow`** / variable height for invoice |
| **Presentation strategy** inside merged **`ListView`** | **Bound:** DataTemplate + DrawOn; **Owner-draw:** `OwnerDrawItem` event / handler |
| **`ItemTemplateCache`** | Explicit state enum; fix Move; support invalidation |
| **`ListCollectionView`** | Remove static init bug; optional sort later |

**ItemsSource v2:**

- Accept **`ObservableCollection`** (v1) + **`IEnumerable`** via one-shot snapshot or adapter interface.
- Document **`CollectionView`** sync with **`SelectedItem`**.

**DataTemplate v2:**

- Shared with **`ItemsControl`** (§2.12.8) — same inflate/clone module.
- Implement **`ICloneable`** on **Image**, **Border** (draw primitives) or use **factory** from XAML node.
- Row layout: eventually **Measure/Arrange per row** instead of **`MoveChild` Design* scale**.

**InvoiceGrid path (§6):** Engine adds **multi-column row layout**, **`MeasureRow(Index, width) → height`**, **row level** (parent/child). POS binds **`SalesOrderItem`** / modifier lines via **`ItemsSource`** or owner-draw. **No editing API** required on the control.

### 2.14.8 Known bugs to fix (P0 — before POS grid migration)

1. **`ListCollectionView.Initialize` static flag** — per-instance init.
2. **`IItemsControl_ItemTemplate` stubs** — wire or remove interface.
3. **`CollectionChangedActionMove`** — reorder `ItemTemplates` with items.
4. **Non-object row draw loop** — call default template once per row.
5. **ItemsSource type gate** — fail loud in debug / log when wrong type.

### 2.14.9 Phasing (adds to §4)

| Phase | ListView deliverable |
|-------|---------------------|
| **0** | This review → **`VCF_LISTVIEW_ARCHITECTURE.md`** in VCF repo; golden tests (bind, template, select, scroll) |
| **4** | **Merge `UnboundListView` → `ListView`**; bug fixes §2.14.8; **`Selector`** DPs |
| **4** | Shared **DataTemplate engine** with ItemsControl |
| **5** | **`ListViewBase` evolution** — multi-column rows, **`MeasureRow`**, parent/child levels; POS InvoiceGrid spike |
| **6+** | Column/field model, virtualization if needed on same engine |

### 2.14.10 Decisions (agreed — see §9)

| # | Question | Decision |
|---|----------|----------|
| **1** | Merge **ListView + UnboundListView**? | **Yes** — one **`ListView`** on **`Selector`**; remove **`UnboundListView`** type (breaking); owner-draw when no `ItemsSource` or explicit owner-draw mode |
| **2** | vbWidgets vs ListViewBase? | **`ListViewBase` / `TextBoxBase` = vbWidgets-derived engine in VCF**; wrappers stay thin; **refactor or new engine implementation** allowed — not a separate DLL/vendor fork |
| **3** | Invoice grid engine? | **Same engine as #2** — extend/replace **`ListViewBase`**; merged **`ListView`** public API; **not** a parallel **`VirtualizingListView`** product |

**Still open (§8):**

- [ ] **Bound list presentation:** owner-draw only long-term vs optional full visual tree per row for small lists?

---

## 2.15 Visibility / Collapsed and StackPanel

**Report:** **`Visibility`** / **`Collapsed`** are **partially** implemented and **inconsistent** across controls. **`StackPanel`** does **not** exist — POS and WPF alignment need both on **`FrameworkElement`** with proper layout semantics (Phase 1–2).

### 2.15.1 Visibility today — split API, incomplete semantics

**WPF model (target):**

| Value | Render | Layout slot |
|-------|--------|-------------|
| **Visible** | Yes | Yes |
| **Hidden** | No | Yes (space reserved) |
| **Collapsed** | No | **No** (parent re-arranges) |

**VCF enum exists** (`IControl.Visibility`: Visible=0, Hidden=1, Collapsed=2) but **`SetVisibility` in `modVisibilityHelper.bas` ignores the distinction:**

```vb
bVisible = (Value = VisibilityVisible)   ' Hidden and Collapsed both → W.Visible = False
```

**No panel skips collapsed children** in measure/arrange — `UniformGrid` still allocates cells; hidden buttons leave empty grid slots (POS menu grids use `Visible="{Binding Visible}"` on buttons).

### 2.15.2 Per-control support (audit)

| Control | API | Bindable DP? | Notes |
|---------|-----|--------------|-------|
| **Panel**, **UserControl** | `Visibility` enum property | **No** (plain property) | `SetVisibility` → widget only |
| **TextBox**, **ListView**, **UnboundListView**, **WindowsFormsHost** | `Visibility` enum | **No** | Same helper |
| **Button**, **UniformGrid** | **`Visible` bool DP** | **Yes** | Different name from WPF; no Hidden/Collapsed |
| **Border**, **Image**, **TextBlock** | **None** | — | Widget always `W.Visible = True` at init |
| **Window** | **None** on element | — | Form visibility separate |

**XAML:** POS uses **`Visible="{Binding …}"`** (Button, grids) and **`Visibility="0"`** on Panel (numeric — maps to **Visible**, not Collapsed). No `Collapsed="…"` or `Visibility="Collapsed"` in use yet.

**Designer** (`PropertyEditor.cls`) knows Visible/Hidden/Collapsed strings — runtime layout does not fully implement them.

### 2.15.3 Gaps vs WPF (and POS pain)

1. **Two property names** — `Visible` vs `Visibility`; blocks WPF XAML migration.
2. **`Collapsed` ≠ `Hidden`** — not implemented; toggling visibility in **UniformGrid** does not reclaim space → workarounds (empty `TextBlock` placeholders, manual grid math).
3. **Not on all elements** — Border/TextBlock/Image cannot bind visibility.
4. **Not a DP on FrameworkElement** — inconsistent binding/rebind (§2.5).
5. **Parent layout unaware** — until **Measure/Arrange**, Collapsed cannot work correctly on any panel.

### 2.15.4 StackPanel — missing, high value

**Not in VCF** (grep: zero references). POS layouts rely on **`UniformGrid`** + fixed **`Design*`** coords — vertical/horizontal stacks (toolbars, form fields, settings rows) are awkward.

**WPF `StackPanel`:**

- **`Orientation`:** Horizontal | Vertical
- Children stacked; **Collapsed** children omitted from layout
- **`Margin`**, alignment per child (with Measure/Arrange)

**POS examples that would simplify:**

- Customer panel rows, navigation command strips (today partial UniformGrid + empty cells)
- Settings / dialog field columns after `.frm` → XAML migration
- **`ItemsPanelTemplate`** default for ItemsControl (§2.12.8) — WPF default is StackPanel

**Dependency:** Meaningful **StackPanel** requires **Phase 1 Measure/Arrange** (or a minimal arrange-only StackPanel on current widget model as interim — document deviation).

### 2.15.5 Recommended direction (VCF team)

| Priority | Deliverable |
|----------|-------------|
| **P0 / Phase 1** | **`Visibility` DP** on **`FrameworkElement`** — values **`Visible`**, **`Hidden`**, **`Collapsed`** (XAML strings + enum); **remove `Visible` bool** (breaking; map in `MIGRATION.md`) |
| **P0 / Phase 1** | Fix **`SetVisibility` / layout integration** — panels call **`UIElement.Visibility`** during **Measure/Arrange**; Collapsed → desired size 0 |
| **P1 / Phase 2** | **`StackPanel`** — Vertical + Horizontal; first real panel after Grid |
| **P1** | Apply Visibility DP to **all** current controls (Border, TextBlock, Image, Button, …) via base class |
| **P2** | **`IsHitTestVisible`** when Hidden (optional) |

**POS migration:**

- `Visible="{Binding …}"` → **`Visibility="{Binding …, Converter=BoolToVisibility}"`** or bind Visibility directly on VM
- Menu grid patterns: prefer **StackPanel** or **Grid** over hiding buttons in fixed **UniformGrid** cells

### 2.15.6 Open questions (see §8)

- [ ] **Interim StackPanel** before full Grid — vertical-only stack on widget coords acceptable for one release?
- [ ] **`Visibility="0"` in POS XAML** — bug (Visible) or intentional? Audit LeftColumn Panel — §2.15.2

---

## 2.16 `res:` includes — WPF-aligned resources (replace app hack)

**Question:** How to re-implement POS **`res:`** fragment loading to match **WPF** resource / include standards?

**Decision:** **`res:` as a custom XML namespace is legacy.** Target = **`ResourceDictionary` + `MergedDictionaries` + `{StaticResource}` / `{DynamicResource}`**, with **`ContentControl`** only for **single-content regions** (navigation / panel swap). **Move loader into VCF** (§2.6); delete POS-only `TryCreateObject` when migration completes. Resolved detail: §2.16.6.

### 2.16.1 What `res:` does today

**Three separate mechanisms share the `res:` prefix:**

| Mechanism | Where | Behavior |
|-----------|-------|----------|
| **Inline XAML element** | `<res:Commands\SKashCheckout/>` in parent tree | `XAMLReader` → `CreateInstance("res", path)` → **`ObjectConstructor.TryCreateObject`** → load XML string from **`MyApp.XAMLResources(path)`** → **`XAMLReader.Load`** → add root object as child |
| **Dynamic panel swap** | `SetPanelContentCommand` param `"NavigationArea1,res:Screens\...\Layout2,1"` | **`VCF.CreateInstance("", Replace(res:, res.))`** → same loader; optional cache in **`Application.Resources(Parent.ChildKey)`** |
| **View / app bootstrap** | `LoadSuperclassData(Me, XAMLResources("Screens\..."))` | **Not `res:`** — code looks up flat dictionary key directly |

**Resource storage (POS):**

```text
LoadXAMLResources()  →  scan Resources\XAML\**\*.xml
                      →  Application.Resources("XAML")[relativePathWithoutExt] = fileText
```

**Parallel to WPF `Application.Resources` in `MyApp.xml`:** styles, `ThemesManager`, `AppCommands` use **`{StaticResource ResourceKey=AppCommands}`** and **`UIElementBase.FindResource`** walk — **different bucket** from flat **`XAMLResources`** file index.

**Flow (inline fragment):**

```text
<res:Widgets\StatusBar/> 
  → CreateInstance Namespace="res" Class="Widgets\StatusBar"
  → CustomConstructor "res.Widgets\StatusBar"
  → XAMLResources("Widgets\StatusBar") → XAMLReader.Load(XML) → IUIElement tree
```

**Problems vs WPF:**

- **App-specific** — only works if POS registers `ObjectConstructor`
- **Silent failure** — missing key → `Nothing`, empty slot
- **No `x:Key`** on fragments — identity = file path only
- **Not in resource lookup chain** for `{StaticResource}` — two mental models
- **Dynamic swap in code** (`SetPanelContentMethod`) — not bindable / not declarative
- **`res:` in CommandParameter strings** — not XAML resource syntax

### 2.16.2 WPF model (target semantics)

| Need | WPF pattern |
|------|-------------|
| **App-wide styles/objects** | `Application.Resources` / `ResourceDictionary` |
| **Split XAML across files** | **`ResourceDictionary.MergedDictionaries`** + **`Source="Commands/SKashCheckout.xaml"`** |
| **Reference by key** | **`{StaticResource Key}`** / **`{DynamicResource Key}`** |
| **Compose fragment into tree** | **`UserControl`**, **`ContentControl Content=`**, or **`DataTemplate`** — not a magic empty element |
| **Swap panel content at runtime** | **`ContentControl`** + binding / **`DataTemplateSelector`** — not comma-separated command params |
| **Type in XAML** | **`xmlns:local`** → CLR namespace + class — for **code-behind views** with `x:Class` |

VCF already has **partial** WPF: `Application.Resources`, `{StaticResource}`, `TryFindResource` parent walk, `ThemesManager` ≈ dynamic theme dict.

### 2.16.3 Target VCF design

**A. Unified resource system (Phase 3)**

1. **`ResourceDictionary` class** — `ObservableDictionary` + **`MergedDictionaries`** collection.
2. **Loader** — on merge or app init, load **`Source="relative/path.xml"`** into that dictionary (same files as today under `Resources\XAML\`).
3. **Keys** — WPF rules (**§2.16.7**):
   - **Required explicit `x:Key`** on every dictionary entry (command chrome, templates, singleton resources).
   - **`x:Class` views:** key = **CLR type name** (same as WPF `local:MyView` registration), not file path.
   - **No path-based keys** (`Commands\SKashCheckout`) in the final spec — paths belong on **`Source=`** only.
4. **Remove flat `XAMLResources` bucket** — merged dicts replace it (breaking; **`MIGRATION.md`**).

**B. Replace inline `<res:…/>` (WPF patterns — §2.16.7)**

| Pattern | When | XAML example |
|---------|------|--------------|
| **1. DataTemplate + clone** | Command/button **fragments** (most POS `res:Commands\…`) | Fragment file: **`<DataTemplate x:Key="SKashCheckout">`** … **`</DataTemplate>`**; parent child slot: **`ResourceReference Key="SKashCheckout"`** (VCF inflates **new instance** per slot — WPF `LoadContent` semantics) |
| **2. UserControl + xmlns** | Screen **views** with **`x:Class`** | **`xmlns:views`** + **`TypeRegistry`** — **`views:SalesOrderView`** (WPF `local:` style); no `res:` path element |
| **3. ContentControl** | **One** navigation / swap region | **`ContentControl Content="{Binding NavArea1View}"`** — not per-button grids |
| **4. ItemsControl** | **Long-term** command grids | **`ItemsSource`** + **`ItemsPanelTemplate` (UniformGrid)** + **`ItemTemplate`** — WPF standard for many similar cells |

**Command fragments:** store as **`DataTemplate`**, not a shared **`Button x:Key`** instance — WPF forbids attaching one **`UIElement`** to multiple parents.

**Transitional shim (one release):** **`res:path`** → resolver loads merged dict entry; if entry is **`DataTemplate`**, clone; if legacy bare root (no template wrapper), wrap-on-load + warn. Emit **deprecation warning** in debug.

**C. Replace `SetPanelContent` / CommandParameter `res:` strings**

WPF-aligned:

```xml
<ContentControl Name="NavigationArea1"
                Content="{Binding NavigationArea1Content}"/>
```

ViewModel sets content key or object; or use **`DataTemplate`** + **`ContentTemplateSelector`**.

Migration: `SetPanelContentCommand` → binding + **`INotifyPropertyChanged`**; cache in VM or resource dict, not **`Application.Resources(Parent.ChildKey)`** ad hoc strings.

**D. Built-in framework API (extends §2.6)**

```text
IXamlTypeResolver.Resolve(namespace, name)
  1. TypeRegistry (x:Class views, converters)
  2. ResourceDictionary.TryGet(name) → parse/inflate if lazy
  3. VCF built-ins (Button, …)
  4. Error — no silent Nothing
```

**`res:` namespace** → alias for **`http://schemas.demac.vcf/2026/xaml/resources`** → step 2 (optional transitional).

**E. URI / Source form (WPF-like)**

```xml
<Application.Resources>
  <ResourceDictionary>
    <ResourceDictionary.MergedDictionaries>
      <ResourceDictionary Source="Commands/SKashCheckout.xml"/>
      <ResourceDictionary Source="Screens/SalesOrder/QuickService/LeftColumn.xml"/>
    </ResourceDictionary.MergedDictionaries>
  </ResourceDictionary>
</Application.Resources>
```

Parent screen references **`{StaticResource …}`** only — no inline `res:` tags.

### 2.16.4 POS migration examples

**Before (inline include):**

```xml
<UniformGrid …>
  <res:Commands\SKashCheckout/>
  <res:Commands\AlternatePaymentCheckout/>
</UniformGrid>
```

**After (merged dictionary + DataTemplate — WPF-correct):**

```xml
<!-- Commands/SKashCheckout.xml (merged via Source=) -->
<DataTemplate x:Key="SKashCheckout">
  <Button Command="…" …/>
</DataTemplate>

<!-- Parent: clone template per cell — no ContentControl wrapper -->
<UniformGrid …>
  <ResourceReference Key="SKashCheckout"/>
  <ResourceReference Key="AlternatePaymentCheckout"/>
</UniformGrid>
```

**Long-term (dense command grids — preferred WPF):**

```xml
<ItemsControl ItemsSource="{Binding CheckoutCommands}">
  <ItemsControl.ItemsPanel>
    <ItemsPanelTemplate>
      <UniformGrid Columns="8"/>
    </ItemsPanelTemplate>
  </ItemsControl.ItemsPanel>
  <ItemsControl.ItemTemplate>
    <DataTemplate>
      <Button Command="{Binding Command}" Content="{Binding Label}" …/>
    </DataTemplate>
  </ItemsControl.ItemTemplate>
</ItemsControl>
```

AI migration: wrap legacy fragment roots in **`<DataTemplate x:Key="…">`**, add explicit keys, replace **`<res:…/>`** with **`ResourceReference`** (or refactor grid to **ItemsControl** where a command list VM exists).

**Before (panel swap):**

```text
CommandParameter="NavigationArea1,res:Screens\...\Layout2,1"
```

**After:**

```xml
<ContentControl Name="NavigationArea1" Content="{Binding NavArea1View}"/>
```

**Before (view load in code):**

```vb
Key = "Screens\SalesOrder\QuickService\Layout"
LoadSuperclassData Me, XAMLResources(Key)
```

**After:**

```vb
' ViewBase.InitializeComponent uses StartupUri / x:ResourceKey — same resolver as MergedDictionaries
```

### 2.16.5 Phasing

| Phase | Deliverable |
|-------|-------------|
| **3** | **`ResourceDictionary` + `MergedDictionaries` + `Source=`**; VCF **`XamlResourceResolver`**; fail-loud load |
| **3** | **`res:` shim** → resolver (deprecate); POS stops using custom **`TryCreateObject`** |
| **4** | **`ContentControl`** panel swap pattern; migrate **`SetPanelContentCommand`** |
| **7 / POS** | AI migration: `<res:…/>` → **`DataTemplate` + `ResourceReference`** or **ItemsControl**; explicit **`x:Key`** on all fragments |

### 2.16.6 Resolved decisions (WPF-aligned)

*(Former open questions — locked for VCF spec.)*

**1. Fragment keys — explicit `x:Key` required**

| Rule | Detail |
|------|--------|
| **Final spec** | Every merged resource entry has **`x:Key`** (or **`x:Class`** view registered in **`TypeRegistry`** by type name). |
| **Paths** | File path is **`ResourceDictionary Source=`** only — **never** the lookup key. |
| **Transitional (one release)** | If a merged file has no **`x:Key`**, loader uses **filename without extension** (e.g. `SKashCheckout` from `Commands/SKashCheckout.xml`) and logs **deprecation warning** — **not** full relative path with backslashes. |

WPF does not use path strings as resource keys; filename fallback is migration-only sugar, not target API.

**2. Command grids — `DataTemplate` clone, not `ContentControl` per cell**

| Rule | Detail |
|------|--------|
| **Why not ContentControl × N** | WPF treats a **`UIElement`** in a dictionary as **one instance**; repeating **`Content="{StaticResource Foo}"`** across 50 cells is invalid (second parent fails). POS **`res:`** worked because each tag **re-parsed** and created a **new tree**. |
| **Target inclusion syntax** | **`ResourceReference Key="…"`** (or equivalent markup extension) in a panel’s child collection → resolver **`DataTemplate.LoadContent()`** → append **cloned** subtree as direct child — **no extra wrapper control**. |
| **Fragment shape** | Command files → **`<DataTemplate x:Key="…">`** root, not bare **`<Button x:Key="…">`**. |
| **Long-term grids** | **`ItemsControl`** + **`ItemsPanelTemplate` (UniformGrid)** + **`ItemTemplate`** / bound **`DataTemplate`** — standard WPF for many similar buttons. |
| **`ContentControl` scope** | **Single-content regions only** — navigation areas, dialog body, **`SetPanelContent`** replacement — **`Content="{Binding …}"`**. |

**3. `ResourceReference` (VCF markup)**

Minimal WPF-aligned child element for “inflate keyed template here” (name aligns with **`StaticResource`** lookup, distinct from WPF’s `{StaticResource}` on properties):

```xml
<ResourceReference Key="SKashCheckout"/>
```

Equivalent to WPF **`ContentPresenter`** + template inflation, but produces **direct panel children** for **`UniformGrid`** cell ordering.

### 2.16.7 Open questions (see §8)

*(None — resource include model locked; see §2.16.6.)*

---

## 2.17 Styles, themes & semantic classes — WPF alignment

**Question:** VCF uses **HTML/CSS-inspired styling** — implicit type defaults, **named style “classes”** (`ButtonSubmit`, `ListWrapper`), and a **theme dictionary** (`ThemesManager`) for colors. How do we align this with **WPF**?

**Decision:** The model is **already mostly WPF** (`Style`, `Setter`, `BasedOn`, `{StaticResource}`). Align by (1) treating **named `x:Key` styles** as the WPF **`Style`** mechanism (not a separate `Class=` attribute), (2) mapping **`{ThemeResource}` → `{DynamicResource}`** with theme palettes as **merged `ResourceDictionary` files**, (3) keeping **semantic button styles** as thin named styles over theme brush keys. Do **not** add HTML-style `Class="submit cancel"` parsing.

### 2.17.1 What VCF does today (CSS metaphor → WPF reality)

| CSS-like idea | VCF mechanism | WPF equivalent |
|---------------|---------------|----------------|
| **Element default** (`button { … }`) | `<Style TargetType="Button">` (no `x:Key`) → registered as **`Button.BaseStyle`** | **Default style** — WPF key **`{x:Type Button}`** |
| **Named class** (`.submit`, `.cancel`) | `<Style x:Key="ButtonSubmit">` + **`Style="{StaticResource ButtonSubmit}"`** | **Keyed `Style`** on **`FrameworkElement.Style`** |
| **CSS variables** | **`ThemesManager`** per-theme **`ObservableDictionary`** of brushes/strings | **Theme `ResourceDictionary`** + **`{DynamicResource}`** |
| **Theme switch** | **`ThemesManager.ActiveThemeName`**; styles listen **`ThemesManager_ThemeChanged`** → **`Style.OnStyleChanged`** → re-apply | Swap **merged theme dictionary**; **`DynamicResource`** invalidation |
| **Property values in theme** | **`{ThemeResource ButtonBackgroundSubmit}`** in style setters | **`{DynamicResource ButtonBackgroundSubmit}`** on **`SolidColorBrush`** |

**POS example (already WPF-shaped):**

```xml
<!-- MyApp.xml — theme palette -->
<SolidColorBrush x:Key="ButtonBackgroundSubmit" Color="#008000"/>
<SolidColorBrush x:Key="ButtonBackgroundCancel" Color="#800000"/>

<!-- Named semantic styles (“classes”) -->
<Style TargetType="Button" x:Key="ButtonSubmit">
  <Setter Property="BackColor" Value="{ThemeResource ButtonBackgroundSubmit}"/>
</Style>
<Style TargetType="Button" x:Key="ButtonCancel">
  <Setter Property="BackColor" Value="{ThemeResource ButtonBackgroundCancel}"/>
</Style>

<!-- Usage -->
<Button Style="{StaticResource ButtonSubmit}" Command="{Binding Login}"/>
<Button Style="{StaticResource ButtonCancel}" Command="{Binding ClearInput}"/>
```

**Implementation notes:**

- **`StyleManager`** applies setters via **`SetCurrentValue`** (WPF local-vs-style precedence intent).
- Keyed styles auto-**`BasedOn`** implicit **`TargetType.BaseStyle`** when `BasedOn` omitted (`Style.Initialize` → **`GetBaseStyle`**).
- **`ThemeResource`** markup extension reads **`ThemesManager.ActiveTheme(Key)`** only — **not** the general **`TryFindResource`** chain.

### 2.17.2 Gaps vs WPF

| Gap | Impact |
|-----|--------|
| **`{ThemeResource}` is separate from `{DynamicResource}`** | Two markup extensions for “resource that changes with theme”; confuses WPF developers |
| **`Button.BaseStyle` key** vs WPF **`{x:Type Button}`** | Same role, different key convention |
| **Theme dict outside `ResourceDictionary` merge** | Theme brushes not in normal resource lookup / merged dict story (§2.16) |
| **Setters use `BackColor` not `Background` + `Brush`** | Acceptable interim; long-term map to brush DPs when property names align (§3.3) |
| **Some default styles hardcode colors** | e.g. `ListView` / `UniformGrid` in `MyApp.xml` — bypass theme (see `THEME_INHERITANCE_AND_IMPROVEMENTS.md`) |
| **No `Style.Triggers` / `ControlTemplate`** | Hover/disabled via widget fields (`HoverColor`, …) in flat setters — not WPF visual state model (Phase 6) |
| **Legacy POS `StyleManager.cls`** | Codejock/VB6 form styling — **out of scope** for VCF XAML path |

### 2.17.3 Target design (WPF-aligned)

**A. Styles — keep WPF model; no `Class=` attribute**

| Rule | Detail |
|------|--------|
| **Implicit default** | `<Style TargetType="T">` without key — register as WPF **`{x:Type T}`** (migrate from **`T.BaseStyle`** alias; transitional dual lookup) |
| **Named semantic styles** | **`x:Key="ButtonSubmit"`** etc. — **`Style="{StaticResource ButtonSubmit}"`** on element |
| **`BasedOn`** | Explicit or auto-chain to type default — **`BasedOn="{StaticResource {x:Type Button}}"`** in migration docs |
| **Do not add** | HTML **`Class="submit cancel"`** multi-class strings — not WPF; use one **`Style`**, or **`BasedOn`** + overrides, or future **`Style.Triggers`** |

**B. Theme palettes — `ResourceDictionary` per theme**

```xml
<!-- Themes/Light.xml (merged or swapped) -->
<ResourceDictionary>
  <SolidColorBrush x:Key="ButtonBackground" Color="#EEEEEE"/>
  <SolidColorBrush x:Key="ButtonBackgroundSubmit" Color="#008000"/>
  <SolidColorBrush x:Key="ButtonBackgroundCancel" Color="#800000"/>
  <sys:String x:Key="FontName">Segoe UI</sys:String>
</ResourceDictionary>
```

- **`ThemesManager`** becomes a thin **active-theme selector** that **merges/unmerges** the active theme **`ResourceDictionary`** into **`Application.Resources`** (same swap model as WPF theme packs).
- **Semantic colors stay in theme dict** (green submit / red cancel) — WPF has no built-in “SubmitButtonBrush”; app resources are correct.

**C. `{ThemeResource}` → `{DynamicResource}` (resolved)**

| Rule | Detail |
|------|--------|
| **Implement `{DynamicResource}`** | Standard WPF markup; participates in **`TryFindResource`** walk including **active theme merged dict** |
| **`{ThemeResource}`** | **Deprecated alias** (one release) → identical to **`{DynamicResource}`**; remove from **`MIGRATION.md`** target XAML |
| **Theme switch** | Unmerge old theme dict, merge new, **invalidate dynamic resources**, **`Style.OnStyleChanged`** re-apply (extend existing **`ThemesManager_ThemeChanged`** path) |
| **`{StaticResource}` in setters** | Use for **non-theme** resources (e.g. **`BasedOn`**, converter keys); use **`DynamicResource`** for **theme brush keys** in setters |

**D. Semantic button pattern (locked — matches POS today)**

Three layers — WPF-standard separation of concerns:

```text
Theme dict     →  ButtonBackgroundSubmit, ButtonBackgroundCancel, ButtonBackgroundWarning  (tokens)
Default style  →  TargetType=Button (font, border, corner radius from theme)
Named styles   →  ButtonSubmit / ButtonCancel / ButtonWarning (one setter: BackColor/Background)
XAML usage     →  Style="{StaticResource ButtonSubmit}" only where semantics needed; else default Button style
```

**Do not** hardcode `#008000` on individual buttons — keep overrides in **theme** or **named style** only (POS theming goal in `THEME_INHERITANCE_AND_IMPROVEMENTS.md`).

**E. Split `MyApp.xml` resources (with §2.16)**

```xml
<Application.Resources>
  <ResourceDictionary>
    <ResourceDictionary.MergedDictionaries>
      <ResourceDictionary Source="Themes/Default.xml"/>
      <ResourceDictionary Source="Styles/Controls.xml"/>   <!-- implicit + ButtonSubmit, ListWrapper, … -->
    </ResourceDictionary.MergedDictionaries>
    <ThemesManager x:Key="ThemesManager" …/>  <!-- selects which Themes/*.xml is merged -->
  </ResourceDictionary>
</Application.Resources>
```

### 2.17.4 Migration examples

**Theme setter (before → after):**

```xml
<!-- Before -->
<Setter Property="BackColor" Value="{ThemeResource ButtonBackground}"/>

<!-- After -->
<Setter Property="BackColor" Value="{DynamicResource ButtonBackground}"/>
```

**Default style key (internal — reader may accept both one release):**

```text
Button.BaseStyle  →  {x:Type Button}   (same Style object, WPF key name)
```

**Anti-pattern → fix:**

```xml
<!-- Bad: semantic color on element -->
<Button BackColor="#008000" …/>

<!-- Good -->
<Button Style="{StaticResource ButtonSubmit}" …/>
```

### 2.17.5 Phasing

| Phase | Deliverable |
|-------|-------------|
| **3** | **`{DynamicResource}`**; theme dicts as **merged `ResourceDictionary`**; **`{ThemeResource}`** deprecated alias |
| **3** | Implicit style key **`{x:Type T}`** (+ **`T.BaseStyle`** shim); unify resource lookup with §2.16 |
| **6** | **`Style.Triggers`** / **`ControlTemplate`** for Button chrome (replace flat **`HoverColor`** widget fields where possible) |
| **7 / POS** | AI pass: **`ThemeResource` → `DynamicResource`**; fix hardcoded default styles; remove per-button color overrides |

### 2.17.6 Resolved decisions

- **No HTML `Class=` attribute** — use WPF **`Style`** + **`StaticResource`**.
- **`{ThemeResource}`** → **`{DynamicResource}`** (alias, then remove).
- **Semantic submit/cancel** → **named styles + theme brush keys** (keep POS pattern).
- **ThemesManager** → **theme dictionary merge/swap**, not a parallel resource universe.

### 2.17.7 Open questions (see §8)

*(None — styling/theming model locked for handoff.)*

---

## 2.18 XAML property resolution — end CallByName / widget fallback

**Question:** Original **`XAMLReader`** resolution (circa 2015): (1) dependency property by name, (2) standard property on control, (3) base control property, (4) **`cWidgetBase`** property. What is the **correct WPF-aligned fix** now that full DP registration is feasible (AI/codegen)?

**Decision:** **Single path — dependency properties only** for all XAML-settable attributes. **`cWidgetBase` is internal** (updated from **`PropertyChanged`** callbacks, never from XAML). **Remove `CallByName` fallback chain**; **fail loud** on unknown attributes. Consolidate registration on **`FrameworkElement`** + type-level **`DependencyPropertyRegistry`**.

### 2.18.1 What the code does today

**`XAMLReader.SetDependencyProperties`:**

```text
If Attached (contains ".")     → AttachedProperties dictionary
Else If DP.Exists(name)        → SetValue on DependencyProperty
Else                           → SetProperty (CallByName fallback)
```

**`SetProperty` fallback (`On Error Resume Next`):**

```text
Try CallByName on control
If fail And IWindow  → retry Base → Form → WidgetRoot
If fail Else         → retry Obj.Widget (cWidgetBase)
```

**Same pattern in `StyleManager.SetObjectProperties`** — styles can also land on widget fields bypassing DPs.

**Why it was done:** VB6 has **no implementation inheritance**; registering DPs per control is verbose; **`cWidgetBase`** already had **`BackColor`**, **`FontSize`**, etc.; hurry + laziness → let XAML “find” a setter anywhere.

**Examples of split storage (`Button.cls`):**

| XAML attribute | Lands on | Styles/bindings |
|----------------|----------|-----------------|
| `BackColor`, `Style`, `Command` | **Registered DP** | Yes |
| `DesignLeft`, `CornerRadius`, `GradientBackground` | **Private field + Property Let** | No — **`CallByName`** only |
| Some visual attrs if not registered | **`W` (cWidgetBase)** silently | No |

### 2.18.2 Why the fallback is wrong (WPF lens)

| Problem | Effect |
|---------|--------|
| **Silent `On Error Resume Next`** | Typos create **no error**, value may vanish or hit wrong target |
| **Dual storage** | Same logical property on DP **and** widget — precedence/style/bindings inconsistent |
| **Leaky abstraction** | XAML can set **implementation** (`cWidgetBase`) — bypasses **`SetCurrentValue`**, inheritance, **`AffectsMeasure`** |
| **Per-instance registration only** | Each control re-**`Register`s** in **`Class_Initialize`** — no type-level metadata (WPF **`Register`** once per type) |
| **Style + XAML diverge** | Style uses DP path when registered, **`CallByName`/widget** when not — unpredictable |
| **Blocks layout engine** | **`Design*`** not DPs → cannot drive Measure/Arrange or unified invalidation (§2.7, §3.3) |

**WPF rule:** Every XAML attribute maps to **`DependencyObject.SetValue(DependencyProperty, value)`** (or **`SetAttachedValue`**). CLR **`Property Let/Get`** are thin wrappers. Visual tree / render backend is updated in **`PropertyChangedCallback`** only.

### 2.18.3 Target architecture

**A. Type-level `DependencyPropertyRegistry` (VB6 substitute for static fields)**

```text
RegisterType("FrameworkElement", … core props …)
RegisterType("Button", extends "FrameworkElement", … button props …)

On control Class_Initialize:
  DependencyProperties.InitializeFromRegistry Me, "Button"
```

- One **`PropertyDefinition`** per property per type: type, default, metadata (**`IsInheritable`**, **`AffectsMeasure`**, **`AffectsRender`**, **`BindingMode`**).
- Optional **`ChangedCallback`** name → **`OnBackColorChanged(Control, New, Old)`** pushes to **`cWidgetBase`** internally.

**B. `FrameworkElement` property surface (shared across controls)**

Register once; all **`IUIElement`** controls inherit via registry merge:

- **Layout (replaces `Design*`):** `Width`, `Height`, `MinWidth`, `MaxWidth`, `Margin`, `HorizontalAlignment`, `VerticalAlignment`, `Visibility`
- **Visual (theme/style):** `BackColor`, `BorderColor`, `ForeColor`, `Opacity`/`Alpha`, `CornerRadius`, `BorderThickness`, `Padding`
- **Tree:** `DataContext`, `Style`, `Name`, `ToolTip`
- **Input (where shared):** `IsEnabled`, focus-related as needed

Control-specific: `Button` → `Command`, `CommandParameter`, `ClickMode`; `TextBox` → `Text`, `TextWrapping`; etc.

**C. `cWidgetBase` — render adapter only**

```text
XAML / Style / Binding  →  DP.SetValue / SetCurrentValue
                              ↓ PropertyChangedCallback (per type or shared)
                         cWidgetBase.W.BackColor = …   (internal only)
```

- **No XAMLReader path to `Obj.Widget`.**
- Widget property names **not** part of public XAML contract.

**D. XAMLReader — strict resolution**

```text
ResolveAttribute(element, name):
  If IsAttached(name)     → GetAttachedPropertyDefinition → SetAttachedValue
  Else
    def = GetPropertyDefinition(element.Type, name)
    If def Is Nothing     → Throw XamlParseException(element, name)
    Else                    SetValue(def, CoerceMarkupValue(value))
```

- **Delete `SetProperty` CallByName chain** (and duplicate in **`StyleManager`**).
- **`StyleManager`:** only iterate style setters against **registered DPs** on target type.
- Optional **one-release shim:** unknown attr tries legacy fallback + **debug warning** — not in release/CI.

**E. CLR property wrappers (AI/codegen-friendly)**

Each control keeps **`Public Property Get/Let BackColor`** for VB6 callers, but body is always:

```vb
' Pattern — no m_BackColor field on control
Public Property Let BackColor(ByVal Value As Long)
    DependencyProperties.SetValue "BackColor", Value
End Property
```

Remove duplicate **`m_CornerRadius`** fields where a DP exists. **`PropertyChanged`** handler (or per-property callback) updates Cairo/widget.

**F. Attached properties — real registry**

Replace ad-hoc **`AttachedProperties("Grid")`** dict with **`RegisterAttached("Grid.Column", …)`** + **`Grid.GetColumn(element)`** / **`SetColumn`** — WPF pattern for **`UniformGrid` → `Grid`**.

### 2.18.4 Phasing (coordinates §2.5, §3.3, Phase 1)

| Phase | Work |
|-------|------|
| **0** | **`DependencyPropertyRegistry`** module; **`InitializeFromRegistry`**; document full **XAML attribute → DP** table per type |
| **1** | **`FrameworkElement` registry** + layout DPs; **`Design*` → `Width`/`Height`/…** (breaking); callbacks sync widget |
| **1** | **Strict XAMLReader** (fail loud); remove widget fallback in framework builds |
| **2** | Migrate each control: register remaining XAML attrs; delete **`CallByName`** targets; codegen **Property Get/Let** from registry |
| **2** | **`StyleManager`** DP-only; **`AffectsMeasure`/`AffectsRender`** → **`InvalidateMeasure`/`InvalidateVisual`** |
| **4+** | Full precedence stack (triggers/template setters) on same DP store |

### 2.18.5 POS / migration tooling

1. **Scan** `pos-v1/UI/Resources/XAML/**/*.xml` — collect all attribute names per element type.
2. **Diff** against registry — gaps = DPs to add before dropping fallback.
3. **AI transform** — no mechanical change for attrs already on DPs; rename **`Design*`** per §3.3.
4. **CI contract test** — load golden XAML; unknown attribute must **throw**.

### 2.18.6 Resolved decisions

- **Do not** keep widget/base **`CallByName`** fallback in target framework.
- **Do not** expose **`cWidgetBase`** to XAML — adapter callbacks only.
- **Do** centralize registration (**`FrameworkElement`** + registry) — AI generates per-control boilerplate, not hand-copy **`Register`** blocks.
- **Do** align with WPF **`SetValue` / `SetCurrentValue` / `ClearValue`** public API (§2.5).

### 2.18.7 Open questions (see §8)

- [ ] **Registry + thin wrappers vs codegen `Implements` blocks** — same trade-off as §2.11 **`FrameworkElement` in VB6**; registry + codegen is recommended default here.

---

## 2.19 Memory use — bindings, DPs & process budget (&lt;100 MB observed)

**Symptom (refined):** VCF **object graph** is structurally heavy (per-instance DPs, binding overhead) — visible on Sales as a large share of process RAM. **Observed POS process:** **always &lt; 100 MB** in normal operation; **&gt; 100 MB only** when **secondary customer display** plays **video** (decode buffers — not VCF framework). Earlier **~70 MB** figure = Sales-heavy / UI slice estimate, not full-process ceiling.

**Verdict:** Memory is high due to **(1) object-per-instance dependency properties**, **(2) heavy per-binding object graph**, **(3) eager pre-built control grids with many `{Binding}`s**, **(4) deep widget duplication per control** — not Cairo alone. Fixes align with **§2.5 BindingExpression**, **§2.8 lazy inheritance**, **§2.18 registry**.

### 2.19.1 Where memory goes (code-level)

**A. Per-control dependency property instances (largest structural cost)**

Each **`Button`**, **`TextBlock`**, etc. calls **`DependencyProperties.Register`** in **`Class_Initialize`** (~9–15 properties per control).

Each **`DependencyProperty`** allocates:

- Its own **`PropertyChangedEvent`** (+ **`Register Me`**)
- **`DependencyPropertyMetadata`**
- **`m_Listeners`** + **`m_Callbacks`** (**`List`**)
- Value storage

WPF uses **static `DependencyProperty` fields per type** (metadata shared). VCF creates **full objects per control instance** → thousands of small COM objects on Sales screen.

**Rough scale (menu panel alone):**

| Area | Pre-allocated cells | Controls per cell | Notes |
|------|---------------------|-------------------|--------|
| Menu items grid | 7×3 = **21** | Button + 3 TextBlocks | **`MenuItemsGridButton.xml`**: **6 `{Binding}`** each |
| Submenu grid | 7×1 = **7** | Button + TextBlock | **5 bindings** |
| Menu groups | 2×8 = **16** | Button + TextBlock | **5 bindings** |
| **Menu subtotal** | **44 cells** | ~**110** VCF elements | ~**240+ Binding** objects |

Full Sales layout adds **LeftColumn/RightColumn**, **command grids** (~8 buttons × multiple fragments), **StatusBar**, **CustomerPanel**, etc. → **300–600+ VCF elements** and **400–800+ live bindings** is plausible before view-models and data.

**B. Per-binding object graph (why bindings hurt)**

Each **`{Binding …}`** creates a **`Binding`** object stored on **`control.Bindings`** (**`List`**) — **never detached** on normal screen life.

Per **`Binding`** (`Binding.cls` + **`BindingsManager`**):

| Object / link | Purpose |
|---------------|---------|
| **`Binding`** | Main binding |
| **`NestedProperty`** (+ child chain for `A.B.C` paths) | Path resolution |
| **`WithEvents`** on **`NestedProperty`**, **`SourcePropertyChangedEvent`**, **`TargetPropertyChangedEvent`** | 3 event sinks |
| **`TargetProperty.AddListener(Me)`** | Listener on target **DP** |
| **Default source = `DataContext` DP object** | **`BindingsManager`** sets **`Source = Target.DependencyProperties.GetProperty("DataContext")`** → **`SrcDepObj.AddCallback`** on **every** binding without explicit `Source` |

So **6 bindings on one menu button** → **6 callbacks** on that button’s **`DataContext` DP**, **6 listeners** on various target DPs, **6 `NestedProperty`**, **6 `Binding`** — while POS **also** updates the same cell via **`ButtonProperties.SetProperty`** in **`clsMenuPanelViewModel`** (dual path).

**C. Widget / visual duplication**

Each **`Button`**: **`cWidgetBase W`** + **`OverlayWidget`** in **`W.Widgets`**, plus **`TextBlock`** / **`Image`** children each with **own widget**. Menu item cell ≈ **5 widgets** × **21** ≈ **105 widgets** for items grid alone.

**D. Other framework buckets**

- **`XAMLResources`**: all XAML file **text** kept in **`Application.Resources("XAML")`** after load (§2.16)
- **`SetPanelContent`**: cached trees in **`Application.Resources`**
- **`Style`** objects with **`WithEvents ThemesManager`**
- **`ImageKey`** / Cairo surfaces (smaller than binding/DP overhead but non-zero)

### 2.19.2 Why Sales screen specifically

1. **`MainButtonsPanel`** → three **full UniformGrids** pre-filled with templated buttons (**`MenuItemsGridButton.xml`**, etc.) — **worst-case binding density**.
2. **Deep XAML tree** via **`res:`** fragments (§2.7) — many elements, each with full DP stack.
3. **Command buttons** — `{Binding Commands.…Command}` on dozens of chrome buttons.
4. **DataContext push** (§2.8) — inherited props copied to entire subtree at load/parent set.

Bindings are not the only cost, but **menu grid template + default DataContext-as-DP source** multiplies small objects fastest.

### 2.19.3 Improvements (framework — priority order)

**P0 — High impact, aligns with §2.5 / §2.18**

| # | Change | Effect |
|---|--------|--------|
| 1 | **`DependencyPropertyRegistry`** — **shared metadata per type**; per-instance **value store only** (array or compact map keyed by property index) | Cuts **PropertyChangedEvent × N controls × M props** — largest win |
| 2 | **`BindingExpression`** — **one object per bound DP**; replace **`NestedProperty` + triple `WithEvents` + listener on GetValue** | ~50–70% fewer objects per binding |
| 3 | **Resolve `DataContext` once per binding** — store **reference to source object**, not **`DependencyProperty` + AddCallback**; **rebind on context change** at element level (walk expressions) | Removes **K callbacks per control** (K = binding count) |
| 4 | **`BindingExpression.Detach`** on unload / **`DataContext` cleared**; clear **`control.Bindings`** | Allows GC when leaving Sales screen |
| 5 | **Lazy inheritance** (§2.8) — stop **`PassPropertyValue`** fan-out on load | Less transient allocation + CPU |

**P1 — Medium impact**

| # | Change | Effect |
|---|--------|--------|
| 6 | **Single `PropertyChanged` hub per `DependencyProperties`** instead of **per-`DependencyProperty` event object** | Fewer COM objects |
| 7 | **Remove `OverlayWidget` per Button** (already noted in `Button.cls`) | −1 widget per button |
| 8 | **Drop `XAMLResources` raw strings** after parse (§2.16 merged dicts) | −1–5 MB typical |
| 9 | **Default `OneTime`** for template bindings where source is **`ButtonProperties`** updated imperatively | Fewer live listeners (POS migration) |

**P2 — Architectural (POS + framework)**

| # | Change | Effect |
|---|--------|--------|
| 10 | **ItemsControl + virtualized/recycled rows** for menu grids — not 21× full button trees at load | Fewer controls when not visible |
| 11 | **Menu grid: code-behind set** (`Btn.Command = …`, `Txt.Text = …`) **or** one **`MultiBinding`/batch update** — drop 6 bindings × 21 cells | Large POS-side win short-term |
| 12 | **Shared command instances** — `{StaticResource AppCommands}` already; avoid per-button binding to nested path where **`CommandParameter` only** differs |

### 2.19.4 POS mitigations (until VCF P0 ships)

1. **Audit `MenuItemsGridButton.xml`** — use **`OneTime`** or remove bindings that **`clsMenuPanelViewModel` already sets** via **`ButtonProperties`**.
2. **Reduce TextBlock bindings** — parent sets text in code when loading items (3 bindings → 0 per cell).
3. **Ensure Sales view trees released** on navigate away (drop refs to **`SalesOrderView`**, clear cached **`Application.Resources`** panel keys if holding duplicate trees).
4. **Profile in Task Manager** before/after — baseline; optional **`Debug.Print TypeName` object counts** in **`SalesOrderView`** load for regression.

### 2.19.5 Phasing

| Phase | Deliverable |
|-------|-------------|
| **4** | **`BindingExpression`** + **DataContext rebind** + **Detach** (§2.5 P0) |
| **1** | **Registry / shared DP metadata** (§2.18) |
| **4** | **Lazy inheritance** (§2.8) |
| **7 / POS** | Menu grid binding reduction; **ItemsControl** long-term |

### 2.19.6 Resolved direction

Memory work is **same program** as DP/binding redesign — not a separate “optimization pass”. **Registry + BindingExpression + DataContext resolution fix** are prerequisites before micro-optimizing Cairo.

### 2.19.7 Expected savings & observed baseline (POS telemetry)

**Observed (product owner):**

| Scenario | Process RAM |
|----------|-------------|
| **Normal operation** (Sales and other screens) | **Always &lt; 100 MB** |
| **Secondary display + video** | **&gt; 100 MB** (expected — video decode/cache, not VCF UI graph) |

**Implication:** Memory is **not a crisis** on current POS hardware; framework rebuild is still warranted for **correctness, CPU (§2.7), leak resistance, and headroom** — not because the app is near OOM today.

**Scope of savings below:** **P0–P1 VCF fixes only** — **not** POS `MenuItemsGridButton` / menu VM changes.

**Recalibrated model** (full process **~80–95 MB** typical on Sales, **&lt; 100 MB** cap):

| Bucket | Est. share of &lt;100 MB | Notes |
|--------|--------------------------|--------|
| **VCF object graph** | **~25–40 MB** | Structural bloat target |
| **Cairo / widgets / static images** | **~8–15 MB** | |
| **POS data & VM** | **~15–25 MB** | Orders, menu, Codejock islands |
| **Runtime / DLLs** | **~10–20 MB** | VB6 + libs |
| **Customer display video** (when on) | **+20–80+ MB** | Separate budget; not framework |

**Framework P0 savings (registry + `BindingExpression` + DataContext + lazy inheritance):** roughly **~10–20 MB** absolute on full process (**~10–20%** of typical **&lt;100 MB**), not the larger figures from the earlier **70 MB = whole process** assumption.

**Framework P0 + P1:** add **~3–7 MB** → **~13–25 MB** total (**~15–25%**).

**Planning summary (full process):**

| Milestone | Typical process RAM | Notes |
|-----------|-------------------|--------|
| **Today** | **&lt; 100 MB** (normal) | Already acceptable vs WPF/Electron |
| **After VCF P0** | **~70–85 MB** | Structural fix; same XAML |
| **After VCF P0 + P1 + menu POS fix** | **~65–80 MB** | Optional further trim |
| **Secondary display video** | **&gt; 100 MB** | OK if documented; tune video pipeline separately |

**Product memory budget (recommended):**

- **Normal POS (no customer video):** **&lt; 100 MB** steady state — **already met**
- **After framework rebuild:** **&lt; 85 MB** typical Sales — **margin, not requirement**
- **Customer display + video:** **&lt; 150–200 MB** acceptable if decode is released on stop
- **No monotonic growth** over an 8-hour shift — **more important than peak MB**

### 2.19.8 Comparison vs WPF / modern frameworks (with observed baseline)

| | **Your POS (observed / target)** | **Typical WPF LOB** | **Electron** |
|--|----------------------------------|---------------------|--------------|
| **Full process, rich screen** | **&lt; 100 MB today**; **~65–85 MB** after fixes | **80–200+ MB** | **150–400+ MB** |
| **Runtime tax** | VB6 (no CLR/Chromium) | .NET CLR + WPF | Chromium |
| **Verdict** | **Already lighter** than most “modern” stacks | Baseline reference | Not suitable for same RAM budget |

The **&lt; 100 MB** observation means VCF+POS is **already competitive** on RAM. Framework rebuild **widens margin** and fixes **object-count architecture**; it does not change the product from “failing” to “passing” on memory.

**Confidence:** Absolute MB ranges are planning figures; **observed &lt;100 MB** is authoritative for normal ops. Re-measure after VCF P0 on same hardware/scenario.

---

## 2.20 North star — light + fast, full WPF feature set

**Goal (product owner):** Make the framework **as light and fast as possible**, **without sacrificing any feature**.

**Decision:** **No “lite mode” API.** The WPF-aligned surface in this document (bindings, styles/themes, templates, Selector, ItemsControl, resources, layout engine, ListView engine, etc.) remains **in scope**. Performance and memory are **first-class acceptance criteria** for every phase, implemented by **replacing inefficient internals**, not by cutting capabilities.

### 2.20.1 What “without sacrificing features” means

| In scope (keep / complete) | Out of scope (not a feature cut) |
|----------------------------|----------------------------------|
| Full **binding** grammar target (Path, Mode, Converter, Source, rebind on DataContext) | Long-term dual APIs (`Design*` + `Width`, `ThemeResource` + `DynamicResource` forever) |
| **Styles, themes, `{DynamicResource}`**, semantic styles | HTML-style `Class=` shortcut |
| **DataTemplate**, **ItemsControl**, **Selector**, **ListView** (merged) | POS `@`-fragment template API in framework |
| **ResourceDictionary**, **MergedDictionaries**, **`ResourceReference`** | Permanent `res:` app hack |
| **Measure/Arrange**, **Grid/StackPanel**, **Visibility/Collapsed** | Keeping **Design* scale cascade** as default |
| **ControlTemplate** / triggers (phased) | Per-control Cairo monoliths forever |
| POS migrates via **`MIGRATION.md`** — features move, not disappear | Silent XAML fallback chains |

**Rule:** If a change removes XAML or API capability POS uses today, it must have a **documented WPF-equivalent replacement** in the same release — not deferred “maybe later.”

### 2.20.2 Light — implementation pillars (memory & object count)

| Pillar | Section | Effect |
|--------|---------|--------|
| **`DependencyPropertyRegistry`** — shared metadata, instance value store only | §2.18 | Largest memory win; fewer COM objects |
| **`BindingExpression`** + **Detach** on unload | §2.5, §2.19 | One object per bound DP; GC-friendly |
| **DataContext → source object** (not DP + K callbacks) | §2.19 | Cuts listener/callback lists |
| **Drop parsed XAML strings**; merged dicts only | §2.16 | Lower baseline RAM |
| **Remove `OverlayWidget` per Button** | §2.19, §2.7 | −1 widget per button |
| **Strict XAML** — no silent `CallByName`/widget fallback | §2.18 | Prevents duplicate storage paths |
| **Unload views** — tear down trees when navigating away | §2.19 | Stable shift-long RAM |

**Budget (normal POS):** retain **&lt; 100 MB** process; target **&lt; 85 MB** typical after rebuild (§2.19.7) — **margin**, not a feature trade.

### 2.20.3 Fast — implementation pillars (CPU & latency)

| Pillar | Section | Effect |
|--------|---------|--------|
| **Measure/Arrange**; remove **Design* `MoveChild`** resize cascade | §2.7, §3 | Main UX win on nested Sales grids |
| **Coalesced `InvalidateMeasure` / `InvalidateVisual`** | §2.7 | No refresh storms on resize |
| **Lazy DP inheritance**; no **`PassPropertyValue`** fan-out | §2.8 | Faster load + context changes |
| **Push bindings** (`BindingExpression`) vs **listener-pull on `GetValue`** | §2.5 | Less work per property read |
| **ListView** bug fixes + bound path perf | §2.14 | Invoice/menu lists |
| **Simpler Border/Button paint**; text layout cache | §2.7 | Lower Cairo cost per frame |
| **Stable widget tree on arrange** — no **`Widgets.RemoveAll`** every resize | §2.7 | Less allocation churn |

**Primary user-visible metric:** Sales screen **resize / theme switch / menu page** feels instant on target POS hardware — not just lower Task Manager MB.

### 2.20.4 How VCF team proves “light + fast” (Phase 0 deliverable)

Add to **`.Tests`** (or POS contract app):

1. **Memory smoke** — load Sales-equivalent XAML tree; assert process **&lt; 100 MB** (normal), **no growth** after N navigations.
2. **Layout bench** — parent resize of nested UniformGrid/Grid; assert **&lt; X ms** (baseline TBD on reference HW).
3. **Binding bench** — N bindings update on INPC; assert **O(changed)** updates, not full-tree scan.
4. **Leak test** — open/close Sales view 50×; working set flat ± tolerance.

Ship **regression thresholds** in CI where feasible; document reference machine spec.

### 2.20.5 Priority when light/fast conflicts with purity

1. **Correct WPF semantics** (no silent wrong behavior)  
2. **Fast path for hot cases** (Sales grid, ListView, theme switch) — may use internal caches if externally equivalent  
3. **Memory** (object count, detach, registry)  
4. **API elegance** — never at cost of (1) or measurable (2)/(3)

**Anti-pattern rejected:** “POS doesn’t need `{Binding}` on menu cells” as a **framework** shortcut — that is a **POS template** choice (§2.19.4), not a capability removal from VCF.

---

## 2.21 Non-visual infrastructure — collections, MVVM primitives & services

**Question:** VCF is not only controls — **collections**, binding helpers, application/resources, value types, and utilities also need review for **light + fast + WPF-aligned** rebuild (§2.20).

**Decision:** Treat **data/collection/MVVM primitives** as **first-class framework surface** (same breaking-change + migration policy as controls). **Split optional non-UI utilities** from core UI DLL where practical. Optimize **allocation on collection change** and **unify notification** patterns.

### 2.21.1 Inventory (VCF today)

| Layer | Types | WPF analogue | Role |
|-------|-------|--------------|------|
| **Collections** | **`ObservableCollection`**, **`ObservableDictionary`**, **`List`**, **`UIElementCollection`** | `ObservableCollection<T>`, `ResourceDictionary` | ItemsSource, app/resources, children, event arg payloads |
| **Collection views** | **`CollectionViewSource`**, **`ListCollectionView`** | `CollectionViewSource`, `ICollectionView` | ListView current item; future ItemsControl |
| **Change events** | **`CollectionChangedEvent`**, **`CollectionChangedEventArgs`**, **`PropertyChangedEvent`** | `NotifyCollectionChangedEventArgs`, `INotifyPropertyChanged` | INPC + INCC |
| **Binding / MVVM** | **`Binding`**, **`BindingsManager`**, **`NestedProperty`**, **`SelfBinding`**, **`IValueConverter`**, **`ICommand`** | Same names | XAML `{Binding}`, commands |
| **Resources / markup** | **`Application`**, **`StaticResourceExtension`**, **`ThemeResource`**, **`MarkupExtensions`**, **`SolidColorBrush`**, **`String`** (XAML) | Application, markup extensions, brushes | Parse + lookup |
| **Styling (non-visual)** | **`Style`**, **`Setter`**, **`StyleManager`**, **`ThemesManager`**, **`XAMLStyleReader`** | Style, Setter | §2.17 |
| **Value types** | **`Thickness`**, **`Color`**, **`CornerRadius`**, **`KeyValuePair`** | Thickness, Color, struct-like setters | Layout/style values |
| **DP core** | **`DependencyProperty`**, **`DependencyProperties`**, **`DependencyPropertiesStatic`** | DependencyObject store | §2.5, §2.18 |
| **Utilities (non-WPF)** | **`Mail`**, **`INIParser`**, **`BackgroundWorker`**, **`StringProcessor`**, **`Information`**, **`Environment`**, **`NamingManager`** | N/A (app concerns) | Historical “kitchen sink” in **`Demac.VCF.dll`** |

**Storage backends:** collections wrap **vbRichClient `cArrayList` / `cCollection`** — fine; overhead is **VCF wrapper objects and per-change allocations**, not the list itself.

### 2.21.2 Collection stack — behavior & issues

**`ObservableCollection`**

- Backing store: **`cArrayList`**.
- Each **`Add`/`Remove`/`Replace`**: allocates **`List`** via **`NewList(...)`** + **`CollectionChangedEventArgs`** + notifies **`CollectionViewSource`** default view + raises **`CollectionChangedEvent`**.
- **`Class_Terminate`**: **`CollectionViewSource.DestroyDefaultView`** — good lifecycle hook.
- **WPF:** often uses a **single-element read-only list** for `NotifyCollectionChangedEventArgs`; does not allocate a new **`List`** class per item change.

**`ObservableDictionary`**

- Backing store: **`cCollection`** (keyed).
- Used for **`Application.Resources`**, **theme dicts**, **attached property** bags — **many instances**, moderate churn at load.
- Same **NewList + EventArgs** pattern on **`Add`/`Remove`**.
- **Target:** becomes **`ResourceDictionary`** base (§2.16, §2.17) — same type, stricter semantics.

**`List`**

- General-purpose **`cArrayList`** wrapper — used for **`control.Bindings`**, style setter enumeration, **payload inside `CollectionChangedEventArgs`**, temporary batches.
- **Issue:** **`List`** and **`ObservableCollection`** overlap; many **`List`** instances are **short-lived wrappers** created only to pass one item to an event.

**`UIElementCollection`**

- Wraps **`ObservableCollection`** + **`CollectionChanged`** forward — appropriate; keep, wire to **`Panel`** children.

**`CollectionViewSource` / `ListCollectionView`**

- **WPF-aligned concept** — default view per **`ObservableCollection`** (keyed by **`GetHashCode`** / pointer).
- **Known bugs (§2.14):** **`ListCollectionView.Initialize` static `bIsInitialized`** — only first view works; **`CollectionChangedActionMove`** stub in ListView path.
- **Gap vs WPF:** no **Sort/Filter/Group** (OK defer); **`CurrentItem`** not exposed on **`Selector`** yet (§2.13).

### 2.21.3 Event object proliferation (ties to §2.19)

| Pattern today | Count driver | Target |
|---------------|--------------|--------|
| **`PropertyChangedEvent` per `DependencyProperty`** | controls × DPs | **Hub on `DependencyProperties`** or registry (§2.18) |
| **`PropertyChangedEvent` per INPC source** | each ViewModel / ButtonProperties | Keep **one per source** (WPF-like) |
| **`CollectionChangedEvent` per `ObservableCollection`** | each collection | Keep **one per collection**; lighten args |
| **`CollectionChangedEventArgs` + `List` per change** | every Add/Remove on hot paths | **Reuse / single-item payload** (§2.21.5) |

### 2.21.4 WPF alignment targets (non-visual)

| Area | Target | Phase |
|------|--------|-------|
| **Collections** | **`ObservableCollection(Of T)`** semantics in docs; accept **`IEnumerable`** snapshot adapter for ItemsSource (§2.14) | 4 |
| **Views** | Fix **`ListCollectionView`** init; expose via **`Selector`** + **`ICollectionView`** subset | 4–5 |
| **Resources** | **`ResourceDictionary : ObservableDictionary`** merge + **`Source=`** (§2.16) | 3 |
| **Brushes / tokens** | **`SolidColorBrush`**, theme keys — **`DynamicResource`** (§2.17) | 3 |
| **Commands** | **`ICommand`** + **`CanExecuteChanged`** (WPF parity) | 4 |
| **Markup** | One **`MarkupExtension`** base; **`StaticResource`**, **`DynamicResource`**, **`Binding`** | 3–4 |
| **Value types** | **`Thickness`**, **`CornerRadius`** — immutable, parse once; optional **freeze** | 1–2 |
| **Utilities** | **`Mail`**, **`INIParser`**, etc. — **move to `Demac.VCF.Core` or app**; not required for UI parse/bind | 0 doc split |

**Features retained:** INCC, INPC, **`CollectionView`**, resource lookup, converters, commands — **implement lighter**, not removed.

### 2.21.5 Light + fast improvements (collections & events)

**P0 (framework rebuild)**

1. **`CollectionChangedEventArgs` lightweight path** — for single-item Add/Remove, pass **item + index** without **`New List` + Add**; optional **args pooling** for bulk updates.
2. **Fix `ListCollectionView.Initialize`** — **per-instance** flag; add regression test.
3. **Implement `CollectionChangedActionMove`** on ListView/ItemsControl path (§2.14.8).
4. **`BindingExpression` list** — replace **`control.Bindings As List`** with typed **`BindingExpressionCollection`** (detach all on unload).

**P1**

5. **Batch notifications** — **`ObservableCollection`**: **`BeginUpdate`/`EndUpdate`** (WPF **`DeferRefresh`**) for menu/order bulk loads.
6. **Dictionary lookup perf** — resources/themes: **`ResourceDictionary`** keyed lookup; avoid linear scan where **`TryFindResource`** walks merge stack naively.
7. **Read-only wrappers** — expose **`ReadOnlyObservableCollection`** for bound views without copy.

**P2**

8. **Sort/Filter** on **`ListCollectionView`** only if POS needs (invoice grid columns later).

### 2.21.6 `List` vs `ObservableCollection` — guidance

| Use | Type |
|-----|------|
| **ItemsSource**, children that notify UI | **`ObservableCollection`** |
| **Application.Resources**, merged dicts | **`ResourceDictionary`** (evolved **`ObservableDictionary`**) |
| **Internal batch/build** | **`List`** or array — **do not** raise INCC |
| **Event arg payloads** | **Lightweight args** — not long-lived **`List`** |

Avoid creating **`ObservableCollection`** for static XAML children that never change after load — **UIElementCollection** can use **non-observable** internal storage until **structural change** (future optimization; optional).

### 2.21.7 Utilities split (optional, Phase 0 doc)

**Problem:** **`Mail`**, **`INIParser`**, **`BackgroundWorker`** live in the same DLL as **`Button`** — pulls unrelated surface into “framework” and confuses WPF alignment scope.

**Proposal:**

- **`Demac.VCF`** — UI, XAML, collections, binding, resources, controls.
- **`Demac.VCF.Core`** or **`Demac.Common`** (existing?) — **`INIParser`**, **`Mail`**, async helpers — POS already uses some elsewhere.

**Not urgent for RAM** — mainly **clarity** and smaller conceptual API. Breaking only if types move namespaces; document in **`MIGRATION.md`** if done.

### 2.21.8 Acceptance tests (add to §2.20.4)

- **Collection churn:** 1000× **`Add`** on **`ObservableCollection`** — assert **bounded** object growth vs today (profile before/after args lightweight path).
- **Multi-list views:** two **`ObservableCollection`** instances each with **`ListCollectionView`** — both **`CurrentItem`** work (static init bug).
- **Resource lookup:** 1000× **`TryFindResource`** on merged dict — **&lt; X ms**.

### 2.21.9 Open questions (see §8)

- [ ] **Split utilities DLL** — Phase 0 doc-only vs actual move in first breaking release?
- [ ] **`BeginUpdate/EndUpdate`** on **`ObservableCollection`** — required for POS menu bulk load, or nice-to-have?
- [ ] **`ReadOnlyObservableCollection`** — expose in public API or internal only?

---

## 3. Layout: today vs WPF target

### 3.1 Today

- Each element stores **design-space** rect (`Design*`).
- On parent resize: **multiply** child coords by `parentActual / parentDesign` (zoom entire subtree).
- **UniformGrid:** cell layout by child order + `Grid.ColumnSpan`/`RowSpan` attached props (not row/column index).

### 3.2 Target (WPF-like)

- **Measure → Arrange** pipeline; **`ActualWidth/Height/Left/Top`**.
- **`Width` / `Height`:** pixels, **`Auto`**, **`*`** (in Grid).
- **`Margin`, `Padding`, `HorizontalAlignment`, `VerticalAlignment`** drive arrange (string enums: Stretch, Center, …).
- **Panels:** `Grid`, `StackPanel`, `DockPanel`, `Border`, `Canvas` (absolute coords **only** on Canvas).
- **Legacy:** root `LayoutMode=ScaleToDesignSurface` or **`Viewbox`** for old 1024×768 screens — **one** scale at root, not per child.

### 3.3 Layout property names — breaking change (agreed direction)

**Decision:** Public XAML and API use **WPF names only** — `Width`, `Height`, `Left`, `Top`, `Margin`, alignments — registered as **dependency properties** on a shared element base. **Remove `Design*`** in the same major release (no indefinite dual names).

**Rationale:** Small consumer surface + AI-assisted POS migration makes a **clean break** cheaper than maintaining two parallel property systems.

**VCF team deliverables (per breaking release):**

- **`BREAKING_CHANGES.md`** — what changed, why, semver bump
- **`MIGRATION.md`** — before/after examples (XAML + VB6); mechanical find/replace patterns where applicable
- **Changelog entry** with linked migration section

**Optional short-term:** A **one-release** XAML reader shim that maps `DesignWidth` → `Width` with a **deprecation warning** is acceptable only if it simplifies POS cutover timing — not a permanent dual API.

**Note:** Until Measure/Arrange lands, document semantic differences (e.g. scale-at-root legacy mode vs true WPF arrange) in the migration guide — not by keeping duplicate property names.

---

## 4. WPF alignment program (draft roadmap for VCF)

**North star:** WPF-experienced developers can write VCF with minimal surprises; **semantic parity** for LOB UI (not full WPF feature parity). **As light and fast as possible** without removing targeted features — §2.20. Every phase lists perf/memory acceptance criteria; **`.Tests`** benches in Phase 0.

### Phase 0 — Spec, tests & migration contract

- `VCF_XAML_WPF_SUBSET.md` in VCF repo
- **`VCF_INFRASTRUCTURE.md`** — collections, views, events, resources, utilities split (§2.21)
- **`VCF_LISTVIEW_ARCHITECTURE.md`** — ItemsSource / DataTemplate / ListViewBase / vbWidgets lineage (§2.14)
- Golden XAML layout tests in `.Tests`; **ListView** bind/template/selection/scroll tests
- **Semver + `BREAKING_CHANGES.md` + `MIGRATION.md`** for every major release that changes XAML/API surface
- Document **breaking-change policy** (small internal consumer base; migration guides required)

### Phase 1 — Layout core

- **`FrameworkElement` base** — start compositional stack (§2.11); **`DependencyPropertyRegistry`** + shared DP set (§2.18)
- **`Visibility` DP** (Visible/Hidden/Collapsed) on all elements; fix Collapsed layout — §2.15
- Measure/Arrange; Actual*; Width/Height/Margin/Alignment as DPs
- **Remove public `Design*`** (breaking); optional one-release XAML reader alias with deprecation warning only
- Root **`LayoutMode=ScaleToDesignSurface`** or **Viewbox** for unmigrated screens until POS XAML is converted
- **Performance (§2.7):** stable widget tree on arrange; delete per-container Design* `MoveChild`; coalesced invalidation

### Phase 2 — Panels

- Grid (Row/Column defs, `*`, Auto), **`StackPanel`** (Vertical/Horizontal — §2.15.4), DockPanel, **Border as child decorator**, Canvas

### Phase 2b — Content model

- **ContentControl**, **ContentPresenter**; refactor **Button** onto composition stack (§2.10, §2.11)

### Phase 3 — XAML & resources

- xmlns mapping; ResourceDictionary + MergedDictionaries; **`res:` loader in framework** (see §2.6); **`{DynamicResource}`** + theme dict merge (§2.17); stricter parse/create errors

### Phase 4 — Bindings & MVVM

- **Dependency property P0:** `BindingExpression`, DataContext rebind, public `ClearValue`, **lazy inheritance** (see §2.5, §2.8)
- **Collections P0:** lightweight **`CollectionChangedEventArgs`**, fix **`ListCollectionView`** static init, **`Move`** handler (§2.21)
- **TextBox TwoWay:** `UpdateSourceDelay` + flush on Enter/LostFocus (§2.9)
- DataContext rebind; RelativeSource/ElementName (minimal — **required for commands in DataTemplates**, §2.12.8)
- **`ItemsControl` + `ItemsPanelTemplate`** — WPF-aligned; unify **`DataTemplate`** engine with ListView (§2.12.8)
- **`Selector`** — `SelectedItem`, `SelectedIndex`, `SelectedValue`, `SelectedValuePath`; refactor **ListView** (§2.13)
- ListView bound performance fixes

### Phase 5 — Controls

- **Button `Content` (string caption)** — see §2.10; high POS value, low risk
- **`ListBox`**, **`ComboBox`**, **`TabControl`** on **`Selector`** base (§2.13)
- CheckBox, RadioButton, ScrollViewer, … (POS priority order)

### Phase 6 — Styles & templates

- **ControlTemplate** — lookless Button/TextBox/… built from Border/TextBlock/Grid (§2.11)
- **`Style.Triggers`** / PropertyTrigger / DataTrigger (subset) — replace flat widget **`HoverColor`** fields where possible (§2.17)
- **Render optimization:** simplified Border/Button paint; text layout cache (§2.7 P0/P2)

### Phase 7 — Tooling & migration

- Designer/XAMLWriter (if feasible)
- **Migration guides** maintained in VCF repo; POS-specific migration checklist in denovo (this doc + screen matrix)
- **denovo:** AI-assisted bulk migration of XAML + VB6 after each tagged VCF release (Cursor)

### Out of scope (explicit)

BAML, FlowDocument, 3D, full animation storyboards, pixel-identical WPF layout rounding, full virtualization v1.

---

## 5. POS priorities for VCF (evidence-based)

Ordered for **Front completion** then **Back Office**:

| P | Item | Why |
|---|------|-----|
| P0 | **ListViewBase evolution + merged ListView** | Replace Codejock InvoiceGrid — read-only hierarchical multi-column list §6 |
| P0 | **ListView P0 bug fixes** | `ListCollectionView` static init, Move handler, ItemTemplate stubs — §2.14.8 |
| P0 | **DataContext → rebind bindings** | Every major view has TODO; screen switching |
| P0 | **Measure/Arrange + Grid** | WPF layout; end scale-factor UX |
| P1 | **`Visibility` / `Collapsed` (WPF semantics)** | Unified DP; layout-aware; replace `Visible` — §2.15 |
| P1 | **`StackPanel`** | Vertical/horizontal stacks; reduce UniformGrid hacks — §2.15.4 |
| P0 | **Selector DPs on ListView** | `SelectedItem` / `SelectedIndex` in XAML — Login, dialogs, invoice grid — §2.13 |
| P1 | **DP inheritance perf** | Push `DataContext` / `PassPropertyValue` on large trees — §2.8 |
| P1 | **ItemsControl + DataTemplate (WPF)** | Replace MessageBox/Dialog `@`-fragments — §2.12.8 |
| P1 | **Render/layout perf (nested grids)** | Border paint, RemoveAll on resize, Design* cascade — §2.7 |
| P1 | **Numeric keypad (`Numpad`)** | Replace numpad .frm + SendKeys — §5.2 B1 |
| P1 | **ListView bound performance** | Login, Back Office lists |
| P1 | **`TabControl`, `CheckBox`, `ComboBox`** | Settings, Z, CRUD `.frm` bulk — §5.2 C1–C3 |
| P2 | **`RadioButton`, `DatePicker`, `ScrollViewer`** | Smaller `.frm` surface — §5.2 D |
| P2 | **Button `Content` / caption** | Remove nested TextBlock boilerplate — §2.10 |
| P2 | **ViewBase** (reduce IUserControl boilerplate) | Migration velocity |
| P3 | **Viewbox / legacy layout mode** | Until all POS XAML migrated |

**Not asking VCF to own:** POS business rules, KernelLib, DeNovo XAML content (except samples), payment/printing.

### 5.1 VCF control inventory vs full POS conversion

**VCF ships today:** `Button`, `TextBox`, `TextBlock`, `Border`, `Panel`, `UniformGrid`, `Window`, `UserControl`, `ListView`, `UnboundListView` (→ merge), `Image`, internal `ScrollBar`, `WindowsFormsHost` (escape hatch).

**Not in VCF:** `CheckBox`, `RadioButton`, `ComboBox`, `TabControl`, `ScrollViewer`, `ProgressBar`, `DatePicker`, `Label` (use `TextBlock`), dedicated **`Numpad`**.

**Front XAML today (~59 screen XML files):** Uses only **Button, TextBlock, TextBox, Border, Panel, UniformGrid, Image, ListView** — **zero** ComboBox/CheckBox/RadioButton/TabControl in XAML. Front is blocked by **InvoiceGrid island** + **~40 modal `.frm` dialogs**, not missing checkboxes on the sales canvas.

**Legacy `.frm` inventory (`pos-v1/UI/Source/Forms`, ~68 forms):**

| Codejock / VB6 control | Approx. footprint | Role |
|------------------------|-------------------|------|
| **`ReportControl`** | **~30 forms** | Read-only **multi-column lists/reports** (Z, day orders, customers, stock, queries) — same engine need as InvoiceGrid §6 |
| **`TabControl`** | **5 forms** (heavy: `frm_settings`, `frm_z2`, …) | Settings, Z report tabs |
| **`CheckBox` / `RadioButton`** | **~10 forms** | Filters, receipt options, search settings |
| **`ComboBox`** | **~3 forms** in UI/Source (+ **UILib** CRUD) | Dropdowns in settings, customer edit |
| **`DatePicker`** | **~3 forms** | Cheques, Z reports |

### 5.2 What must be implemented for **complete** POS UI → VCF

Grouped by **blocking impact**, not control alphabet.

#### Tier A — P0 (Front “done” + any bound list)

| # | Deliverable | Type | Why |
|---|-------------|------|-----|
| A1 | **Merged `ListView` + `ListViewBase` evolution** | List engine | **InvoiceGrid** + read-only column lists — §6; largest Front blocker |
| A2 | **`Selector`** (`SelectedItem`, …) | List API | Login, invoice line, dialogs |
| A3 | **Binding P0** (DataContext rebind, `BindingExpression`) | Framework | All MVVM screens |
| A4 | **Measure/Arrange + `Grid`** | Layout | Replace Design* / UniformGrid perf issues |
| A5 | **ListView P0 bug fixes** | List engine | §2.14.8 — multi-list apps unsafe today |

#### Tier B — P1 (modal / operator workflows)

| # | Deliverable | Replaces | Notes |
|---|-------------|----------|-------|
| B1 | **`Numpad` / numeric keypad** (VCF or documented POS `UserControl`) | `frmNumpad`, `PanelNumpad2`, SendKeys | Discount/qty/tender entry on Front |
| B2 | **Modal `Window` + dialog patterns** (already exists; polish) | Many small `.frm` | MessageBox/DialogWindow pattern §2.12 |
| B3 | **`ListView` bound mode perf** | Login, pick lists | ItemSource slowness §UI baseline |
| B4 | **`ItemsControl` + `DataTemplate`** | MessageBox `@` fragments | §2.12.8 |

#### Tier C — P1/P2 (convert settings, Z, Back Office `.frm` bulk)

| # | Control | WPF priority | POS evidence |
|---|---------|--------------|--------------|
| C1 | **`TabControl`** | High | `frm_settings`, `frm_z2`, `frm_query`, UILib product editor |
| C2 | **`CheckBox`** | High | Settings, filters, `frm_receipt`, data lists |
| C3 | **`ComboBox`** | High | Settings, customer edit, UILib CRUD (category, VAT, supplier, …) |
| C4 | **`RadioButton`** | Medium | `frm_search_settings`, option groups (often few — can defer if tabs/checkboxes suffice) |

Implement on **`Selector`** / **`ContentControl`** stack where applicable (§2.13, §2.11).

#### Tier D — P2 (secondary)

| # | Control | Replaces |
|---|---------|----------|
| D1 | **`DatePicker`** (or `TextBox` + mask interim) | `frm_cheque`, Z report date fields |
| D2 | **`ScrollViewer`** | Long tab pages, scrollable settings |
| D3 | **`ProgressBar`** | Rare (unload progress) |
| D4 | **`Button.Content`** | XAML verbosity §2.10 |

#### Tier E — Not a “control” but required for parity

| Item | Notes |
|------|--------|
| **Report/list grids → `ListView`** | Same **Tier A** engine; ~30 forms use Codejock **ReportControl** for read-only columnar data — **not** a separate DataGrid product |
| **`WindowsFormsHost`** | Short-term bridge only — discouraged for “complete” claim |
| **UILib forms** | Same C1–C4 controls in `pos-v1/Libs/UILib/Forms` — included in full conversion |

### 5.3 Minimum control set summary (answer)

To claim **complete POS UI on VCF**, implement in this order:

1. **List engine** (merged **`ListView`**, hierarchical multi-column, **`MeasureRow`**) — covers **InvoiceGrid + most ReportControl screens**
2. **Framework binding + layout + Selector** (Tier A3–A4)
3. **`Numpad`** (Tier B1)
4. **`TabControl`, `CheckBox`, `ComboBox`** (Tier C — settings/Z/CRUD)
5. **`RadioButton`, `DatePicker`, `ScrollViewer`** (Tier D — smaller surface)

**CheckBox / RadioButton / ComboBox are not Front-sales XAML blockers today**, but **are mandatory** for converting the remaining **~50+ classic forms** and UILib. **`TabControl`** is the highest-impact **missing input chrome** after the list engine.

---

## 6. InvoiceGrid (Codejock) vs VCF ListView

**Deep dive:** §2.14 (ListView stack, ItemsSource, DataTemplate, **`ListViewBase` = engine**).

**Target:** Merged **`ListView`** on evolved **`ListViewBase`** — not a separate **`UnboundListView`** type (§2.14.10).

### 6.1 What POS InvoiceGrid actually is (clarified)

**Not a DataGrid.** Codejock **`XtremeReportControl.ReportControl`** on `FormMain.frm` is used as a **read-only, multi-column list** with **parent/child rows** and **variable row height**. There is **no in-cell editing** in the POS invoice UI — lines come from the sales order model; user actions (qty, modifiers, etc.) happen elsewhere in the UI.

**Evidence (POS today):**

| Behavior | Implementation |
|----------|----------------|
| **Columns** | 16 columns (DESCRIPTION, QTY, PRICE, …) — `InvoiceGridHelper.Initialize` |
| **Parent/child** | Parent line + modifier/child lines via `AddRecordEx2(ChildRecord, ParentRecord)` — `FormMain.ConstructItemList` |
| **Variable height** | `InvoiceGrid_MeasureRow`: parent **40px**, child **20px**; `FixedRowHeight = False` |
| **Row styling** | `BeforeDrawRow`: bold/larger font on parent, smaller on child (`ParentRecord Is Nothing`) |
| **Selection** | `FocusedRow` / highlight colors — maps to VCF **`SelectedItem`** |
| **Editing** | **None** — sort, reorder, column resize disabled in helper |

**WPF analog:** Closer to **`ListView` + `GridView` columns** (or custom column layout) with **hierarchical items** — **not** `DataGrid` / editable cells. VCF target name **`ListView`** is correct; avoid designing a general-purpose editable grid for this migration.

### 6.2 VCF engine requirements (InvoiceGrid parity)

| Requirement | VCF `ListViewBase` / merged `ListView` |
|-------------|----------------------------------------|
| **Multi-column rows** | Use/extend existing **column** model in `ListViewBase` (headers optional) |
| **Variable row height** | **`MeasureRow(RowIndex, width) As Long`** (or WPF-like row height callback) |
| **Parent/child display** | **Row level** / indent on flat index, or hierarchical **`ItemsSource`** + **`HierarchicalDataTemplate`** (Phase 5+ decision) |
| **Read-only display** | Owner-draw or **`DataTemplate`** per row — **no TextBox-in-cell** |
| **Selection** | **`Selector.SelectedItem`** bound to active line VM |
| **Populate from order** | POS **`ItemsSource`** = flattened/hierarchical collection of line VMs |

**Explicitly out of scope for InvoiceGrid replacement:**

- In-cell editing, commit/cancel per cell
- Full **`DataGrid`** control
- Codejock record/field mutation API on the control

### 6.3 Migration mapping

| Codejock today | VCF target |
|----------------|------------|
| `Records` + `AddRecordEx2` parent/child | **`ObservableCollection`** of line VMs with **`ParentId`** / level, or tree items |
| `MeasureRow` | **`ListView.MeasureRow`** event or engine callback |
| `BeforeDrawRow` | **`DataTemplate`** by level/type, or **`OwnerDrawItem`** metrics |
| `FocusedRow.Record` | **`SelectedItem`** `{Binding ActiveLine}` |
| `Populate` / `DeleteAll` | **`ItemsSource`** replace/clear on order VM |
| Column settings string | Optional: column width DP persistence (P2) |

**Path (agreed):** Refactor or **new `ListViewBase` implementation** in VCF; POS supplies **`SalesOrderItem`** line VMs and templates/draw — **not** a POS-specific grid control in denovo.

---

## 7. Working model (two repos)

| VCF repo | denovo repo |
|----------|-------------|
| Framework code, `.Tests`, spec | XAML, ViewModels, `ObjectConstructor`, migration |
| Tag releases; golden tests; **`BREAKING_CHANGES.md` / `MIGRATION.md`** | Pin min VCF version; integration smoke after each tag |
| POS sample or contract test app | **AI-assisted migration** of XAML + source; this document + screen matrix (TBD) |

**Integration smoke:** login → sales → line grid → one dialog → commit path.

---

## 8. Open questions

*(Add items here as we discuss; strike through when resolved and move to §9.)*

- [ ] **Grid vs extend UniformGrid** — enough for POS if Row/Column index added, or full Grid required?
- ~~**`ThemeResource`**~~ — **resolved:** **`{DynamicResource}`**; **`{ThemeResource}`** deprecated alias; theme dicts as merged **`ResourceDictionary`** — §2.17
- ~~**`res:` vs `MergedDictionaries`**~~ — **resolved:** §2.16
- [ ] **ViewBase in VCF** — see §2.6.7
- [ ] **Strict XAML create failures** — see §2.6.7
- [ ] **Legacy layout default** — new major version defaults to **Arrange**; opt-in `ScaleToDesignSurface` only for screens not yet migrated?
- [ ] **One-release XAML alias** for `Design*` → WPF names — yes/no, or hard break only?
- ~~**Invoice grid** — framework (ListViewBase) vs POS-specific control in denovo?~~ — **resolved:** **VCF `ListViewBase` engine** (evolved/rewritten) + merged **`ListView`** in framework; POS owns line VMs/adapters only — §2.14.10
- [ ] **BindingExpression vs listener-pull** — see §2.5.7
- [ ] **Split utilities DLL** (`Mail`, `INIParser`, …) — §2.21.7
- [ ] **`ObservableCollection BeginUpdate/EndUpdate`** — required for POS? §2.21.5
- ~~**Design* → DP migration**~~ — **resolved:** one breaking release; WPF names only; see §3.3, §9
- [ ] **Z report / Codejock island** — temporary exception vs must be VCF for “done”?
- [ ] **ANSI/Greek (Windows-1253)** — document at VCF string boundary; any API needed?
- [ ] **Compositional phasing** — Button `Content` string before full ContentControl/ControlTemplate? see §2.11.7
- [ ] **`FrameworkElement` in VB6** — one base class + thin wrappers vs codegen? see §2.11.7
- [ ] **Border refactor timing** — decorator model + breaking visual change in same release as layout engine? see §2.11.7
- ~~**`@`-fragment templates vs binding-only `DataTemplate`**~~ — **resolved:** WPF-aligned **binding-only**; no `@` in framework; migrate POS MessageBox/DialogWindow in Phase 7 — §2.12.8, §9
- [ ] **ListView owner-draw vs ItemsControl full tree** — see §2.12.9
- [ ] **Command binding inside DataTemplate** — RelativeSource timing — see §2.12.9
- [ ] **`ListBox` vs ListView-only** for dialog grids — see §2.13.9
- [ ] **`ListIndex` → `SelectedIndex` migration** — alias vs hard break — see §2.13.9
- ~~**UnboundListView Selector surface**~~ — **resolved:** merged **`ListView`** exposes **`Selector`** for both bound and owner-draw — §2.14.10
- ~~**Merge ListView + UnboundListView**~~ — **resolved:** merge — §2.14.10
- ~~**Vendor vbWidgets / ListViewBase relationship**~~ — **resolved:** **`ListViewBase` is the engine**; refactor or new impl — §2.14.6, §2.14.10
- [ ] **Interim StackPanel before Grid** — see §2.15.6
- [ ] **POS `Visibility="0"` on Panel** — audit LeftColumn — see §2.15.6

---

## 9. Agreed direction

*(Stable decisions only — moved from discussion log when confirmed.)*

1. **VCF remains a separate repo** — not vendored into denovo except DLL reference.
2. **Strategic goal:** Framework **as close to WPF as practical** (semantic parity for LOB UI).
3. **Consumer scope:** POS + small internal apps only — **not** a general-purpose public framework.
4. **Breaking changes are OK** when they improve design, efficiency, and WPF alignment — **required:** VCF team documents each breaking release (`BREAKING_CHANGES.md`, `MIGRATION.md`, changelog).
5. **denovo migration:** Use **AI assistance (Cursor)** to migrate POS XAML and VB6 source after VCF releases — favors clean API over long dual-support.
6. **Replace scale-factor child layout** with **Measure/Arrange + Grid/StackPanel** as the default model; **legacy scale mode** only at root for unmigrated screens.
7. **Layout/XAML names:** **WPF names only** (`Width`, `Height`, …); **remove `Design*`** in the same major release as DPs + layout engine (see §3.3).
8. **Compositional architecture:** Complex controls (**Button**, **Border**, lists) should be built from **simple elements + panels + templates**, WPF-style (§2.11) — not one-off Cairo in each control long-term.
9. **Templates:** Align with **WPF** — `DataTemplate`, `ItemsControl`, `ItemsPanelTemplate`, then `ControlTemplate` / `ContentTemplate` (§2.12.8). **No** framework `@`-fragment API; POS `@` templates are legacy to migrate away.
10. **Selector:** WPF-aligned **`SelectedItem`**, **`SelectedIndex`**, **`SelectedValue`**, **`SelectedValuePath`** on a **`Selector`** base; **ListView** and future list controls expose these in XAML (§2.13). Deprecate public **`ListIndex`** on bound controls.
11. **ListView stack:** **`ListViewBase` / `TextBoxBase` = vbWidgets-derived engines** in VCF; public controls are thin wrappers. **Merge `ListView` + `UnboundListView`** into one **`Selector`**-based **`ListView`**. Evolve engine (**refactor or new implementation**) for InvoiceGrid — not a separate control family (§2.14). Fix P0 bugs (§2.14.8).
12. **Visibility + panels:** WPF **`Visibility`** (Visible/Hidden/**Collapsed**) as unified DP on **`FrameworkElement`** with **layout-aware Collapsed**; deprecate bool **`Visible`**. Implement **`StackPanel`** (Phase 2) — §2.15.
13. **`res:` includes:** **`ResourceDictionary` + `MergedDictionaries` + `Source=`** + **`{StaticResource}`**; move loader into VCF; **`res:` XML namespace legacy** (transitional shim); dynamic panels → **`ContentControl`** + binding; command fragments → **`DataTemplate` + `ResourceReference` clone** (not ContentControl-per-cell); explicit **`x:Key`** required — §2.16.
14. **Styles & themes:** WPF **`Style`** + **`Setter`** + **`BasedOn`**; named semantic styles (`ButtonSubmit`, …) via **`Style="{StaticResource …}"`** — **no HTML `Class=`**; theme palettes in per-theme **`ResourceDictionary`**; **`{ThemeResource}` → `{DynamicResource}`** — §2.17.
15. **XAML property resolution:** **DP-only** XAML setters; **`DependencyPropertyRegistry`** + **`FrameworkElement`** shared props; **`cWidgetBase` internal** via change callbacks; **remove CallByName/widget fallback** — §2.18.
17. **Memory / bindings:** Shared **DP registry** + **`BindingExpression`** + fix **DataContext-as-DP callback fan-out** — §2.19.
18. **Light + fast, full features:** No lite-mode API; optimize internals (registry, push bindings, Measure/Arrange, coalesced paint, detach); WPF feature set in this doc stays in scope; **`.Tests`** perf/memory benches — §2.20.
19. **Non-visual infrastructure:** Collections, collection views, change events, resources, value types — lightweight **`CollectionChangedEventArgs`**, fix **`ListCollectionView`**, **`ResourceDictionary`** evolution; optional utilities split — §2.21.
20. **Handoff to VCF team** will be sent **after** discussion topics in §8 are covered (or explicitly deferred).

---

## 10. Final handoff draft (for VCF team)

**Status:** Draft — ready for VCF team review; §8 open items may be explicitly deferred in kickoff meeting.

**Full rewrite blueprint:** [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md) (master catalog, per-type actions, bug registry, test matrix).

### Subject

**Demac.VCF — WPF alignment program: rewrite specification and POS-driven priorities**

### Executive summary

Demac.VCF should evolve into a **WPF-semantics LOB UI framework** for VB6/Cairo: dependency properties, bindings, styles, resources, templates, Selector, Measure/Arrange — **without sacrificing features** for performance. Implementation should be **as light and fast as practical** via registry-based DPs, `BindingExpression`, coalesced layout/paint, and lightweight collection notifications (not a “lite” API).

**Consumer base:** POS (DeNovo) + a few internal Demac apps. **Breaking changes are acceptable** when they improve design and WPF parity, provided each major release ships **`BREAKING_CHANGES.md`**, **`MIGRATION.md`**, and semver. DeNovo will use **AI-assisted migration** after VCF releases.

**VCF remains in its own repo;** denovo hosts coordination docs and will pin DLL versions per release.

### Non-negotiable outcomes

1. **WPF semantic parity** for LOB UI (layout, DPs, bindings, styles, resources, templates, selection) — see §2 and [rewrite spec §1–§2](./VCF_FRAMEWORK_REWRITE_SPEC.md).
2. **Fail-loud XAML** — no silent `Nothing`, no `CallByName`/widget property fallback (§2.18).
3. **`res:` → ResourceDictionary** — merged dictionaries, `{StaticResource}`/`{DynamicResource}`, built-in resource resolver (§2.16).
4. **ListView unified** — merge `ListView`/`UnboundListView` on `Selector`; rewrite `ListViewBase` for InvoiceGrid-style variable rows (§2.14).
5. **Memory/perf hygiene** — shared DP registry, binding detach, lightweight `CollectionChangedEventArgs`, fix `ListCollectionView` static bug (§2.19–§2.21).
6. **Documentation** — `VCF_XAML_WPF_SUBSET.md`, `VCF_INFRASTRUCTURE.md`, `VCF_LISTVIEW_ARCHITECTURE.md`, property registry, benchmarks (rewrite spec §15).

### Priority table (P0 first)

| P | Work item | Phase | Spec ref |
|---|-----------|-------|----------|
| P0 | `DependencyPropertyRegistry` + `FrameworkElement` | 1 | Rewrite §2.1, §6 |
| P0 | Strict `XAMLReader` + `TypeRegistry` + `XamlLoadException` | 0 | Rewrite §2.4, §9.1 |
| P0 | Remove orphans (`_Image`, `_TextBlock`, duplicate markup/API files) | 0 | Rewrite §18 |
| P0 | `ResourceDictionary` + `DynamicResource` + `res:` resolver | 3 | Alignment §2.16–§2.17 |
| P0 | `BindingExpression` + DataContext rebind + detach | 4 | Rewrite §7 |
| P0 | Collection fixes (lightweight args, `ListCollectionView`, Move) | 4 | Rewrite §8, bugs B1–B2 |
| P0 | ListView merge + `ListViewBase` rewrite | 4–5 | Alignment §2.14 |
| P0 | Measure/Arrange; drop public `Design*` | 1–2 | Alignment §2.7 |
| P1 | `Grid`, `StackPanel`, `ContentControl`, Border decorator | 2 | Rewrite §6 |
| P1 | Button `Content` DP; theme/style cleanup | 2b–6 | Alignment §2.10–§2.17 |
| P2 | Remaining controls, designer sync, optional utilities split | 0, 6+ | Rewrite §10 |

### Open questions (§8) — resolve or defer at kickoff

- Grid vs UniformGrid long-term default for POS migration
- `ViewBase` / strict `x:Class` failure mode
- `BindingExpression` vs listener-pull interim period
- Utilities DLL split (`Mail`, `INIParser`, AsyncKit)
- `ObservableCollection.BeginUpdate`/`EndUpdate` in Phase 4 vs later

### POS reference artifacts

- `pos-v1/UI/Resources/XAML/` — production XAML
- `pos-v1/UI/Source/Classes/MyApp.cls`, `ObjectConstructor.cls`, `*View*.cls`
- `pos-v1/UI/Resources/XAML/OrderItemsView.xml` — grid stub
- `pos-v1/UI/Resources/XAML/MenuItemsGridButton.xml` — binding density hotspot (framework-first; POS VM fix deferred)
- `pos-v1/UI/Source/Forms/FormMain.frm`, `InvoiceGridHelper.cls` — Codejock anchor (out of VCF scope)
- `docs/architecture/UI_AND_PARTITIONING_BASELINE.md`
- `docs/architecture/VCF_FRAMEWORK_REWRITE_SPEC.md`

---

## 11. Discussion log

*(Newest first.)*

### 2026-06-19 (continued)

- **Non-visual infrastructure (§2.21):** Review collections (`ObservableCollection`, `ObservableDictionary`, `List`), `CollectionViewSource`/`ListCollectionView`, change-event allocation, resources/markup types, utilities in DLL — align with WPF, light+fast fixes.

### 2026-06-19 (continued)

- **North star (§2.20):** Framework **as light and fast as possible** without **feature sacrifice** — no lite mode; wins from registry, BindingExpression, Measure/Arrange, coalesced paint; `.Tests` benches; observed POS **&lt;100 MB** normal ops.

### 2026-06-19 (continued)

- **Styles & themes (§2.17):** VCF style “classes” = WPF keyed **`Style`** + **`StaticResource`** (no HTML `Class=`). **`{ThemeResource}` → `{DynamicResource}`**; theme palettes as merged **`ResourceDictionary`**; keep **`ButtonSubmit`/`ButtonCancel`** as named styles over theme brush keys.

### 2026-06-19 (continued)

- **`res:` decisions locked (§2.16.6):** Explicit **`x:Key`** required (filename-only fallback one release). Command fragments → **`DataTemplate`**; grid cells → **`ResourceReference`** clone (not ContentControl × N); long-term → **ItemsControl** + UniformGrid panel. ContentControl only for single navigation/swap regions.

### 2026-06-19 (continued)

- **`res:` → WPF resources (§2.16):** Three mechanisms today (inline `<res:…/>`, `SetPanelContent` command params, flat `XAMLResources` dict). Target: `MergedDictionaries` + `Source=`, `{StaticResource}`, VCF `XamlResourceResolver`, `ContentControl` for panel swap; transitional `res:` shim in framework; delete POS `TryCreateObject`.

### 2026-06-19 (continued)

- **Visibility / StackPanel (§2.15):** Partial Visibility — enum exists but Hidden=Collapsed at widget level; split `Visible` bool vs `Visibility`; Border/TextBlock/Image lack support; Collapsed does not affect UniformGrid layout. StackPanel missing — P1 with Measure/Arrange.

### 2026-06-19 (continued)

- **InvoiceGrid clarified (§6.1):** Not editable DataGrid — **read-only multi-column ListView** with **parent/child rows** and **variable row height** (Codejock MeasureRow: 40/20px). VCF target: ListViewBase columns + MeasureRow + hierarchy; SelectedItem for focus.

### 2026-06-19 (continued)

- **ListView decisions (§2.14.10):** (1) **Merge** ListView + UnboundListView → one ListView on Selector. (2) **ListViewBase = vbWidgets engine** in VCF (same as TextBoxBase); wrappers thin; new engine impl OK. (3) InvoiceGrid = evolve that engine, not separate VirtualizingListView.

### 2026-06-19 (continued)

- **ListView stack review (§2.14):** Three-layer model (ListView/UnboundListView → ListViewBase → cWidgetBase). Not vbWidgets DLL List — owner-draw Cairo engine (cwWidget/Colin Edwards lineage). ItemsSource ObservableCollection-only; DataTemplate clone+DrawOn; bugs (ListCollectionView static init, Move, IItemsControl stubs). TextBox/ScrollBar same widget pattern. Refactor plan + vbWidgets vendoring option documented.

### 2026-06-19 (continued)

- **Selector support (§2.13):** User requested WPF-aligned selection in XAML. Target: `Selector` base on `ItemsControl`; DPs `SelectedItem`, `SelectedIndex`, `SelectedValue`, `SelectedValuePath`; refactor ListView; ListBox/ComboBox/TabControl Phase 5. Replaces `ListIndex` + DialogWindow `@Selected` hacks.

### 2026-06-19 (continued)

- **Templates → WPF alignment (§2.12.8):** User confirmed templates should match WPF where possible. Agreed: `DataTemplate` + `ItemsControl` + `ItemsPanelTemplate` (Phase 4); `ControlTemplate`/`ContentTemplate` later; **drop `@`-fragment pattern** in POS migration. MessageBox target XAML documented.

### 2026-06-19 (continued)

- **Template evaluation (§2.12):** MessageBox = view shell XAML + POS `@`-fragment templates (`MessageBoxButton.xml`), **not** VCF `DataTemplate`. VCF `DataTemplate` is ListView-only (clone + `{Binding}` + owner-draw). Recommend ItemsControl + DataTemplate to replace `Replace$` loops.

### 2026-06-19 (continued)

- **Compositional architecture (§2.11):** WPF builds complex controls from primitives + panels + ContentControl + templates. VCF partially composes (XAML trees, ListView DataTemplate) but Button/Border are monolithic Cairo. Target: FrameworkElement base, Border-as-decorator, ContentControl, ControlTemplate — phased with layout program.

### 2026-06-19 (continued)

- **Button Content / caption (§2.10):** VCF Button draws chrome only; text requires nested TextBlock child. Recommend WPF-like **`Content` DP** + draw string in `W_Paint` using existing font style setters; XAML `Content="..."` / `{Binding}`.

### 2026-06-19 (continued)

- **TextBox TwoWay / barcode (§2.9):** Per-char TwoWay sync → VM `OnPropertyChanged` + optional target write-back (`Text` Let resets caret). Debounced **UpdateSource** with **flush on Enter** should fix scanner jank; implement on `Binding`, not delayed display.

### 2026-06-19 (continued)

- **DP inheritance perf (§2.8):** Push model (`PassPropertyValue`, `InheritPropertyValues`) vs WPF lazy pull. `DataContext` change fans out O(n) `SetCurrentValue` + binding callbacks on nested POS trees. Fix: lazy inheritance + `BindingExpression`; batch/remove Button visual inheritance if needed.

### 2026-06-19 (continued)

- **Render/layout perf (§2.7):** Confirmed slowness with nested UniformGrids is architectural — Design* `MoveChild` cascade + `Widgets.RemoveAll` on resize + heavy Border/Button/TextBlock paint + uncoalesced `W.Refresh`. Dropping Design* + Measure/Arrange is the main fix; render/coalesce optimizations still needed.

### 2026-06-19 (continued)

- **Constructor / custom constructor (§2.6):** VCF merges internal factory + global `IObjectConstructor`. XAML uses `CreateInstance(prefix, name)`; `x:Class` two-phase init; POS `ObjectConstructor` = giant `Select Case` + app-specific `res.` loading from `XAMLResources`. Improvements: fail-loud load, built-in `res:` in framework, `TypeRegistry` / `IXamlTypeResolver`, split `Constructor`, optional `ViewBase`, remove dangerous `x:Class` fallback.

### 2026-06-19 (continued)

- **Breaking-change policy (§9):** VCF used only by POS + few internal apps. **Breaking changes acceptable** if VCF team provides **`BREAKING_CHANGES.md` + `MIGRATION.md`**. denovo will use **Cursor AI** to bulk-migrate XAML and POS source — enables **one decisive redesign** (e.g. drop `Design*`, unified DPs) instead of long dual APIs.

### 2026-06-19 (continued)

- **Dependency properties evaluation (§2.5):** Compared VCF to WPF. VCF uses per-instance DPs with `SetValue`/`SetCurrentValue` + simplified precedence; binding via listeners/callbacks. **Layout (`Design*`) is outside the DP system.** Priority improvements: `BindingExpression`, DataContext rebind, register Width/Margin/Alignment as DPs, static Register pattern, `ClearValue`, metadata callbacks, full precedence for triggers later.

### 2026-06-19

- Created this living document to collect notes until final VCF team suggestion is ready.
- **Prior thread summary captured:** CRLF/VB6 automation in denovo; VCF not in repo; full VCF source scan; XAML vs WPF dialect analysis; Design* rename + cache at load; layout evolution toward Measure/Arrange; WPF alignment phased roadmap; user goal **“as close to WPF as possible.”**
- **Pending:** User to supply rough bullets (worries, must-haves) for §8 and §10.

---

## 12. Changelog (document)

| Date | Change |
|------|--------|
| 2026-06-19 | Complete VCF team handoff package: [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) + 6 companion docs |
| 2026-06-19 | Added [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md); populated §10 final handoff draft |
| 2026-06-19 | §2.15 Visibility/Collapsed audit + StackPanel requirement |
| 2026-06-19 | §5.1–5.3 POS control gap matrix (complete VCF conversion) |
| 2026-06-19 | §6.1 InvoiceGrid = read-only hierarchical multi-column list (not DataGrid) |
| 2026-06-19 | §2.14.10 ListView merge + ListViewBase=engine decisions |
| 2026-06-19 | Added §2.14 ListView/ItemsSource/DataTemplate/vbWidgets architecture review |
| 2026-06-19 | Added §2.13 Selector (WPF selection DPs in XAML) |
| 2026-06-19 | §2.12.8 WPF-aligned template target; `@`-fragments resolved as legacy |
| 2026-06-19 | Added §2.12 template mechanisms (DataTemplate vs POS @-fragments) |
| 2026-06-19 | Added §2.11 compositional architecture (WPF-style visual tree) |
| 2026-06-19 | Added §2.10 Button Content / caption |
| 2026-06-19 | Added §2.9 TextBox TwoWay / barcode scanner binding |
| 2026-06-19 | Added §2.8 DP inheritance performance review |
| 2026-06-19 | Added §2.7 render/layout performance (nested grids) |
| 2026-06-19 | Added §2.6 constructor / custom constructor evaluation |
| 2026-06-19 | Breaking-change policy + AI migration plan; §3.3/§9 updated |
| 2026-06-19 | Added §2.5 dependency properties evaluation vs WPF |
| 2026-06-19 | Initial draft from POS/VCF coordination discussions |
