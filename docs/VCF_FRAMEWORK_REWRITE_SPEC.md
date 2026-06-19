# VCF Framework — comprehensive rewrite specification

**Status:** Draft for VCF team handoff  
**Audience:** Demac VCF maintainers · POS / DeNovo (requirements & migration)  
**VCF source:** `Demac.VCF` (separate repo, `Demac.VCF.dll`)  
**Master guide:** **[VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md)** — start here  
**Companion docs:** [VCF_WPF_ALIGNMENT_NOTES.md](./VCF_WPF_ALIGNMENT_NOTES.md) · [VCF_CLASS_REFERENCE.md](./VCF_CLASS_REFERENCE.md) · [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md) · [VCF_XAML_WPF_SUBSET.md](./VCF_XAML_WPF_SUBSET.md) · [VCF_INFRASTRUCTURE.md](./VCF_INFRASTRUCTURE.md) · [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md) · [UI_AND_PARTITIONING_BASELINE.md](./UI_AND_PARTITIONING_BASELINE.md)  
**Last updated:** 2026-06-19  

---

## Purpose

This document is the **complete rewrite blueprint** for **Demac.VCF**: every registered type, module, and public surface — mapped from **current implementation** to **target design** using **WPF semantics**, **light + fast internals** (§2.20 alignment notes), and **business-grade** practices (fail-loud, testable, documented breaking migrations).

It supplements [VCF_WPF_ALIGNMENT_NOTES.md](./VCF_WPF_ALIGNMENT_NOTES.md) ( *why* and *phasing* ) with *what to build per type*.

**Scope:** ~98 registered classes, 9 modules, 19 interfaces, ~15,500 LOC in `Classes/`. Designer and `.Tests` referenced but not line-by-line rewritten here.

**Out of scope for VCF rewrite:** POS ViewModels, `ObjectConstructor`, business data — covered by denovo `MIGRATION.md` after VCF releases.

---

## 1. North star & engineering principles

| Principle | Requirement |
|-----------|-------------|
| **WPF semantic parity** | LOB UI: DPs, bindings, styles, resources, templates, Selector, layout — see alignment §2 |
| **Light + fast** | Shared DP registry, `BindingExpression`, Measure/Arrange, lightweight collection notifications — §2.19–§2.21 |
| **No feature subset** | No “lite API”; optimize implementation, not capability — §2.20 |
| **Fail loud** | `XamlLoadException` with element/attribute context; no silent `Nothing` — §2.6 |
| **One responsibility** | Split `Constructor` god-object; split utilities from UI core — §2.21.7 |
| **Testable** | Golden XAML, binding/collection/layout benches, ListView regression — §2.20.4 |
| **Breaking OK** | Each major release: `BREAKING_CHANGES.md`, `MIGRATION.md`, semver — alignment §9 |

### 1.1 Target layer architecture

```text
┌─────────────────────────────────────────────────────────────┐
│  Application · Window · Scene · ResourceDictionary          │
│  XamlServices (read/write) · TypeRegistry · ThemesManager    │
├─────────────────────────────────────────────────────────────┤
│  FrameworkElement (DP store, layout, visibility, resources)   │
│  Control · ContentControl · Selector · ItemsControl           │
├─────────────────────────────────────────────────────────────┤
│  Panels: Grid · StackPanel · DockPanel · Border · Canvas    │
│  UniformGrid (legacy compat / migration)                     │
├─────────────────────────────────────────────────────────────┤
│  Primitives: TextBlock · Image · Shape (future)             │
├─────────────────────────────────────────────────────────────┤
│  BindingExpression · CollectionView · Markup extensions      │
│  ObservableCollection · ResourceDictionary · Style/Setter    │
├─────────────────────────────────────────────────────────────┤
│  DependencyPropertyRegistry · Value precedence · Invalidation│
├─────────────────────────────────────────────────────────────┤
│  Cairo adapter (internal): cWidgetBase sync from DP callbacks│
└─────────────────────────────────────────────────────────────┘
     Optional satellite: Demac.VCF.Core (Mail, INIParser, …)
```

### 1.2 VB6 implementation strategy

VB6 has **no implementation inheritance** for controls. Rewrite uses:

1. **`FrameworkElement` composed helper** + **codegen** for `Implements` boilerplate on thin wrappers.
2. **`DependencyPropertyRegistry`** — static metadata per type; instance **value store only**.
3. **Single paint/layout invalidation pipeline** on `FrameworkElement`.
4. **Delete** orphans: `_Image.cls`, `_TextBlock.cls`, duplicate `MarkupExtensions.cls`, stub `IDependencyPropertyCallbackListener.cls`, orphan `API.bas`.

---

## 2. Cross-cutting rewrite (applies to all types)

### 2.1 Dependency properties

| Current | Target |
|---------|--------|
| Per-instance `DependencyProperties.Register` in each control | **`DependencyPropertyRegistry.RegisterType("Button", defs…)`** once |
| Each `DependencyProperty` owns `PropertyChangedEvent` | **One change hub** per element or shared metadata callbacks |
| `GetValue` listener-pull for bindings | **`BindingExpression`** push updates — §2.5 |
| Layout as fields (`Design*`) | **DPs:** `Width`, `Height`, `Margin`, alignments — §3.3 |
| `StyleManager` + `CallByName` fallback | **DP-only** setters — §2.18 |
| No public `ClearValue` | **`ClearValue(property)`** |

### 2.2 Bindings

| Current | Target |
|---------|--------|
| `Binding` + `NestedProperty` + 3× `WithEvents` | **`BindingExpression`** per target DP |
| Source = `DataContext` **DP object** by default | Source = **resolved object**; rebind on context change |
| Stored in `control.Bindings` List forever | **`BindingExpressionCollection`** with **`Detach()`** |
| `BindingsManager.CreateBindingFromMarkup` only | Same entry + used by **`Binding` markup extension** |

### 2.3 Collections & views

| Current | Target |
|---------|--------|
| `NewList(item)` per collection change | **Lightweight `NotifyCollectionChangedEventArgs`** |
| `ListCollectionView.Initialize` static flag | **Per-instance** initialization |
| `CollectionChangedActionMove` stub | **Implement Move** |
| `ObservableDictionary` only | **`ResourceDictionary`** + merge + `Source=` — §2.16 |
| ItemsSource = `ObservableCollection` only | **`IEnumerable` adapter** + snapshot option — §2.14 |

### 2.4 XAML & creation

| Current | Target |
|---------|--------|
| `CreateInstance` + `On Error Resume Next` | **`IXamlTypeResolver`** + **`XamlLoadException`** |
| `res:` in POS `ObjectConstructor` | Built-in **`XamlResourceResolver`** — §2.16 |
| `SetProperty` → widget fallback | **Removed** |
| `x:Class` silent fallback to root tag | **Error if custom create fails** |
| Flat `XAMLResources` string dict | **Merged ResourceDictionaries** |

### 2.5 Layout & render

| Current | Target |
|---------|--------|
| `MoveChild` Design* scale cascade | **Measure → Arrange** |
| `Widgets.RemoveAll` on resize | **Stable widget tree** |
| Per-control Cairo in Button/Border | **Composition** + optional **ControlTemplate** — §2.11 |
| `OverlayWidget` per Button | **Remove**; effects via template or single overlay DP |

---

## 3. Master catalog — action summary

**Legend:** **Keep** = minimal change · **Refactor** = same role, new internals · **Merge** = combine types · **Replace** = new type supersedes · **Remove** = delete · **Split** = extract to another assembly · **Evolve** = significant API extension

| Type | Lines | Action | Phase | Priority |
|------|------:|--------|-------|----------|
| **FrameworkElement** *(new)* | — | **Replace** boilerplate base | 1 | P0 |
| **DependencyPropertyRegistry** *(new)* | — | **New** | 0–1 | P0 |
| **BindingExpression** *(new)* | — | **Replace** `Binding` graph | 4 | P0 |
| **XamlServices** *(new)* | — | **Split** from `Constructor` | 0 | P0 |
| **TypeRegistry** *(new)* | — | **Replace** ad hoc `CreateInstance` | 0 | P0 |
| **ResourceDictionary** *(new)* | — | **Evolve** from `ObservableDictionary` | 3 | P0 |
| **Selector** *(new base)* | — | **New** | 4–5 | P0 |
| **ItemsControl** *(new)* | — | **New** | 4 | P0 |
| **ContentControl** *(new)* | — | **New** | 2b | P1 |
| **Grid / StackPanel** *(new)* | — | **New** | 2 | P0–P1 |
| Button | 712 | **Refactor** → `FrameworkElement` | 1–2b | P0 |
| TextBlock | 630 | **Refactor** | 1 | P1 |
| TextBox | 486 | **Refactor** thin over TextBoxBase | 1 | P1 |
| TextBoxBase | 1573 | **Refactor** engine; keep widget core | 1–5 | P1 |
| Image | 328 | **Refactor** | 2 | P2 |
| Border | 414 | **Refactor** → decorator | 2 | P0 |
| Panel | 376 | **Refactor** | 2 | P1 |
| UniformGrid | 515 | **Refactor**; legacy compat | 1–2 | P0 |
| UserControl | 391 | **Refactor** | 1 | P1 |
| Window | 552 | **Refactor** | 1 | P1 |
| Scene | 366 | **Refactor** / merge with visual root | 1 | P2 |
| ListView | 755 | **Merge** + **Refactor** → `Selector` | 4–5 | P0 |
| UnboundListView | 448 | **Remove** (merged into ListView) | 5 | P0 |
| ListViewBase | 1178 | **Refactor** / rewrite engine | 4–5 | P0 |
| ScrollBar | 265 | **Keep** / minor refactor | 5 | P2 |
| OverlayWidget | 55 | **Remove** | 1 | P1 |
| WindowsFormsHost | 391 | **Keep** | — | P3 |
| ListViewPropertyChangedHandler | 36 | **Remove** → part of BindingExpression | 4 | P1 |
| DependencyObjectBase | 42 | **Merge** into FrameworkElement | 1 | P0 |
| DependencyProperties | 101 | **Refactor** → instance store | 1 | P0 |
| DependencyProperty | 234 | **Refactor** → metadata in registry | 1 | P0 |
| DependencyPropertyMetadata | 23 | **Keep**; add callbacks | 1 | P0 |
| DependencyPropertiesStatic | 65 | **Refactor** lazy inheritance | 1–4 | P0 |
| Binding | 353 | **Replace** → BindingExpression | 4 | P0 |
| BindingsManager | 156 | **Refactor** | 4 | P0 |
| NestedProperty | 342 | **Remove** (inside BindingExpression) | 4 | P0 |
| SelfBinding | 27 | **Keep** | 4 | P2 |
| PropertyChangedEvent | 59 | **Refactor** hub pattern | 1–4 | P1 |
| StaticResourceExtension | 180 | **Keep** / refactor lookup | 3 | P0 |
| ThemeResource | 63 | **Deprecate** → DynamicResource | 3 | P1 |
| DynamicResourceExtension *(new)* | — | **New** | 3 | P0 |
| XAMLDependencyPropertyManager | 72 | **Merge** into XamlServices | 0 | P1 |
| XAMLImagePropertyManager | 44 | **Keep** | 3 | P2 |
| XAMLThicknessConstructor | 36 | **Keep** | 1 | P2 |
| List | 244 | **Keep** internal; limit public use | 4 | P1 |
| ObservableCollection | 381 | **Refactor** notifications + DeferRefresh | 4 | P0 |
| ObservableDictionary | 286 | **Evolve** → ResourceDictionary | 3 | P0 |
| UIElementCollection | 148 | **Keep** | 1 | P1 |
| CollectionChangedEventArgs | 65 | **Refactor** lightweight | 4 | P0 |
| CollectionChangedEvent | 60 | **Keep** | 4 | P1 |
| CollectionViewSource | 45 | **Keep** | 4 | P0 |
| ListCollectionView | 207 | **Fix** + extend ICollectionView | 4 | P0 |
| XAMLReader | 581 | **Refactor** strict resolver | 0–3 | P0 |
| XAMLStyleReader | 87 | **Keep** | 3 | P1 |
| MarkupExtensions *(MakupExtensions.cls)* | 217 | **Refactor**; fix filename typo | 3 | P1 |
| Style | 254 | **Keep**; ThemesManager event | 3–6 | P1 |
| Setter | 26 | **Keep** | 3 | P2 |
| StyleManager | 153 | **Refactor** DP-only | 3–6 | P0 |
| ThemesManager | 199 | **Refactor** merge/swap themes | 3 | P0 |
| DataTemplate | 28 | **Evolve** full tree inflation | 4 | P0 |
| Application | 132 | **Refactor** | 3 | P0 |
| ApplicationStatic | 52 | **Keep** | 3 | P2 |
| Constructor | 164 | **Split** → XamlServices + thin VCF | 0 | P0 |
| NamingManager | 63 | **Keep** | 3 | P2 |
| Interaction | 18 | **Keep** | — | P3 |
| UIElementBase | 74 | **Merge** into FrameworkElement | 1 | P0 |
| Thickness | 19 | **Keep** immutable | 1 | P2 |
| SolidColorBrush | 25 | **Keep** | 3 | P2 |
| Color | 145 | **Keep** | — | P2 |
| Variable | 93 | **Keep** XAML boxed values | 3 | P2 |
| ArrayWrapper | 158 | **Keep** | 3 | P3 |
| API | 90 | **Keep** | — | — |
| Conversion | 118 | **Keep** | — | — |
| ObjectStatic | 194 | **Keep** | — | — |
| Information | 28 | **Keep** | — | P3 |
| StringConversion | 87 | **Split** optional | 0 | P3 |
| StringProcessor | 65 | **Split** optional | 0 | P3 |
| INIParser | 48 | **Split** → VCF.Core | 0 | P3 |
| Mail | 80 | **Split** → VCF.Core | 0 | P3 |
| Environment | 20 | **Split** optional | 0 | P3 |
| BackgroundWorker | 235 | **Split** → VCF.Core or AsyncKit only | 0 | P3 |
| InternalWorker | 76 | **Split** | 0 | P3 |
| ErrorInfo | 40 | **Split** | 0 | P3 |
| Function | 74 | **Keep** (delegate) | — | P3 |
| StaticClasses | 58 | **Refactor** facade | 0 | P2 |
| _Image / _TextBlock | 292/529 | **Remove** | 0 | P0 |
| MarkupExtensions.cls orphan | 193 | **Remove** | 0 | P0 |

---

## 4. Interfaces — contract target

| Interface | Current | Target changes |
|-----------|---------|----------------|
| **IDependencyObject** | `DependencyProperties`, `Parent`, `Children` | Add **`ClearValue`**, **`GetValue`/`SetValue`** surface if exposed publicly |
| **IControl** | `Widget`, `Widgets`, `Children`, enums | **`FrameworkElement`** implements; **`Widget` Friend/internal** only |
| **IUIElement** | Design* layout | **WPF layout DPs** on `IUIElement` / `FrameworkElement` |
| **IUserControl** | Large duplicate of layout | **Thin**; `InitializeComponent` via **ViewBase** optional |
| **IWindow** | `Base`, `InitializeComponent` | **Keep** |
| **IApplication** | Resources, Run, FindResource | **Keep**; Resources = **ResourceDictionary** |
| **IItemsControl** | `ItemsSource`, `ItemTemplate` | Add **`ItemsPanel`**, **`ItemContainerStyle`** (phased) |
| **ICommand** | `Execute`, `CanExecute` | Add **`CanExecuteChanged`** event (WPF) |
| **IValueConverter** | `Convert`, `ConvertBack` | **Keep** |
| **IMarkupExtension** | `ProvideValue()` | **Keep**; all extensions implement |
| **IObjectConstructor** | Single `CreateInstance(String)` | **Extend** or replace with **TypeRegistry** |
| **INotifyPropertyChanged** | `PropertyChangedEvent()` | **Keep** |
| **INotifyCollectionChanged** | `CollectionChangedEvent()` | **Keep** |
| **IVisualChild** | `DrawOn` | **Keep** for ListView owner-draw; not on TextBlock/Image public path |
| **ICloneable** | `Clone()` | **Keep** on TextBlock if still needed for templates |
| **IBackgroundTask** | `Execute` | Move to **VCF.Core** |
| **IDependencyPropertyCallbackListener** | Full in `IDependencyPropertyCallback.cls` | **Remove** after BindingExpression; delete stub duplicate file |

---

## 5. Modules — complete API & rewrite

### 5.1 `modConstructors.bas`

| Member | Role | Target |
|--------|------|--------|
| `CustomConstructor` | Global app hook | **`IApplication.RegisterTypes`** / **TypeRegistry** |
| `NewCollectionChangedEventArgs` | Factory | **Lightweight** factory; optional pool |
| `NewObject` | Markup type create | **`XamlServices.CreateObject`** |
| `NewCustomObject` | App types | **TypeRegistry** |
| `NewList` | List factory | **Internal** only |
| `NewUIElementCollection` | Child collection | **Keep** |
| `NewDependencyProperties` | DP bag | **`FrameworkElement.InitProperties`** |
| `NewThickness` | Parse helper | **Keep** |
| `NewBinding` | Code bindings | **`BindingExpression.Attach`** |
| `NewDependencyPropertyMetadata` | Metadata factory | **Registry** internal |
| `NewFunction` | Delegate | **Keep** |
| `CreateInstance` | XAML type resolver | **`XamlTypeResolver.Resolve(ns, name)`** |
| `NewUIElementBase` | Compose base | **Fold into FrameworkElement** |
| `NewStyle` | Style factory | **Keep** |
| `Variable` / `ArrayWrapper` | XAML literals | **Keep** |

### 5.2 `modStaticClasses.bas`

| Singleton | Target |
|-----------|--------|
| `CollectionViewSource` | **Keep** |
| `DependencyPropertiesStatic` | **Refactor** lazy inherit |
| `BindingsManager` | **Merge** into XamlServices or keep stateless |
| `Application` | **Keep** |
| `API`, `Conversion`, `Object`, `Color` | **Keep** in VCF |
| `INIParser`, `Mail`, `Environment`, `StringConversion`, `StringProcessor` | **Move** to VCF.Core (optional) |
| `NamingManager` | **Keep** |

### 5.3 Other modules

| Module | Public API | Target |
|--------|------------|--------|
| **modVisibilityHelper** | `SetVisibility(W, Value)` | **Merge** into `FrameworkElement.Visibility` DP handler |
| **modWindowsFormsHostHelper** | `SetChild`, `ShowWindow` | **Keep** |
| **modInformation** | `IsNothing` | **Duplicate** of `Information` — consolidate |
| **modUDFConstructors** | `NewCornerRadius` | **Keep** |
| **modInternalSignals** | APC slots for BackgroundWorker | **Move** with AsyncKit |
| **modAPI** | `DllPath()` | **Keep** |
| **modStyleWriter** | (commented out) | **Remove** or revive for designer only |
| **API.bas** orphan | Duplicates `modAPI` | **Delete** |

---

## 6. Detailed rewrite — visual controls

### 6.1 `Button` (712 lines)

**Public API today:** `Name`, `Bindings`, `Style`, `Selected`, `Margin`, `Command`, `CommandParameter`, `ClickMode`, `BorderWidth`, `GradientBackground`, `CornerRadius`, `Design*`, `Move`, `Children`, `Widget`.

**Issues:** Dual storage (DP + fields + widget); `OverlayWidget`; Design* layout; monolithic Cairo paint; ~200 lines interface delegation.

**Target:**

- Inherits **`FrameworkElement`** (composed).
- **DPs:** `Content` (string Phase 5 / object later), `Command`, `CommandParameter`, `IsPressed` internal, layout/visual DPs from registry.
- **Template:** Phase 6 **ControlTemplate** default; until then simplified single-pass paint.
- **Remove:** `OverlayWidget`, public `GradientBackground` → template/trigger.
- **Widget sync:** `OnBackColorChanged` etc. in registry callbacks only.

### 6.2 `TextBlock` (630 lines)

**Public:** `Text`, font DPs, alignments, `ScaleFont`, `Clone`, `DrawOn` via widget path.

**Target:** `FrameworkElement`; **`FormattedText`** cache optional P2; **`DrawOn`** for ListView templates via **IVisualChild** internal path only.

### 6.3 `TextBox` / `TextBoxBase` (486 / 1573)

**Target:** `TextBox` thin wrapper; engine stays **vbWidgets/cWidget** lineage; **TwoWay** binding + **`UpdateSourceDelay`** §2.9; DPs for `Text`, `SelectionStart`, etc. phased.

### 6.4 `Border` (414)

**Target:** **Decorator** — single **`Child`** DP; border/background/corner via DPs; **Measure/Arrange** passes to child.

### 6.5 `Panel`, `UniformGrid`, `UserControl`, `Window`, `Scene`

**UniformGrid:** **Keep** for POS migration; add **Grid** as primary; cell layout via **Arrange**.

**UserControl / Window:** **ViewBase** optional `InitializeComponent`; **Fail-loud** XAML load.

### 6.6 `Image`, `ScrollBar`, `WindowsFormsHost`

**Keep** with FrameworkElement migration; **WindowsFormsHost** low churn.

### 6.7 ListView stack — see alignment §2.14

**ListViewBase:** Rewrite **MeasureRow**, columns, scroll, owner-draw; variable row height for InvoiceGrid.

**ListView:** Single public control on **`Selector`**; **`ItemsSource`**, **`ItemTemplate`**, **`ItemContainerGenerator`**-style (phased).

**UnboundListView:** **Delete**; mode flag internal if needed short-term.

---

## 7. Detailed rewrite — data & binding

### 7.1 `Binding` → `BindingExpression`

**Current public:** `Mode`, `Source`, `Target`, `TargetProperty`, `Path`, `Converter`, `StringFormat`, `ProvideValue`.

**Target class `BindingExpression`:**

```text
Attach(target, targetProperty, source, path, mode, converter, …)
Detach()
UpdateTarget()
UpdateSource()
OnDataContextChanged(newContext)
```

**Remove:** `NestedProperty`, triple `WithEvents`, `SrcDepObj` callback fan-out.

### 7.2 `BindingsManager`

**Keep** `CreateBindingFromMarkup`; delegate to **`BindingExpression.AttachFromMarkup`**.

### 7.3 `DependencyProperty` / `DependencyProperties`

**Replace** per-property `PropertyChangedEvent` with registry metadata + element hub.

**Public add:** consumers may read **`GetValue`/`SetValue`/`ClearValue`** on `FrameworkElement`.

---

## 8. Detailed rewrite — collections

### 8.1 `ObservableCollection`

**Public API:** Full list interface + `CollectionChanged` + `GetHashCode`.

**Rewrite:**

- **`BeginUpdate`/`EndUpdate`** (DeferRefresh).
- **Single-item notification** without `New List`.
- **Remove** duplicate `Exists` linear scan hot paths where possible.

### 8.2 `ObservableDictionary` → `ResourceDictionary`

**Add:** `MergedDictionaries`, `Source` URI loader, **`TryGetResource(key)`**.

### 8.3 `ListCollectionView`

**Fix:** Remove static `bIsInitialized`.

**Add:** `CurrentItem`, `MoveCurrentTo*`, **Move** action handler, optional **Sort** P2.

**Wire:** **`Selector.SelectedItem`** sync — §2.13.

### 8.4 `List`

**Role:** Internal batches, tests — **not** for INCC payloads long-term.

---

## 9. Detailed rewrite — XAML, styles, application

### 9.1 `XAMLReader` (581)

**Public:** `Load`, `LoadSuperclassData`, `LoadApp`.

**Rewrite phases:**

1. Strict **`SetProperty`** removal; **resolver** only.
2. **ResourceReference**, **merged dictionaries**.
3. **x:Class** + **ViewBase** hook.

**Internal split:** DOM walk · property setter · name registration · template inflate.

### 9.2 `MarkupExtensions` (fix `MakupExtensions.cls` typo)

**Extensions:** `StaticResource`, `DynamicResource`, `Binding`, `ThemeResource` (shim), `Self`.

### 9.3 `Style` / `StyleManager` / `ThemesManager`

**Style:** Keep **`BasedOn`**, **`SetSetter`**, theme change re-apply.

**ThemesManager:** **Active theme = merged dictionary swap**; fix **`ThemeCkanged`** typo → **`ThemeChanged`**.

### 9.4 `Application` / `ApplicationStatic`

**Resources:** root **ResourceDictionary**; **StartupUri** via resolver; **Shutdown** mode optional.

---

## 10. Utilities & async (split recommendation)

| Type | Recommendation |
|------|----------------|
| Mail | **Demac.VCF.Core** or app — CDO dependency |
| INIParser | **Core** |
| BackgroundWorker + InternalWorker + modInternalSignals | **Demac.VCF.Async** package |
| StringProcessor, StringConversion | **Core** or **Demac.Common** |
| Environment.TickCount | **Keep** or use Win32 inline |

---

## 11. Designer (parallel track)

| Component | Action |
|-----------|--------|
| **XAMLWriter** | Update for WPF property names; no `Design*` |
| **PropertyEditor** | Bind to **DP registry** metadata |
| **DesignSurface** | Use new **Measure/Arrange** |
| **.NET Designer** | Sync with VB6 or deprecate one |

---

## 12. Known bugs — must fix in rewrite

| ID | Location | Issue |
|----|----------|-------|
| B1 | `ListCollectionView.Initialize` | Static flag — only first view works |
| B2 | ListView | `CollectionChangedActionMove` not handled |
| B3 | ListView | `IItemsControl_ItemTemplate` stubs |
| B4 | ListView | ItemsSource non-ObservableCollection silent fail |
| B5 | XAMLReader | Silent failures / widget fallback |
| B6 | DataContext change | Bindings not recreated — TODO in POS views |
| B7 | `ThemesManager` | Typo event name |
| B8 | Style | `BasedOn` + missing key edge cases |
| B9 | Visibility | Hidden=Collapsed; Collapsed no layout — §2.15 |
| B10 | Duplicate `modAPI` / interface files | Load ambiguity risk |

---

## 13. Complete public API index (registered classes)

*(Public + exposed Friend factories; Private helpers omitted.)*

### 13.1 Controls & layout

**Button:** `Bindings`, `Selected`, `Style`, `Margin`, `Command`, `CommandParameter`, `ClickMode`, `BorderWidth`, `GradientBackground`, `CornerRadius`, `DesignLeft/Top/Width/Height`, `Move`, `Children`, `Widget`, `Widgets`, `Parent`, `DependencyProperties`, `Name`.

**TextBlock:** `Text`, `ForeColor`, font properties, `HorizontalAlignment`, `VerticalAlignment`, `ScaleFont`, `Design*`, `Style`, `Bindings`, `Clone`.

**TextBox:** `Text`, `Alignment`, `VCenter`, `PasswordChar`, `CueBanner`, `Visibility`, `Border`, `Focused`, `Design*`, `Style`, `Bindings`.

**TextBoxBase:** Full text widget API (see inventory §1) — **document subset in VCF_XAML_WPF_SUBSET.md**.

**Border:** `CornerRadius`, `Design*`, `Style`, `Children`, standard layout.

**Panel / UniformGrid / UserControl / Window / Scene:** Layout DPs, `Children`, `Visibility` (where exposed), `Style`, `Bindings`.

**Image:** `ImageKey`, `KeepAspectRatio`, `DrawOn`.

**ScrollBar:** `Value`, `Min`, `Max`, `Vertical`, `SmallChange`, `LargeChange`.

**ListView:** `ItemsSource`, `ItemTemplate`, `Base` (ListViewBase), `Resources`, row/col colors, `Name`.

**UnboundListView:** Same minus ItemsSource; `Refresh`.

**WindowsFormsHost:** `AutomaticallyUnloadContent`, `Children`.

### 13.2 Data & binding

**Binding:** `Mode`, `Source`, `Target`, `TargetProperty`, `Path`, `Converter`, `StringFormat`, `ProvideValue`.

**DependencyProperty:** `Name`, `PropertyType`, `Metadata`, `GetValue`, `SetValue`, `SetCurrentValue`, `PropertyChangedEvent`.

**DependencyProperties:** `Register`, `GetValue`, `SetValue`, `SetCurrentValue`, `Exists`, `GetProperty`.

**NestedProperty:** `Path`, `Source`, `GetValue`, `SetValue`, `Initialize` — **REMOVE**.

**ObservableCollection / List / ObservableDictionary:** Standard collection ops — see §8.

**CollectionViewSource:** `GetDefaultView`.

**ListCollectionView:** Collection ops + `CurrentItem`, `CurrentPosition`, `MoveCurrentTo*`.

### 13.3 XAML & app

**XAMLReader:** `Load`, `LoadSuperclassData`, `LoadApp`.

**Application:** `Run`, `Resources`, `StartupURI`, `FindResource`, `TryFindResource`, `Windows`.

**Constructor / VCF module:** `NewWindow`, `NewUserControl`, `NewBinding`, `NewStyle`, `CreateInstance`, `SetCustomConstructor`.

**Style:** `Initialize`, `TargetType`, `Key`, `BasedOn`, `GetSetter`, `SetSetter`, `Clear`.

**ThemesManager:** `ActiveThemeName`, `ActiveTheme`.

**Markup extensions:** `StaticResourceExtension.ProvideValue`, `ThemeResource.ProvideValue`, `SelfBinding`.

### 13.4 Value types & helpers

**Thickness:** `Left`, `Top`, `Right`, `Bottom`.

**Color:** `FromHtml`, `ToHtml`, `FromRGB`, `Multiply`, etc.

**SolidColorBrush:** `Color`.

**Variable / ArrayWrapper:** XAML literal support.

**API / Conversion / ObjectStatic:** Utility methods — stable, minimal changes.

---

## 14. Test matrix (rewrite acceptance)

| Test | Validates |
|------|-----------|
| Golden XAML load (POS Sales subset) | XAMLReader + resolver |
| DP registry memory bench | §2.19 |
| 1000× ObservableCollection.Add | §2.21 lightweight args |
| Two ListCollectionView instances | Bug B1 |
| ListView Move + bound ItemsSource | B2, B4 |
| DataContext swap rebind | BindingExpression |
| Theme switch | DynamicResource + Styles |
| Resize nested grid bench | Measure/Arrange §2.7 |
| Navigation leak 50× | Detach |
| InvoiceGrid MeasureRow golden | ListViewBase |

---

## 15. Documentation deliverables (VCF repo)

| Document | Content |
|----------|---------|
| **VCF_XAML_WPF_SUBSET.md** | Supported markup grammar |
| **VCF_INFRASTRUCTURE.md** | Collections, views, events |
| **VCF_LISTVIEW_ARCHITECTURE.md** | Engine + Selector |
| **VCF_PROPERTY_REGISTRY.md** | All DPs per type |
| **BREAKING_CHANGES.md** | Per release |
| **MIGRATION.md** | XAML + VB6 migration |
| **VCF_PERFORMANCE_BENCHMARKS.md** | Baseline numbers + CI thresholds |

---

## 16. Phased execution order (recommended)

```text
Phase 0  Spec + TypeRegistry + XamlLoadException + delete orphans + tests scaffold
Phase 1  DependencyPropertyRegistry + FrameworkElement + Measure/Arrange + Visibility
Phase 2  Grid, StackPanel, Border decorator, ContentControl
Phase 3  ResourceDictionary, DynamicResource, XamlResourceResolver, strict reader
Phase 4  BindingExpression, collection fixes, ItemsControl, Selector/ListView merge start
Phase 5  ListViewBase rewrite, remaining controls, Button Content
Phase 6  ControlTemplate, Style.Triggers, render optimizations
Phase 7  POS migration support (denovo AI-assisted)
```

---

## 17. denovo coordination

| denovo artifact | When |
|-----------------|------|
| Pin VCF version in `DeNovo.vbp` | Each VCF tag |
| Integration smoke | login → sales → grid → dialog |
| AI XAML migration scripts | After Phase 3+ docs |
| This spec + alignment §9 | VCF team kickoff |

---

## 18. Appendix — file inventory & orphans

**Registered:** 98 classes in `Demac.VCF.vbp`.

**Delete candidates:**

- `Classes/_Image.cls` (`Image_OLD`)
- `Classes/_TextBlock.cls` (`TextBlock_BAK`)
- `Classes/MarkupExtensions.cls` (duplicate of `MakupExtensions.cls`)
- `Classes/IDependencyPropertyCallbackListener.cls` (stub; use `IDependencyPropertyCallback.cls`)
- `Modules/API.bas` (duplicate `modAPI`)

**Duplicate trees:**

- `.DevComponents/AsyncKit/` mirrors BackgroundWorker stack — **single source**.

**Styles:** `Styles/*.xml` — regenerate from merged app theme or keep as merge inputs.

---

## 19. Appendix — complete `Demac.VCF.vbp` registry (98 entries)

Every class registered in the shipping DLL. **Action** matches §3 master catalog.

| # | Type | Kind | Action | Notes |
|---|------|------|--------|-------|
| 1 | Button | Control | Refactor | §6.1 |
| 2 | CollectionChangedEventArgs | Data | Refactor | Lightweight payloads |
| 3 | IControl | Interface | Evolve | Widget internal |
| 4 | IUIElement | Interface | Evolve | WPF layout DPs |
| 5 | IVisualChild | Interface | Keep | ListView owner-draw |
| 6 | List | Collection | Keep internal | Not for INCC batches |
| 7 | ObservableCollection | Collection | Refactor | §8.1 |
| 8 | Panel | Layout | Refactor | FrameworkElement |
| 9 | UIElementCollection | Collection | Keep | Child list |
| 10 | Constructor | Factory | Split | → XamlServices |
| 11 | Window | Shell | Refactor | Fail-loud load |
| 12 | Scene | Visual | Refactor | Visual root |
| 13 | XAMLReader | XAML | Refactor | §9.1 strict |
| 14 | UniformGrid | Layout | Refactor | Legacy compat |
| 15 | ICommand | Interface | Evolve | CanExecuteChanged |
| 16 | DependencyProperty | Core | Refactor | Registry metadata |
| 17 | DependencyProperties | Core | Refactor | Instance store |
| 18 | IDependencyObject | Interface | Keep | +ClearValue |
| 19 | DependencyObjectBase | Core | Merge | → FrameworkElement |
| 20 | Border | Decorator | Refactor | Single Child DP |
| 21 | Thickness | Value | Keep | Immutable |
| 22 | IObjectConstructor | Interface | Evolve | TypeRegistry |
| 23 | XAMLImagePropertyManager | XAML | Keep | Image attrs |
| 24 | XAMLDependencyPropertyManager | XAML | Merge | XamlServices |
| 25 | XAMLThicknessConstructor | XAML | Keep | Thickness parse |
| 26 | IMarkupExtension | Interface | Keep | |
| 27 | MarkupExtensions | XAML | Refactor | Fix filename typo |
| 28 | Binding | Binding | Replace | → BindingExpression |
| 29 | INotifyPropertyChanged | Interface | Keep | |
| 30 | PropertyChangedEvent | Event | Refactor | Hub pattern |
| 31 | IValueConverter | Interface | Keep | |
| 32 | ListCollectionView | View | Fix + extend | Bug B1, Move |
| 33 | CollectionViewSource | View | Keep | |
| 34 | StaticClasses | Facade | Refactor | modStaticClasses |
| 35 | TextBoxBase | Control | Refactor | Widget engine |
| 36 | ScrollBar | Control | Keep | Minor refactor |
| 37 | ListViewBase | Control | Refactor | §6.7 engine |
| 38 | TextBox | Control | Refactor | Thin wrapper |
| 39 | ListView | Control | Merge | → Selector |
| 40 | ObservableDictionary | Resources | Evolve | ResourceDictionary |
| 41 | UserControl | Shell | Refactor | ViewBase optional |
| 42 | IUserControl | Interface | Thin | |
| 43 | DependencyPropertyMetadata | Core | Keep | + callbacks |
| 44 | DependencyPropertiesStatic | Core | Refactor | Lazy inherit |
| 45 | Function | Delegate | Keep | Command param |
| 46 | BindingsManager | Binding | Refactor | §7.2 |
| 47 | IDependencyPropertyCallbackListener | Interface | Remove | Use BindingExpression |
| 48 | Application | App | Refactor | ResourceDictionary root |
| 49 | IApplication | Interface | Keep | |
| 50 | ApplicationStatic | App | Keep | Singleton |
| 51 | IWindow | Interface | Keep | |
| 52 | API | Utility | Keep | |
| 53 | ObjectStatic | Utility | Keep | |
| 54 | UIElementBase | Core | Merge | → FrameworkElement |
| 55 | StaticResourceExtension | Markup | Keep | Faster lookup |
| 56 | WindowsFormsHost | Interop | Keep | |
| 57 | StringConversion | Utility | Split optional | VCF.Core |
| 58 | NestedProperty | Binding | Remove | Inside BindingExpression |
| 59 | ListViewPropertyChangedHandler | Binding | Remove | BindingExpression |
| 60 | DataTemplate | Template | Evolve | Full inflate |
| 61 | IItemsControl | Interface | Evolve | ItemsPanel |
| 62 | ICloneable | Interface | Keep | TextBlock clone |
| 63 | Interaction | Utility | Keep | |
| 64 | NamingManager | XAML | Keep | x:Name |
| 65 | INotifyCollectionChanged | Interface | Keep | |
| 66 | CollectionChangedEvent | Event | Keep | |
| 67 | Color | Value | Keep | |
| 68 | StringProcessor | Utility | Split optional | VCF.Core |
| 69 | Information | Utility | Keep | Consolidate modInformation |
| 70 | SelfBinding | Markup | Keep | |
| 71 | UnboundListView | Control | Remove | Merged ListView |
| 72 | StyleManager | Style | Refactor | DP-only |
| 73 | XAMLStyleReader | XAML | Keep | |
| 74 | Style | Style | Keep | ThemeChanged |
| 75 | Environment | Utility | Split optional | |
| 76 | OverlayWidget | Control | Remove | Button overlay |
| 77 | TextBlock | Primitive | Refactor | §6.2 |
| 78 | Image | Primitive | Refactor | |
| 79 | ThemeResource | Markup | Deprecate | DynamicResource |
| 80 | Setter | Style | Keep | |
| 81 | ThemesManager | Theme | Refactor | Merged dict swap |
| 82 | SolidColorBrush | Value | Keep | |
| 83 | Conversion | Utility | Keep | Internal |
| 84 | Mail | Utility | Split | VCF.Core |
| 85 | INIParser | Utility | Split | VCF.Core |
| 86 | Variable | XAML | Keep | Boxed literals |
| 87 | ArrayWrapper | XAML | Keep | |
| 88 | BackgroundWorker | Async | Split | AsyncKit |
| 89 | IBackgroundTask | Interface | Split | AsyncKit |
| 90 | InternalWorker | Async | Split | AsyncKit |
| 91 | ErrorInfo | Async | Split | AsyncKit |

**New types (not in vbp today):** `FrameworkElement`, `DependencyPropertyRegistry`, `BindingExpression`, `XamlServices`, `TypeRegistry`, `ResourceDictionary`, `Selector`, `ItemsControl`, `ContentControl`, `Grid`, `StackPanel`, `DynamicResourceExtension`, `XamlLoadException`, optional `ViewBase`.

**Unregistered orphans (delete):** `_Image.cls`, `_TextBlock.cls`, `MarkupExtensions.cls` (duplicate), `IDependencyPropertyCallbackListener.cls` (stub file), `Modules/API.bas`.

**Modules (9):** `modConstructors`, `modStaticClasses`, `modVisibilityHelper`, `modWindowsFormsHostHelper`, `modInformation`, `modUDFConstructors`, `modInternalSignals`, `modAPI`, `modStyleWriter` (commented) — see §5.

**Designer / tests (parallel):** `Designer/` (~15 classes), `.Tests/` (~10 classes) — update after Phase 1–3 core lands.

---

*This spec is living documentation. Update when VCF team resolves §8 open questions in [VCF_WPF_ALIGNMENT_NOTES.md](./VCF_WPF_ALIGNMENT_NOTES.md).*
