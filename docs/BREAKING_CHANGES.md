# Demac.VCF — breaking changes log

**Program:** WPF alignment (Phases 0–7)  
**Maintained by:** VCF team · POS migration steps in [MIGRATION.md](./MIGRATION.md)

---

## [2.0.0] — 2026-06-20 — Phase 0 (foundation)

### Breaking

- **Orphan source removed:** Unregistered duplicate files deleted from the repo (`_Image.cls`, `_TextBlock.cls`, duplicate `MarkupExtensions.cls`, stub `IDependencyPropertyCallbackListener.cls`, orphan `Modules/API.bas`). **Migration:** None — these were not in `Demac.VCF.vbp`.

### Added (non-breaking until strict mode enabled)

- **`XamlLoadException`** — structured XAML load errors (element, property, line context).
- **`TypeRegistry`** — register app types by name; used by `CreateInstance` before `CreateObject`.
- **`VCF.StrictXamlLoad`** — when `True`, malformed XML and unknown types raise `XamlLoadException` instead of returning `Nothing`. Default **`False`** for POS compatibility; enable in `.Tests/Phase0` and CI.

### Bug fixes

- **B1:** `ListCollectionView.Initialize` — static init flag replaced with per-instance initialization (second view no longer blocked).

### Deprecated (remove in Phase 1+)

- Public `DesignLeft/Top/Width/Height` → `Width`, `Height`, `Margin` DPs (Phase 1).
- `UnboundListView` → merged `ListView` (Phase 5 — **removed in 2.9.0**).
- `ThemeResource` markup → `{DynamicResource}` (Phase 3).
- `CallByName` XAML property fallback → DP-only setters (Phase 3).

---

---

---

## [2.10.0] — 2026-06-21 — Phase 5b (MeasureRow — validated)

Tag: **`v2.10.0-wpf-alignment-p5b`** · Phase0 **24/24** pass.

### Added

- **`ListViewBase.MeasureRow`** event — host sets per-row height when **`FixedRowHeight = False`**.
- **`ListViewBase.FixedRowHeight`** (default **`True`**) · **`InvalidateRowHeights`** · **`MeasuredRowHeight(Index)`**.
- **`ListView.MeasureRow`** — bubbled from engine for owner-draw / InvoiceGrid adapters.

### Changed

- Variable-height paint, hit-test, scroll, and **`EnsureVisibleSelection`** for single-column lists (**`ImageView = False`**). Fixed-height and grid **`ImageView`** paths unchanged.

### Test

- **P5b-MSR** in `.Tests/Phase0` (suite → **24** tests).

---

## [2.9.0] — 2026-06-21 — Phase 5a (ListView merge — validated)

Tag: **`v2.9.0-wpf-alignment-p5a`** · Phase0 **23/23** pass.

### Breaking

- **`UnboundListView` removed** — use **`ListView`** with no **`ItemsSource`** for owner-draw mode. XAML **`<UnboundListView/>`** still loads via alias (creates **`ListView`**).

### Changed

- **`ListView`** — forwards **`ListViewBase`** input/draw/scroll events; **`Refresh()`**; owner-draw when **`ItemsSource`** is **`Nothing`**; **`SelectedIndex`** syncs with **`ListIndex`** in owner-draw mode.

### Test

- **P5a-OWN** in `.Tests/Phase0` (suite → **23** tests).

---

## [2.8.0] — 2026-06-21 — Phase 4d (Selector — validated)

Tag: **`v2.8.0-wpf-alignment-p4d`** · Phase0 **22/22** pass.

### Added

- **`Selector`** — WPF-aligned selection base on **`ItemsControl`**: **`SelectedItem`**, **`SelectedIndex`**, **`SelectedValue`**, **`SelectedValuePath`**; syncs with **`ListCollectionView`**.
- **`ISelector`** interface · **`modSelectorEngine`** shared selection helpers.

### Changed

- **`ListView`** — implements **`ISelector`**; exposes selection DPs in XAML/code; **`IItemsControl_ItemTemplate`** wired.

### Bug fixes

- **B4 (ListView / Selector / ItemsControl):** non-**`ObservableCollection`** **`ItemsSource`** raises **`Err`** at **`DependencyProperties.SetValue`** (before CSEH-wrapped DP callback; no spurious IDE modal).
- **ListView init:** silent list index/count setters avoid benchmark side effects; **`SelectedItem`** scalar values use **`SelectedValue`** in tests.

### Test

- **P4d-SEL** in `.Tests/Phase0` (suite → **22** tests).

---

## [2.7.0] — 2026-06-20 — Phase 4c (ItemsControl — validated)

Tag: **`v2.7.0-wpf-alignment-p4c`** · Phase0 **21/21** pass.

### Added

- **`ItemsControl`** — WPF-aligned items presenter: **`ItemsSource`**, **`ItemTemplate`**, default vertical **`StackPanel`** items host; incremental **`CollectionChanged`** updates.
- **`modItemTemplateEngine`** — shared item template cloning (**`CloneItemVisualForItem`** for **`ItemsControl`**, **`CloneDataTemplateForItem`** for **`ListView`**).
- **`UIElementCollection.Insert`** — insert child at index (items host updates).

### Bug fixes

- **B4 (partial, ItemsControl):** non-**`ObservableCollection`** **`ItemsSource`** raises **`Err`** with clear message (ListView unchanged this release).

### Test

- **P4b-ICtrl** in `.Tests/Phase0` (suite → **21** tests).

---

## [2.6.0] — 2026-06-20 — Phase 4b (collections — validated)

Tag: **`v2.6.0-wpf-alignment-p4b`** · Phase0 **20/20** pass.

### Added

- **`ObservableCollection.BeginUpdate` / `EndUpdate`** — batch mutations; coalesce to a single **`Reset`** notification on `EndUpdate`.
- **`ObservableCollection.Move(OldIndex, NewIndex)`** — raises **`CollectionChangedActionMove`**.
- **`ObservableCollection.IsUpdating`** — read-only batch depth indicator.
- **`modCollectionNotifications`** — reusable single-item **`List`** scratch buffers for Add/Remove/Replace/Move (avoids `New List` per single-item change).

### Bug fixes

- **B2 (partial):** **`ListView`** now handles **`CollectionChangedActionMove`** (item template reorder).

### Notes

- Multi-item **`AddRange`** / **`Clear`** still allocate **`List`** payloads where required; single-item paths use scratch buffers.
- **`ObservableDictionary`** unchanged this release (same notification pattern as before).

### Test

- **P4b-DEFER**, **P4b-MOVE** in `.Tests/Phase0` (with Phase 0–4 suite → **20** tests).

---

## [2.5.0] — 2026-06-20 — Phase 4 (bindings — validated)

Tag: **`v2.5.0-wpf-alignment-p4`** · Phase0 **18/18** pass.

### Added

- **`BindingExpression`** — `Attach`, `Detach`, `UpdateTarget`; wraps legacy `Binding` graph (transitional).
- **`modBindingExpressions`** — `RefreshTargetBindings`, `DetachTargetBindings`; `OnDataContextChanged` hook on `DataContext` DP change.
- **`Binding.IsListenerActive`**, **`Binding.DetachBinding`** — deterministic teardown of listeners, callbacks, and INPC `WithEvents` subscriptions.

### Bug fixes

- **Binding detach hang:** `DependencyProperty` listeners/callbacks now stored as **object references** (not `ObjPtr`); `GetValue` no longer revives stale pointers after `Detach`. Fixes IDE freeze in **P4-DETACH** when reading a bound target property after source INPC.
- **`GetValue` re-entrancy guard** — prevents recursive listener fan-out during effective-value resolution.

### Notes

- Legacy **`Binding`** remains in use; **`BindingsManager`** unchanged this release.
- **`BindingExpression`** entries stored in control **`Bindings`** list alongside legacy **`Binding`** objects.

### Test

- **P4-BIND**, **P4-DCTX**, **P4-DETACH** in `.Tests/Phase0` (**18/18** with Phase 0–3).

---

## [2.4.0] — 2026-06-20 — Phase 3 (resources)

### Breaking (when `StrictXamlLoad = True`)

- **`IApplication.Resources`** and **`IUIElement.Resources`** are now **`ResourceDictionary`** (was **`ObservableDictionary`**). **Migration:** Change property types; use **`Resources.LocalResources`** where flat dictionary access is required; **`Merge`** / **`MergedDictionaries`** for WPF-style includes.
- **Unknown XAML attributes** on **`IDependencyObject`** types raise **`XamlLoadException`** instead of **`CallByName`** / widget fallback. **Migration:** Use registered dependency properties only; set **`VCF.StrictXamlLoad = False`** temporarily for legacy XAML.

### Added

- **`ResourceDictionary`** — local resources + **`MergedDictionaries`** + lazy **`Source=`** load.
- **`XamlResourceResolver`** — load dictionary fragments from disk (`BasePath` + relative **`Source`**).
- **`DynamicResourceExtension`** — **`{DynamicResource Key}`** markup; **`{ThemeResource}`** routes here (deprecated alias).
- **`XAMLReader.LoadElement`** — public node instantiation for resource entries.
- **`Application.Resources` / element tree** — merged lookup via **`TryGetResource`**.

### Deprecated

- **`ThemeResource`** class — use **`{DynamicResource}`** in new XAML.
- Flat **`ObservableDictionary`** on **`Application.Resources`** — use **`ResourceDictionary`** + merge.

### Test

- **P3-MERGE**, **P3-SOURCE**, **P3-DYNAMIC**, **P3-STRICT-PROP** in `.Tests/Phase0` (target **15/15** with Phase 0–2).

---

## [2.3.0] — 2026-06-20 — Phase 2 (panels)

### Added

- **`StackPanel`** — vertical/horizontal stack layout (`Orientation`); `LegacyScaleLayout` off by default.
- **`Grid`** — `RowDefinitions` / `ColumnDefinitions`, `Grid.Row` / `Grid.Column` / span attached props, `*` / `Auto` / pixel tracks.
- **`ContentControl`** — single-content host with decorator arrange.
- **`RowDefinition`**, **`ColumnDefinition`** — grid track specs for XAML.
- **`Border.Child`** DP — decorator semantics; single child fills client (multi-child legacy arrange retained).
- **`UniformGrid`** — **Collapsed** children skipped in cell assignment (B9 partial).

### XAML

- `<StackPanel Orientation="Vertical|Horizontal" Width Height>`
- `<Grid>` with `<Grid.RowDefinitions>` / `<Grid.ColumnDefinitions>`
- `<ContentControl>` with one visual child
- `Grid.Row`, `Grid.Column` attached properties on **Grid** children

### Test

- **P2-STACK**, **P2-STACK-LAY**, **P2-GRID** in `.Tests/Phase0` (target **11/11** with Phase 0/1).

---

## [2.2.0] — 2026-06-20 — Phase 1b (layout core — shell controls)

### Added

- **`FrameworkElement`** on **Border**, **UserControl**, **Window**, **Button** — layout DPs (`Width`, `Height`, `Margin`, `Visibility` where applicable).
- **Window** child layout uses form client scale (`Form.ScaleWidth` / `ScaleHeight`) via `ArrangeChildren` overrides.
- **Registry types:** `Border`, `UserControl`, `Button`, `Window` extend `FrameworkElement`.

### XAML

- `Width` / `Height` accepted on migrated types (Panel, Border, UserControl, Window; Button width/height via DPs).
- `DesignWidth` / `DesignHeight` still accepted (alias when layout DPs registered).

### Notes

- **Button** retains custom `MoveChild` inset logic for content/overlay; layout DPs drive scale factors.
- **Phase 1 compile/runtime fixes** included (ByRef registry, `IsWidgetVisible`, UDT ByRef, init order, `Empty` reserved word).

### Test

- **P1-BORDER** — Border `Width="320"` from XAML (`.Tests/Phase0`).

---

## [2.1.0] — 2026-06-20 — Phase 1 (layout core — partial)

### Added

- **`DependencyPropertyRegistry`** — shared DP metadata per type; `ApplyTo` registers layout properties once per instance.
- **`FrameworkElement`** — composed layout helper (Measure/Arrange, Visibility DP, legacy scale layout default).
- **`modLayoutEngine`** — layout rects, `Design*` XAML alias helper, collapsed visibility checks.
- **`Panel`** — first control migrated: `Width`, `Height`, `Margin`, `Visibility` DPs; `DesignWidth`/`DesignHeight` forward to `Width`/`Height`.

### XAML (transitional — non-breaking)

- **`DesignWidth` → `Width`**, **`DesignHeight` → `Height`** when target type registers layout DPs (Panel today; more controls in 2.2.x).
- **`DesignLeft` / `DesignTop`** unchanged — still scale-layout until Margin-based arrange (set `FrameworkElement.LegacyScaleLayout = False` when ready).

### Deprecated

- Public **`DesignWidth` / `DesignHeight`** on migrated controls — use **`Width` / `Height`** in new XAML and VB6.
- **`Visible` bool DP** on Button — migrate to **`Visibility`** enum when Button is migrated (Phase 1b).

### Behavior

- **`Visibility=Collapsed`** on Panel — child omitted from widget tree (layout-aware); Hidden still hides widget (legacy Cairo; full Hidden semantics Phase 2).

### Not yet in 2.1.0

- Button, Border, Window, UserControl, UniformGrid migration to `FrameworkElement`.
- Full removal of public `Design*`.
- `DependencyProperty` shared instance store (registry metadata only in 2.1.0).

---

## [Unreleased] — planned (Phases 1b–7)

See [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md) and [VCF_BREAKING_CHANGES_TEMPLATE.md](./VCF_BREAKING_CHANGES_TEMPLATE.md).

---

## Release template

```markdown
## [X.Y.Z] — YYYY-MM-DD

### Breaking

- **Area:** Description. **Migration:** one-line fix.

### Deprecated (remove in X+1)

- ...
```
