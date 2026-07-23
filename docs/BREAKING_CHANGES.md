# Demac.VCF — breaking changes log

**Program:** WPF alignment (Phases 0–7)  
**Maintained by:** VCF team · POS migration steps in [MIGRATION.md](./MIGRATION.md)

---

## [3.23.0] — 2026-07-24 — TextBlock / Image Visibility

### Added
- **`TextBlock`** / **`Image`** — `Visibility` DP (Visible / Hidden / Collapsed) + legacy **`Visible`** bool sync (`False` → Collapsed), matching Button / UniformGrid.
- TextBlock change invalidates parent layout (Hidden reserves slot; Collapsed reclaims).
- Phase0 **P1-VIS-TB**, **P1-VIS-IMG**. Suite **87**.
- Docs: stale “still deferred” notes corrected; thin [CHANGELOG.md](../CHANGELOG.md).

- DLL version **3.23.0**.

## [3.22.0] — 2026-07-24 — Attached layout DP invalidation

### Added
- Changing attached layout DPs (`Grid.Row`/`Column`/`Span`, `DockPanel.Dock`, `Canvas.Left`/`Top`) via **SetValue** / binding now calls **`InvalidateParentLayout`** so the parent panel rearranges.
- **`IsAttachedLayoutProperty`** in `modLayoutEngine` — central name list for the hook in `DependencyProperties.OnDependencyPropertyChanged`.
- Phase0 **P4-BIND-LAY** (bind `Grid.Row`, INPC, assert widget **Top** moves). Suite **85**.

- DLL version **3.22.0**.

## [3.21.0] — 2026-07-24 — Canvas panel

### Added
- **`Canvas`** — absolute child positioning via **`Canvas.Left` / `Canvas.Top`** (Double DPs, default 0).
- **`GetLeft` / `SetLeft` / `GetTop` / `SetTop`**, Measure/Arrange bounding-box content size.
- **`EnsureAttachedProperty`** uses registry **PropertyType** (so Canvas Left/Top register as `vbDouble`, not Long).
- Phase0 **P2-CANVAS-XAML**, **P2-CANVAS-LAY**. Suite **83**.

- DLL version **3.21.0**.

## [3.20.0] — 2026-07-23 — Path=(Attached.Prop) source syntax

### Added
- **`NestedProperty`** parses WPF **`Path=(Grid.Row)`** / **`(DockPanel.Dock)`** — reads/writes the source’s attached DP bag (EnsureAttached + GetValue/SetValue).
- Subscribes to the attached DP’s **PropertyChanged** for live OneWay updates.
- Phase0 **P4-BIND-APATH**. Suite **81**.

- DLL version **3.20.0**.

## [3.19.0] — 2026-07-23 — Binding to attached DPs

### Added
- **`EnsureAttachedTargetIfRegistered`** — before binding Attach/`CreateBinding`, lazy-registers registered attached names (`Grid.Row`, `DockPanel.Dock`, …) on the target bag.
- Enables **`Expr.Attach Target, "Grid.Row", …`** and XAML **`Grid.Row="{Binding RowIndex}"`** without a prior `SetRow`/`EnsureAttached`.
- Phase0 **P4-BIND-ATTACH**. Suite **80**.

### Notes
- OneWay **to** attached target DPs. ~~Source `Path=(Grid.Row)` parentheses syntax still deferred~~ → **3.20.0**.

- DLL version **3.19.0**.

## [3.18.0] — 2026-07-23 — MeasureOverride / ArrangeOverride surface

### Added
- **`MeasureOverride` / `ArrangeOverride`** on layout hosts (`StackPanel`, `Grid`, `DockPanel`, `Border`, `ContentControl`, `UniformGrid`, `Panel`) — WPF-named duck-typed entry points (VB6 has no FE inheritance).
- **`MeasureLayout`** remains a one-line **compat alias** → `MeasureOverride`.
- **`MeasureElementSize`** prefers `CallByName` **`MeasureOverride`**, falls back to **`MeasureLayout`**.
- Phase0 **P2-MEAS-OVR**. Suite **79**.

- DLL version **3.18.0**.

## [3.17.0] — 2026-07-23 — DockPanel

### Added
- **`DockPanel`** — WPF-style dock layout with **`LastChildFill`** (default **True**).
- **`DockPanel.Dock`** attached (`Left` / `Top` / `Right` / `Bottom`, default **Left**) — DP bag + dict shim (same pattern as Grid.*).
- **`Dock` enum**, `ParseDock`, `GetDock` / `SetDock`, `MeasureDockPanelContent` / `ArrangeDockPanelChildren`.
- XAML `DockPanel.Dock="Top"` coerced via `ParseDock` (not `Val`).
- Phase0 **P2-DOCK-XAML**, **P2-DOCK-LAY**. Suite **78**.

- DLL version **3.17.0**.

## [3.16.0] — 2026-07-23 — Grid attached props in per-element DP bag

### Added
- **`EnsureAttachedProperty`** — lazy-registers `Grid.Row` / `Column` / `RowSpan` / `ColumnSpan` on the target’s `DependencyProperties` (metadata default from `RegisterAttached`).
- **`SetGridAttachedLong` / XAML `Grid.*`** — write DP bag + keep nested `AttachedProperties("Grid")` shim.
- **`GetGridAttachedLong`** — prefers DP `GetValue` (supports **`ClearValue`** → default) then falls back to dict.
- Phase0 **P2-GRID-ATTACH-DP**. Suite **77**.

### Notes
- Not eager-registered on every instance; no full attached-binding epic yet.

- DLL version **3.16.0**.

## [3.15.0] — 2026-07-23 — Grid cell Horizontal/VerticalAlignment

### Added
- **`HorizontalAlignment` / `VerticalAlignment`** on FrameworkElement layout hosts (string: `Left`/`Center`/`Right`/`Stretch`, `Top`/`Center`/`Bottom`/`Stretch`; default **Stretch**).
- **`ArrangeGridChildren`** positions non-Stretch children from Desired/explicit size within the cell (WPF-style).
- TextBlock keeps Long text-align DPs; **layout always Stretch** in Grid (text align ≠ layout align).
- **`TextBlock.Move`** no longer writes arranged size into Width/Height DPs (avoids sticky narrow wrap / one-glyph-per-row).
- Stretch + explicit Width/Height keeps author size (does not expand fixed children to the cell).
- **`Button`** default Width/Height is **0** (unset), not 100×30 — Stretch in Grid/UniformGrid fills the slot.
- Phase0 **P2-GRID-ALIGN**. Suite **76**.

- DLL version **3.15.0**.

## [3.14.0] — 2026-07-23 — Retire design-canvas scale bridge

### Breaking
- **`FrameworkElement.ArrangeChildren`** no longer scales absolute `Margin`/`Width`/`Height` by `hostWidget / hostDP` on **UserControl**, multi-child **Border**, or **Panel**. Those values are **fixed pixels**.
- Absolute POS trees that relied on canvas scale will not resize proportionally — migrate to **Grid** / **StackPanel** / single-child **Border** (see [POS_LAYOUT_MIGRATION.md](./POS_LAYOUT_MIGRATION.md)).

### Added / changed
- **UserControl** with a **single root** child fills the UC client (typical migrated screen = root `Grid`).
- Multi-child **Border** comment: fixed absolute; no scale.
- Phase0 **P7d-LAY-PANEL** replaces **P7d-LAY-RESIZE** (panel star geometry on resize).
- DeNovoSmoke Login/Sales load **Migrated** XAML under `Resources/XAML/Migrated/`.
- **Layout default Width/Height** for panel hosts is **0** (unset), not 300 — so Grid `Auto` tracks measure content instead of a 300px phantom size (Login star-row starvation).

- DLL version **3.14.0**.

## [3.13.0] — 2026-07-23 — Measure/Arrange UniformGrid slice

### Added / changed
- **`UniformGrid.MeasureLayout`** — max child cell × Rows/Columns when Width/Height unset; Desired/Actual exposed.
- Arrange (`MoveChildren`) measures first then fills equal cells; public `RelayoutChildren`.
- Phase0 **P2a-UG-MEAS**. Suite **75**.

### Not in this release
- Retiring UC/multi-child Border/Panel canvas-scale bridge.

- DLL version **3.13.0**.

## [3.12.0] — 2026-07-23 — Measure/Arrange ContentControl slice

### Added / changed
- **`ContentControl.MeasureLayout`** — IUIElement Content child via `MeasureDecoratorContent`; string/empty uses `FrameworkElement.Measure`.
- Arrange path Measure→`ArrangeDecoratorChild`→`ReportActualSize`; Desired/Actual exposed.
- Phase0 **P6e-CC-MEAS**. Suite **74**.

### Not in this release
- ~~UniformGrid MeasureLayout~~ → **3.13.0**; ~~retiring UC/multi-child Border/Panel canvas-scale bridge~~ → **3.14.0**.

- DLL version **3.12.0**.

## [3.11.0] — 2026-07-23 — Measure/Arrange Border decorator slice

### Added / changed
- **`Border.MeasureLayout`** — single-child Desired = child Desired + Margin insets (`MeasureDecoratorContent`); multi-child falls back to `FrameworkElement.Measure`.
- Arrange path Measure→decorator fill / canvas-scale→`ReportActualSize`; Desired/Actual exposed.
- Phase0 **P1-BORDER-MEAS**. Suite **73**.

### Not in this release
- ~~Retiring multi-child Border / UC / Panel canvas-scale bridge~~ → **3.14.0** (`P7d-LAY-PANEL`).

- DLL version **3.11.0**.

## [3.10.0] — 2026-07-23 — Measure/Arrange Grid slice

### Added / changed
- **`Grid.MeasureLayout`** — content-driven measure via `MeasureGridContent` (same track solver as arrange); exposes Desired/Actual; arrange path Measure→Arrange→`ReportActualSize`.
- **`GridAutoTrackSize`** — uses `MeasureElementSize` (nested StackPanel/Panel Desired*) instead of Width/Height DPs only.
- Phase0 **P2-GRID-MEAS** (Auto+Pixel+Star Desired/Actual). Suite **72**.

### Not in this release
- Full WPF Auto multi-pass / span redistribution; ~~retiring UC/Border/Panel canvas-scale bridge~~ → **3.14.0**.

- DLL version **3.10.0**.

## [3.9.0] — 2026-07-23 — Measure/Arrange first slice (StackPanel)

### Added / changed
- **`FrameworkElement`** — `DesiredWidth`/`DesiredHeight`, `ActualWidth`/`ActualHeight`, `InvalidateMeasure`/`InvalidateArrange`, dirty flags; layout DP changes invalidate measure and parent layout.
- **`StackPanel.MeasureLayout`** — content-driven measure (`MeasureStackPanelContent`); arrange path measures then positions children from desired sizes (vertical unset height stays 0, not host height).
- **`Panel.MeasureLayout`** — forwards to `FrameworkElement.Measure`; exposes Desired/Actual.
- Phase0 **P2-STACK-MEAS** (Desired/Actual assertions). Suite **71**.

### Not in this release
- ~~Grid MeasureOverride~~ → **3.10.0**; ~~retiring the UC/Border/Panel canvas-scale bridge~~ → **3.14.0**.

- DLL version **3.9.0**.

## [3.8.0] — 2026-07-23 — ContentTemplate on ContentControl

### Added

- **`ContentControl.ContentTemplate`** (`DataTemplate` DP) — clones template visual via `modItemTemplateEngine`; `Content` is the clone item/`DataContext` when object.
- **`ContentHost.ApplyContentVisual`** — precedence: explicit children → ContentTemplate clone → string/`Content` presenter paint.
- Phase0 **P6e-CTMPL** — suite **70**.

### Notes

- No `ContentTemplateSelector`; Button default chrome unchanged (self-paint). ItemContainerStyle unify / LCV sort-filter still deferred.
- Layout contract (canvas-scale bridge vs Measure/Arrange) documented in [POS_LAYOUT_RESIZE.md](./POS_LAYOUT_RESIZE.md) §0.
- DLL version **3.8.0**.

---

## [3.7.0] — 2026-07-23 — DeNovo layout bridge + hover / TextBox DPs

### Changed (layout)

- **Design-canvas scale bridge** — `FrameworkElement.ArrangeChildren` scales absolute `Margin`/`Width`/`Height` children when the host widget size differs from host `Width`/`Height` DPs, for **UserControl**, multi-child **Border**, and **Panel**. Restores proportional resize for absolute POS/DeNovo trees after **3.0.0** removed `Design*` / `LegacyScaleLayout`.
- **P7d-LAY-RESIZE** — expects half-size scaled positions (400×300 → 200×150), not frozen absolute pixels.
- Preferred long-term path remains Grid / StackPanel / single-child Border fill — see [POS_LAYOUT_RESIZE.md](./POS_LAYOUT_RESIZE.md).

### Added

- **`TextBox`** — `Width` / `Height` / Min/Max layout DPs (Grid `Auto` track measure; fixes Login password vs numpad overlap when only widget `Height` was set).

### Fixed

- **`Button.IsMouseOver`** — `W.Refresh` on change (3.2.1 Notify no-ops without triggers; enter paint was delayed vs leave).
- **`Button` MouseLeave** — ignore leave when pointer enters a descendant content widget (Image/TextBlock).
- **`StringConversion`** — Greek accent maps via `ChrW` (corrupted `?` maps caused duplicate-key spam).
- **`NamingManager` / `Window.AddNamedItem`** — soft-skip duplicate names.

### DeNovoSmoke

- Harness XAML migrated off `Design*`; LoginPad uses `Content=` captions; Pay stub enabled for no-Command hover check; Clock In stays `Enabled="False"`.

### Notes

- DLL version **3.7.0**.
- Suggested tag: `v3.7.0-denovo-layout-hover`.

---

## [3.6.0] — 2026-07-23 — SelectionChanged naming

### Added

- **`SelectionChanged`** event on **`ListView`** and **`ListViewBase`** (WPF name; parameterless in this release).
- Raised together with **`ListIndexChanged`** at every existing raise site (user / keyboard / `ListIndex` Let).
- Phase0 **P4d-SELCHG** — suite **69**.

### Soft deprecation

- Prefer **`SelectionChanged`** for new code. **`ListIndexChanged`** remains raised as a compat alias (no hard break).

### Notes

- DP-driven **`SelectedIndex=`** still uses silent base index (no event) — same as before for `ListIndexChanged`.
- No `SelectionChangedEventArgs` / AddedItems yet.
- DLL version **3.6.0**.

---

## [3.5.0] — 2026-07-23 — Min/Max layout constraints

### Added

- **`MaxWidth` / `MaxHeight`** DPs on FrameworkElement (with existing **`MinWidth` / `MinHeight`**).
- Layout clamp in Measure + StackPanel / Grid / Panel / decorator arrange: `size = Max(Min, size)`; if `Max > 0` then `size = Min(Max, size)`.
- **`Max* = 0` means unbounded** (no WPF Infinity).
- Public Min/Max Get/Let on Panel, Border, Button, StackPanel, Grid.
- Phase0 **P1-MINMAX-MIN**, **P1-MINMAX-MAX** — suite **68**.

### Notes

- Additive for apps that never set Max*; setting Max can shrink arranged size vs prior builds.
- DLL version **3.5.0**.

---

## [3.4.0] — 2026-07-23 — RegisterAttached (Grid)

### Added

- **`DependencyPropertyRegistry.RegisterAttached`** — metadata for attached props (`IsAttachedRegistered`, `HasAttachedOwner`, `GetAttachedDefault`).
- Built-in: **`Grid.Row`**, **`Grid.Column`**, **`Grid.RowSpan`**, **`Grid.ColumnSpan`** (Long; defaults 0/0/1/1).
- **`Grid.GetRow` / `SetRow`**, **`GetColumn` / `SetColumn`**, **`GetRowSpan` / `SetRowSpan`**, **`GetColumnSpan` / `SetColumnSpan`** — WPF-shaped accessors; storage remains `AttachedProperties("Grid")`.
- **`modLayoutEngine.SetGridAttachedLong`** — shared write path with `GetGridAttachedLong`.
- Phase0 **P2-GRID-ATTACH**, **P2-GRID-XAML** — suite **66**.

### Breaking (strict only)

- With **`StrictXamlLoad=True`**, unknown attached attributes under an owner that has registered attached props (e.g. `Grid.Foo`) raise **`XamlLoadException`**. Non-strict and unknown owners unchanged.

### Notes

- Does **not** migrate attached values into the per-element DP bag (deferred).
- DLL version **3.4.0**.

---

## [3.3.0] — 2026-07-23 — Visibility / Hidden / Collapsed layout

### Breaking

- **`Visible=False` → `Visibility=Collapsed`** on **Button** and **UniformGrid** (and any element where layout reads the legacy `Visible` bool). Previously `Visible=False` only flipped `W.Visible` while the `Visibility` DP stayed **Visible**, so panels kept the layout slot. **Migration:** use `Visibility=Hidden` when you need an invisible element that still reserves space; keep `Visible=False` / `Visibility=Collapsed` when the parent should reclaim space (POS UniformGrid menus).

### Added / Changed

- WPF layout semantics gated: **Visible** = paint + slot; **Hidden** = no paint, slot kept; **Collapsed** = no paint, slot removed.
- **`modLayoutEngine.InvalidateParentLayout`** — child Visibility changes rearrange StackPanel / Grid / UniformGrid / Panel / Border / ContentControl.
- Public **`Visibility`** on **Button** / **Border**; **`Visible`** facade on Button / UniformGrid synced to Visibility.
- Phase0 **P1-VIS-HIDDEN**, **P1-VIS-COLLAPSE**, **P1-VIS-BOOL** — suite **64**.

### Notes

- Does **not** remove the `Visible` bool (compat). ~~TextBlock / Image Visibility deferred~~ → **3.23.0**.
- DLL version **3.3.0**.

---

## [3.2.2] — 2026-07-23 — PropertyTrigger conditions → ReapplyStyleValues

### Changed

- **`modStyleTriggerEngine.NotifyConditionPropertyChanged`** — if the target Style watches that condition, calls **`ReapplyStyleValues`** (reentrancy-guarded).
- **`DependencyProperties`** — after every DP change (except while style reapply is in progress), notifies trigger engine so **any DP used as a Trigger Property** reactivates/deactivates correctly.
- **`Style.WatchesTriggerCondition`** — includes `BasedOn` chain.
- **`Button.IsMouseOver`** — uses Notify (non-DP condition) instead of a private reapply helper.
- **P6b-TRIG** — also gates **Selected** DP trigger activate/restore.

### Notes

- Closes interim **B3** overlap under the **3.2.0** two-slot model.
- **WPF-faithful precedence layers** remain deferred — [VCF_DP_PRECEDENCE_ROADMAP.md](./VCF_DP_PRECEDENCE_ROADMAP.md).
- DLL version **3.2.2**.

---

## [3.2.1] — 2026-07-23 — Style reapply vs full ApplyStyle

### Changed

- **`StyleManager.ReapplyStyleValues`** — style setters + active triggers + refresh only (3.2.0 trigger lifecycle).
- **`StyleManager.ApplyStyle`** — still full path including **ControlTemplate** + content-align (Style DP change).
- **`Button.IsMouseOver`** — calls **`ReapplyStyleValues`** so hover does not re-clone lookless chrome.
- Phase0 **P6b-TRIG** asserts ContentPresenter instance stable across hover.

### Notes

- Non-breaking; aligns interim trigger restore with locked two-slot Current model.
- Full WPF precedence stack remains deferred — [VCF_DP_PRECEDENCE_ROADMAP.md](./VCF_DP_PRECEDENCE_ROADMAP.md).
- DLL version **3.2.1**.

---

## [3.2.0] — 2026-07-23 — DP precedence contract (ClearValue + metadata default)

### Locked contract (conflict #4)

Effective value order (high → low):

1. **Local** — `SetValue` (XAML, CLR, Binding, TemplateBinding)
2. **Current** — `SetCurrentValue` (style setters, triggers, init seeds, template align) — last write wins *inside* this slot
3. **Inherit** — lazy parent walk when both slots unset and `IsInheritable`
4. **Metadata `DefaultValue`** — when still unset

No separate binding / trigger / animation layers yet (B3+ deferred).

### Added

- **`DependencyProperties.ClearValue`** — clears **local** slot only (WPF `ClearValue`); style/current, inherit, or metadata default can reappear.
- **`DependencyProperties.ReadLocalValue`** — raw local slot (Unset sentinel when none).
- **`DependencyPropertyMetadata.DefaultValue` / `HasDefaultValue`** — `NewDependencyPropertyMetadata(..., DefaultValue)`.
- Phase0 **P4-PREC** — suite **61** tests (includes prior under-count fix in Done tally).

### Changed / breaking

- **`Binding.Detach` / `BindingExpression.Detach`** — after removing the listener, **clears the target local value** (WPF `ClearBinding` intent). Prior behavior left the last transferred value frozen (`P4-DETACH` updated).
- **`Window.BorderStyle`** — default `vbSizable` is **metadata DefaultValue**, not `SetCurrentValue`.
- **`ContentControl.Content`** — metadata default `""`.

### Notes

- Friend `DependencyProperty.ClearValue` no longer clears the current slot.
- Migrating remaining `Class_Initialize` `SetCurrentValue` defaults → metadata is incremental (not all controls in this release).
- **Trigger lifecycle (interim):** condition changes re-run **`ApplyStyle`** (style setters, then active triggers only) — see **P6b-TRIG** / Button `IsMouseOver`. Not a separate trigger layer.
- **Full WPF effective-value stack** — accepted roadmap, **deferred (not now)**: [VCF_DP_PRECEDENCE_ROADMAP.md](./VCF_DP_PRECEDENCE_ROADMAP.md).
- DLL version **3.2.0**.

---

## [3.1.0] — 2026-07-22 — IContentControl + ContentHost

### Added

- **`IContentControl`** — WPF ContentControl contract (`Content`, `ContentPresenter`, `SyncContentPresenter`). Implemented by **`ContentControl`** and **`Button`** (VB6 has no class inheritance).
- **`ContentHost`** — shared composition helper (host presenter, optional template presenter slot, content/suppress/alignment sync, `ApplyContentValue` for IUIElement Content).

### Changed

- **`ContentControl` / `Button`** — compose `ContentHost` instead of duplicating presenter sync logic. Public APIs unchanged (`Content`, `ContentPresenter`, `SyncContentPresenter`).
- Phase0 **P6e-CC** / **P6e-PRES** assert `TypeOf … Is IContentControl`.

### Notes

- Non-breaking for existing callers; new interface is additive.
- ~~**`ContentTemplate`** still deferred~~ → **3.8.0** (`ContentControl` + `ContentHost.ApplyContentVisual`).
- DLL version **3.1.0**.

### Follow-up decision (docs only — 2026-07-23)

- **ListView dual presenter** locked permanent (bound vs owner-draw). No API change. See [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md) §8.4.

---

## [3.0.0] — 2026-07-22 — Remove Design* / LegacyScaleLayout

### Breaking

- **`DesignLeft` / `DesignTop` / `DesignWidth` / `DesignHeight` removed** from `IUIElement`, `IUserControl`, and all implementers.
- **`LegacyScaleLayout` removed** from `FrameworkElement` — no host design-canvas scale (`left = DesignLeft × host/design`).
- **Border Option B** (multi-child Design* detection → scale) removed.
- **XAML:** `Design*` attributes are no longer accepted (no alias to Width/Height; no Margin→Design* shim).
- **`ApplyLegacyLayoutProperty` / `LayoutRectFromDesign` / `MarginFromDesignWhenUnset` removed.**

### Migration

| Old | New |
|-----|-----|
| `DesignWidth` / `DesignHeight` | `Width` / `Height` |
| `DesignLeft` / `DesignTop` | `Margin` (Left/Top) |
| Absolute Design* canvas + scale | Prefer `Grid` / `StackPanel` / `Border` composition; resize via panel layout, not uniform X/Y scale |

### Added / changed (supporting)

- **`TextBlock` / `UniformGrid`** — `Width` / `Height` / `Margin` DPs (required once Design* sizing was removed).
- Phase0 fixtures and **P7c-LAY** / **P7d-LAY-RESIZE** rewritten for Margin/Width absolute layout (no Design* scale).

### Notes

- Paint/rendering paths unchanged; only positioning/sizing contract.
- Option C (parent-scale propagate) cancelled — not needed without Design*.
- DLL version **3.0.0**.

---

## [2.43.0] — 2026-07-22 — Phase 2a ({TemplateBinding} markup)

### Added

- **`TemplateBinding`** — WPF shorthand for `{Binding Path, RelativeSource=TemplatedParent, Mode=OneWay}`; markup `{TemplateBinding Content}` + imperative `Attach`.
- **`ContentPresenter.Content`** — registered as `vbVariant` DP (enables Binding/`TemplateBinding` onto the content slot).
- Phase0 **P6k-TBMK** — gate **58/58**.

### Fixed

- **`ResolveRelativeSource`** — `API.CObj` widen so `TemplatedParent`/`Parent` resolve when the target is passed as `IDependencyObject`.

---

## [2.42.0] — 2026-07-22 — Phase 2a (CanExecuteChanged → Button.Enabled)

### Added

- **`INotifyCanExecuteChanged`** — optional companion to `ICommand` exposing `CanExecuteChangedEvent` (non-breaking for existing `ICommand` implementers).
- **`Button`** — when `Command` implements `INotifyCanExecuteChanged`, listens and syncs `Widget.Enabled` from `CanExecute` (also on Command/CommandParameter change).
- Phase0 **P4-CCMD** — gate **57/57**.

### Changed

- **`Phase0DialogCommand`** — mutable `CanExecute` + raises CanExecuteChanged (still defaults True for P7c-DLG).

---

## [2.41.0] — 2026-07-22 — Phase 2a (TextBox caret preserve)

### Changed

- **`TextBoxBase.Text` Let** — when the new value is a pure extension of the old text, preserve caret (end stays at end; mid caret kept). Unrelated replaces still reset selection to 0 (prior behavior).
- **`TextBox.SelStart` / `SelLength`** — exposed wrappers over `TextBoxBase` (were commented out).

### Added

- Phase0 **P4-CARET** — gate **56/56**.

### Notes

- Completes the barcode/typing hardening arc with **P4-UDELAY** (debounce) + caret-safe write-back.

---

## [2.40.0] — 2026-07-22 — Phase 2a (UpdateSourceDelay debounce)

### Added

- **`Binding.UpdateSourceDelay`** (ms) — `0` = immediate PropertyChanged push (default); `>0` debounce Target→Source only.
- Debounce timer + **`FlushUpdateSource`** / **`FlushSourceBindings`** — TextBox **Enter** and **LostFocus** force-flush pending TwoWay bindings before app handlers.
- Round-trip guards **`IsUpdatingSource` / `IsUpdatingTarget`** to reduce caret write-back loops.
- Markup: `{Binding … UpdateSourceDelay=50}`.
- Phase0 **P4-UDELAY** — gate **55/55**.

### Notes

- Opt-in only (no Text metadata default delay) — non-breaking for POS. Barcode guidance: 30–80 ms + Enter flush.

---

## [2.39.0] — 2026-07-22 — Phase 2a (UpdateSourceTrigger)

### Added

- **`Binding.UpdateSourceTrigger`** — `ustPropertyChanged` / `ustLostFocus` / `ustExplicit` / `UpdateSourceTriggerDefault` (Default ≡ PropertyChanged; non-breaking). Markup still accepts `LostFocus` / `PropertyChanged` / `Explicit`.
- **`Binding.FlushUpdateSource`** / **`BindingExpression.UpdateSource`** — push Target→Source on demand.
- **`FlushLostFocusBindings`** — TextBox LostFocus flushes TwoWay bindings with `LostFocus` trigger.
- Markup: `{Binding … UpdateSourceTrigger=LostFocus}`.
- Phase0 **P4-UST** — gate **54/54**.

### Notes

- `UpdateSourceDelay` (barcode debounce) remains a follow-up; this slice is trigger modes + flush only.

---

## [2.38.0] — 2026-07-22 — Phase 2a (ElementName → Command)

### Added

- Phase0 **P4-ECMD** — named host `ICommand` bound to `Button.Command` via ElementName (window-level Command pattern). Gate **53/53**.

### Notes

- Uses existing `NamingManager.ResolveElementName` + `BindingExpression`; no new framework API beyond the RelativeSource/ElementName slice.

---

## [2.37.0] — 2026-07-22 — Phase 2a (RelativeSource / ElementName)

### Added

- **`RelativeSource`** — markup extension + modes `Self`, `TemplatedParent`, `FindAncestor` (`AncestorType` / `AncestorLevel`).
- **`modBindingSourceResolve`** — `ResolveRelativeSource` / `ResolveElementName` / `FindTemplatedParent` / `FindAncestor`.
- **`BindingsManager`** — honors `RelativeSource=` and `ElementName=` in `{Binding}` markup (alongside `Source` / DataContext).
- **`TemplatedParent`** — stamped on live template clones (`Border` / `Grid` / `StackPanel` / `ContentPresenter`).
- Phase0 **P4-RSELF**, **P4-ENAME**, **P4-RTP** — gate **52/52**.

### Notes

- `{Self}` remains the existing `SelfBinding` shortcut; `{RelativeSource Self}` is the WPF-aligned form for binding Source resolution.
- Window-level Command via ElementName is now unblocked for app markup; Phase 2b POS pin still deferred.

---

## [2.36.0] — 2026-07-22 — Phase 2a (multi-node ControlTemplate tree)

### Changed

- **`ApplyButtonTemplate`** — recursively clones template visuals: `Border.Child` may be `ContentPresenter`, `Grid`, or `StackPanel` with a deeper presenter. Flat Border+CP siblings remain supported.
- **`Button.AttachTemplatePresenter`** — optional `NestUnderChrome`; when False, presenter is already in the cloned subtree (do not overwrite `Border.Child`).

### Added

- Phase0 **P6j-MULTI** — gate **49/49**.

### Notes

- RelativeSource / ElementName shipped in **2.37.0**.

---

## [2.35.0] — 2026-07-22 — Phase 2a (OS System theme → Light/Dark)

### Changed

- **`ThemesManager`** — `ActiveThemeName = "System"` resolves to OS preferred Light/Dark (`AppsUseLightTheme` registry) and merges that bag. `SystemThemeOverride` forces resolution for tests/harness. `EffectiveThemeName` / `RefreshSystemTheme` expose the concrete key and re-publish.

### Added

- Phase0 **P2a-THEME-OS** — gate **48/48**.

### Notes

- ~~Deeper multi-node templates and RelativeSource/ElementName remain later/deferred~~ → **P6j-MULTI** / **P4-RSELF** / **P4-ENAME** / **P4-RTP** (2.x–3.x).

---

## [2.34.0] — 2026-07-22 — Phase 2a (theme dictionary merge/swap)

### Changed

- **`ThemesManager`** — on `ActiveThemeName` change, publishes the active theme into the host `ResourceDictionary.MergedDictionaries` (WPF-style swap). Supports theme bags as `ObservableDictionary` (copied into a transient RD) or `ResourceDictionary` (merged directly). `AttachToResources` wires the host (Application or test).
- **`Application`** — when a `ThemesManager` is added to `Resources`, calls `AttachToResources` so the active theme is merged immediately.

### Added

- Phase0 **P2a-THEME-SWAP** — gate **47/47**.

### Notes

- ~~OS Light/Dark detection is still deferred~~ → **2.35.0** / **P2a-THEME-OS**. `{ThemeResource}` markup already aliases `{DynamicResource}`.

---

## [2.33.0] — 2026-07-22 — Phase 2a (nested Border.Child ContentPresenter)

### Changed

- **`ContentPresenter`** — implements `IUIElement` / `IControl` with Cairo widget; caption paints in `W_Paint`. Public `Draw()` retained for host paint path.
- **`ApplyButtonTemplate` / `Button.AttachTemplatePresenter`** — nests live `ContentPresenter` under cloned `Border.Child` (WPF visual tree). Host skips `DrawContent` when the presenter is nested. Hits pass through so Button still receives clicks.

### Added

- Phase0 **P6i-NEST** — gate **46/46**.

### Notes

- Flat template bags (`Border` + `ContentPresenter` siblings) are still accepted; apply-time nesting builds the WPF tree. ~~System/OS theme still deferred~~ → **2.35.0** / **P2a-THEME-OS**.

---

## [2.32.0] — 2026-07-22 — Phase 2a (live ContentPresenter TemplateBinding)

### Changed

- **`ApplyButtonTemplate`** — clones template `ContentPresenter` (or marker-only) as a paint-only live content slot on `Button`; Content/align sync from host DPs (TemplateBinding). Still clones `Border` chrome into `Children`.
- **`Button.ContentPresenter`** — returns the live template slot when present; falls back to the host presenter when Style/template is cleared.

### Added

- Phase0 **P6h-CP** — gate **45/45**.

### Notes

- ContentPresenter remains paint-only (no `cWidgetBase` / Border.Child widget). ~~System/OS theme still deferred~~ → **2.35.0** / **P2a-THEME-OS**.

---

## [2.31.0] — 2026-07-22 — Phase 2a (live lookless ControlTemplate chrome)

### Changed

- **`ApplyButtonTemplate`** — clones the first template `Border` into `Button.Children` (never attaches the template-bag instance). Host `CornerRadius` copy and `Exists("BackColor")` guard retained for overlay / align safety.
- **`Button`** — `SuppressContent` ignores live template chrome (string caption still draws). When template chrome is live, host skips `DrawBackground` / `DrawBorder` (template Border owns chrome). `Style = Nothing` clears chrome. Stable widget key `VCF_TPL_<ObjPtr>`.

### Added

- Phase0 **P6g-LIVE** — gate **44/44**.

### Notes

- No public `TemplateVisual` API; assert via `Children`. ContentPresenter remains paint-only (no live widget / TemplateBinding yet).

---

## [2.30.0] — 2026-07-22 — Phase 2a (ControlTemplate ContentPresenter marker)

### Changed

- **`ApplyControlTemplate` / Button** — flat `ControlTemplate` children: `Border` (chrome) + optional **`ContentPresenter` alignment slot** (`SetContentAlignmentMarker`). Never attaches template nodes as live Button children.
- **`StyleManager`** — after chrome, pushes slot H/V via `SetCurrentValue`. `GetValue("BackColor")` on template `Border` is gated with `Exists` (Border has no BackColor DP; unconditional get raised 424 and aborted apply before align push).

### Added

- Phase0 **P6f-TBIND** — gate **43/43**.
- **`modStyleApplyLog`** — `On Error GoTo` + `Erl` logging for the ApplyStyle chain (`%TEMP%\VCF_StyleApply.log`); kept for upcoming Style/template work.

### Notes

- ~~Live lookless visual tree (Border widget under Button) remains deferred~~ → **P6g-LIVE** / **P6h-CP** / **P6i-NEST**.

---

## [2.29.0] — 2026-07-21 — Phase 2a (UniformGrid ItemsHost harden)

### Fixed

- **`UniformGrid.MoveChildren`** — honors `LockRefresh` / reentrancy guard; uses `ControlWidgetKey`; skips per-add layout during ItemsControl silent batch (was O(N²) `Widgets.RemoveAll` → IDE hang).
- **`ItemsControl`** — silent Add for UniformGrid hosts again; **`ArrangeGeneratedItems`** calls `UniformGrid.ArrangeChildren` once after rebuild.

### Changed

- Phase0 **P7c-PANEL** — also asserts UniformGrid + TextBlock/Button item inflate (KeepAlive). Gate remains **42/42**.

---

## [2.28.0] — 2026-07-21 — Phase 2a (ContentControl Content unify)

### Changed

- **`ContentControl.Content`** — now a bindable **`vbVariant` DP** (was children-only `Object` Set). String Content paints via **`ContentPresenter`**; `IUIElement` Content remains a single visual child (suppresses caption).

### Added

- **`ContentControl.ContentPresenter`** / **`SyncContentPresenter`** — same paint-only model as Button.
- Phase0 **P6e-CC** — gate **42/42**.

### Notes

- VB6 does not inherit `Button` from `ContentControl`; shared semantics only. ~~Lookless template trees still deferred~~ → **P6g**–**P6k**.

---

## [2.27.0] — 2026-07-21 — Phase 2a (ContentAlignment on ContentPresenter / Button)

### Added

- **`ContentPresenter.HorizontalContentAlignment`** / **`VerticalContentAlignment`** — paint-path caption alignment (H: `AlignmentConstants`; V: 0 Top / 1 Bottom / 2 Center).
- **`Button`** matching DPs (defaults Center/Center); synced into presenter before draw.
- Phase0 **P6e-ALIGN** — gate **41/41**.

---

## [2.26.0] — 2026-07-21 — Phase 2a (§2.11 ContentPresenter paint-only)

### Added

- **`ContentPresenter`** — paint-only string `Content` path (`ContentCaption` / `WouldDrawCaption` / `Draw`); no Cairo widget tree in this slice.
- **`Button.ContentPresenter`** + **`SyncContentPresenter`** — caption paint delegates to presenter; `SuppressContent` when `Children.Count > 0`.
- Phase0 **P6e-PRES** — gate **40/40**.

### Notes

- ~~ControlTemplate visual-tree expansion (Border + live presenter widgets) remains deferred~~ → **P6g**–**P6j**.
- UniformGrid ItemsHost item inflate hardened in **2.29.0** (**P7c-PANEL**).

---

## [2.24.0] — 2026-07-21 — Phase 2a (B-RESZ / B-NAV / bind fidelity / Margin-Padding families)

### Fixed

- **`modItemTemplateEngine`** — ItemTemplate `TextBlock` clones copy **Bindings** via `CloneTextBlockWithBindings` (not `TextBlock.Clone` / `cProperties.BindTo`, which crashed the IDE under dense clones).
- **`VCF.CloneDataTemplateForItem`** — public wrapper for Phase0 / tests (module not in typelib).
- **`DetachBindingsTree`** — skip `For Each` when `IControl.Children` is **Nothing** (TextBlock/Image); previously Access Violation in vbRichClient5.

### Added

- **`ListView.Margin`** / **`ListView.Padding`** DPs — defaults **Margin=0**, **Padding=4,1,4,1** (Win10 ListBoxItem); default item text draw reads **Padding** (no longer hard-coded).
- **`Button.Padding`** DP — default **1** (Aero2); caption/`MoveChild` insets use **BorderWidth + Padding**.
- **`TextBox.Margin`** / **`TextBox.Padding`** DPs — defaults **Margin=0**, **Padding=1** (Aero2; matches legacy `InnerSpace=1`); `TextBoxBase.SetContentPadding` supports asymmetric LTRB.
- **`UniformGrid.Padding`** — default **2** locked (cell inset; intentional VCF behavior, not changed).
- **7c-dialog** — `modItemTemplateEngine.CloneButtonWithBindings`; ItemsControl can inflate **Button** ItemTemplates with `{Binding}` (Content / Command / CommandParameter). No `@` substitution in framework.
- **`ItemsPanelTemplate`** — `ItemsControl.ItemsPanel`; hosts **StackPanel** (default) or **UniformGrid**; `ItemsHost` is now **Object** (was StackPanel).
- Phase0 **B-RESZ**, **B-NAV**, **B-BIND-DENSE**, **P2a-PAD**, **P2a-PAD-TB**, **P2a-PAD-UG**, **P7c-DLG**, **P7c-PANEL** — gate **39/39**.

### Notes

- Margin/Padding content-control families complete; layout hosts stay Margin=0 / no content Padding.
- **P7c-DLG** / **P7c-PANEL** prove dialog DataTemplate + ItemsPanelTemplate. UniformGrid ItemsHost with generated items is gated (LockRefresh + one arrange per rebuild).
- Full `DeNovo.vbp` remains **Phase 2b**.

---

## [2.23.0] — 2026-07-21 — Phase 1 harness exit (DeNovoSmoke m3)

Tag: **`v2.23.0-wpf-alignment-phase1-complete`**

### Phase 1 sign-off (UI/XAML)

Agreed exit criteria met ([DENOVO_HARNESS_PROPOSAL_RESPONSE.md](./DENOVO_HARNESS_PROPOSAL_RESPONSE.md) §2.4):

| Gate | Status |
|------|--------|
| Phase0 **31/31** | Pass (includes P8-INHERIT) |
| Splash → Login → MainMenu → Sales **layout-only** | Pass in DeNovoSmoke |

**Still out of scope for Phase 1 / Phase 2a:** full `DeNovo.exe` / `DeNovo.vbp` pin (**Phase 2b** — deferred while Data/Kernel finish). VCF continues on Phase 2a with Phase0 + DeNovoSmoke only.

### Harness (m3)

- **SalesOrderView** + **StubSalesOrderViewModel** — Design* two-column shell, stub `Lines` ListView, Add/Clear, **Back to menu**, **Shift+S**, lazy `[P7d-LOAD-SALES]`.
- Builds on **2.22.0** (8b + m2).

### Verification

- [x] DeNovoSmoke — Login → MainMenu → Sales; Back; Add/Clear lines; named LeftColumn/RightColumn
- [x] Phase0 still green on **2.22.0** DLL (no framework API change in m3)

---

## [2.22.0] — 2026-07-20 — Phase 8b (lazy GetValue inheritance) + DeNovoSmoke m2

Tag: **`v2.22.0-wpf-alignment-p8b-denovo-m2`**

### Changed

- **Inheritable DPs** (notably **`DataContext`**) resolve via **`DependencyProperty.GetValue`** parent-walk when no explicit local/`SetCurrentValue` is present (WPF pull model).
- **`PassPropertyValue`** no longer copies values to children (call sites kept; no fan-out).
- **`InheritPropertyValues`** — no `SetCurrentValue` copy; notifies unset inheritable DPs on attach, or **defers inside `InheritanceBatch`** to one **`PropagateInheritableFrom`** on End.
- On inheritable change, **`NotifyInheritableToDescendants`** refreshes **Binding** callbacks on pull descendants (also batched during load/style).
- **`SetInheritanceBatchRoot`** only applies at batch depth 1 — nested `XAMLReader.Load` (`res:` fragments) must not replace the outer load root (fixes DataContext command/bindings on screens that embed StatusBar/LoginPad).

### Behavior notes

- Reading **`Child.DataContext`** after **`Root.DataContext = Vm`** returns **`Vm`** without storing a copy on the child.
- A local **`SetValue`** / style **`SetCurrentValue`** on a child still blocks inheritance for that subtree.
- Object DPs whose unset sentinel is **`Nothing`** cannot distinguish “cleared to null” from “unset” (pre-existing VCF limitation).

### Harness (DeNovoSmoke m2)

- **MainMenuView** + **StubMainMenuViewModel**; Login → MainMenu / Logout; **Shift+M**; lazy `[P7d-LOAD-MAINMENU]`.
- Also includes Phase **7f** (lazy Login, load benches) and **8a** batch plumbing shipped with this pin.

### Verification

- [x] Phase0 — **P8-INHERIT** (`PassDuringLoad=0`) + **P4-DCTX** PASS (31/31)
- [x] DeNovoSmoke — Splash → Login → MainMenu; named lookup; Sales stub
- [ ] Optional: compiled VCF re-measure Splash/Login (expected tighter than IDE)

---

## [2.21.0] — 2026-07-20 — Phase 8a (inheritance batch + DataContext coalesce)

### Changed

- **`PassPropertyValue`** suppressed during **`InheritanceBatch`** (XAML `Load` / `LoadSuperclassData` + `StyleManager.ApplyStyle`); one coalesce propagate from batch root on `EndInheritanceUpdate`.
- **`InheritPropertyValues`** iterates **inheritable DPs only** (tracked at `Register`).
- **`Button`:** `Selected` / `BackColor` / `BorderColor` **`IsInheritable=False`** (no longer pushed to children).

### Added

- **`modInheritanceBatch`**, counters via `VCF.ResetInheritanceCounters` / `PassPropertyValueCalls` / `InheritPropertyValuesCalls`.
- Phase0 **P8-INHERIT**.

### Superseded by 8b

- Full WPF lazy `GetValue` parent-walk — see **[2.22.0]**.

### Verification

- [ ] Rebuild `Demac.VCF.dll`, Phase0 — **P8-INHERIT** + **P4-DCTX** PASS
- [ ] DeNovoSmoke — Login binds; compare `[P7d-LOAD-LOGIN]` vs prior baseline

---

## [2.20.0] — 2026-07-18 — Phase 7f + Layout Option B

### Changed

- **`XAMLReader.LoadSuperclassData`** — wraps tree build in root `LockRefresh` + `BeginRenderUpdate` / `EndRenderUpdate` (single refresh after load). Behavior-compatible; fewer intermediate paints.
- **`VCF_SHUTDOWN_DIAG` removed** — `modShutdownDiag`, `Constructor.ShutdownDiag*` wrappers, and all TraceEnter/Leave/Step instrumentation deleted. Functional shutdown/detach logic unchanged.
- **`Border` layout Option B** — multi-child Borders with `DesignLeft`/`DesignTop` use `LegacyScaleLayout`. **Single-child Borders always decorator-fill** (WPF `Border.Child`), honoring child **Margin** (uniform `Margin="4"` mirrored when only DesignLeft/Top shimmed).
- **`ListView` item text** — default content inset **4,1,4,1** (Win10 WPF `ListBoxItem` Padding; classic was `2,0,0,0`).
- **`ListView.Move`** — uses `W.Move` instead of Parent.Widgets Remove/Add.

### Harness (DeNovoSmoke)

- **Lazy Login** by default (`EAGER_LOGIN_LOAD = False`) — Login XAML loads on first `ShowLogin`.
- **`[P7d-LOAD-*]`** Immediate timings via `ENABLE_LOAD_BENCH` (`modHarnessLoadBench`).

### Tests

- **P7d-LAY-RESIZE** (Phase0) — Border Design* children scale 400×300 → 200×150.

### Verification

- [ ] Rebuild `Demac.VCF.dll`, F5 Phase0 — **P7d-LAY-RESIZE** PASS
- [ ] F5 DeNovoSmoke — `[P7d-LOAD-SPLASH]` at startup; `[P7d-LOAD-LOGIN]` only after Shift+L / splash timer
- [ ] `USE_WPF_LOGIN_LAYOUT = False` → Shift+2/3 — legacy Login inner content scales
- [ ] `USE_WPF_LOGIN_LAYOUT = True` → Shift+3 — list ~204px, star column grows

---

## [2.19.0] — 2026-06-27 — Phase 7d (DeNovoSmoke harness + host navigation)

Tag: **`v2.19.0-wpf-alignment-p7d-denovo-smoke`**

### Added

- **`.Tests/DeNovoSmoke/`** — minimal POS UI harness (VCF + vbRichClient5): Splash → Login, stub VMs, fixture sync from denovo.
- **`Window.RelayoutChildren`**, **`Window.RebuildNamedItemsList`**, **`Window.ApplyDeferredChildLayout`** — documented host-app APIs for visibility-based view navigation.
- **`UserControl.ApplyDeferredHostLayout`** (Friend) — invoked from `Window.RelayoutChildren`.
- **Docs:** [DENOVO_HARNESS_PROPOSAL_RESPONSE.md](./DENOVO_HARNESS_PROPOSAL_RESPONSE.md), [POS_LAYOUT_RESIZE.md](./POS_LAYOUT_RESIZE.md).

### Changed

- **`FrameworkElement` / `modLayoutEngine`** — legacy arrange reads `Margin`/`Design*` consistently when `LegacyScaleLayout` is on (migrated POS XAML).
- **`UserControl.Move`** — propagates `OnHostResize` to children after parent arrange.
- **`Border.Move`** — calls **`ArrangeBorderChild`** after resize so a single child (e.g. inner **Grid**) fills the border client; avoids 300×300 default grid clip.
- **`Grid.Move`** — propagates **`ArrangePanel`** after widget resize (nested grid reflow).

### Verification

- [ ] Phase0 **31/31**
- [ ] DeNovoSmoke milestone 1 (Splash → Login, pinned **2.18.0+** DLL)

---

## [2.18.0] — 2026-06-19 — POS layout shim (DeNovo integration — validate in IDE)

Tag: **`v2.18.0-wpf-alignment-p7c-layout-shim`** · Phase0 **31/31** (includes **P7c-LAY**).

**Source:** [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md) — DeNovo pin `2.15.0` + mechanical XAML migration.

### Added

- **`ApplyLegacyLayoutProperty`** (`modLayoutEngine.bas`) — when XAML `SetProperty` cannot assign `Margin` / `Width` / `Height`, map to `IUIElement.DesignLeft` / `DesignTop` / `DesignWidth` / `DesignHeight`.
- **`XAMLReader.SetProperty`** — invokes shim for `IUIElement` after `CallByName` failure.
- **P7c-LAY** — loads migrated POS XAML (`Margin` on `TextBlock` → `Design*`).
- **`Invoke-VcfXamlMigration.ps1`** — skip layout transforms on legacy types (`TextBlock`, `Image`, `Scene`, `UniformGrid`, `TextBox`, `WindowsFormsHost`); report `Margin` on legacy tags.

### Migration note (not breaking API)

- **`Design*` → `Margin`** on legacy types is unsafe on **`v2.15.0`** alone; use **this tag** or keep `Design*` on those elements.
- Set **`ActiveThemeName`** non-empty in `MyApp.xml` when using `{DynamicResource}`.

### Verification

- [ ] Phase0 **31/31** (validate in IDE)
- [ ] DeNovo POS smoke §3 on pinned build

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

---

---

---

---

## [2.17.0] — 2026-06-21 — Phase 7b (XAML transform tooling — validated)

Tag: **`v2.17.0-wpf-alignment-p7b`** · **No DLL API changes** · Phase0 still **30/30** (no new bench test).

### Added

- **`tools/xaml-migrate/Invoke-VcfXamlMigration.ps1`** — mechanical transforms: `Design*` → layout DPs, `UnboundListView` → `ListView`, `{ThemeResource}` → `{DynamicResource}`, `Button Text` → `Content`; `-WhatIf`, `-ReportOnly`, `-SelfTest`.
- **[XAML_MIGRATION_PROMPTS.md](./XAML_MIGRATION_PROMPTS.md)** — Cursor prompts for scan, review, Button Content, Scene BackColor, `res:` fragments, MyApp styles.

---

## [2.16.0] — 2026-06-21 — Phase 7a (POS migration package — validated)

Tag: **`v2.16.0-wpf-alignment-p7a`** · Phase0 **30/30** pass. **No DLL API changes** — same `Demac.VCF.dll` as 2.15.0.

### Added

- **[POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md)** — DeNovo manual smoke checklist (login → sales → grid → dialog).
- **[MIGRATION.md](./MIGRATION.md)** — **Upgrading to 2.15.0** master section (Phases 0–6 pin guide).
- **P7a-SMOKE** — loads POS-shaped `PosSalesOrderShell.xml` (Scene + UniformGrid + legacy `Design*`). Real POS `SalesOrder.xml` still uses `Scene BackColor=` (legacy widget setter); omitted here because strict Phase0 rejects unknown Scene DPs.

### Test

- Suite → **30** tests.

---

## [2.15.0] — 2026-06-21 — Phase 6d (Render coalescing — validated)

Tag: **`v2.15.0-wpf-alignment-p6d`** · Phase0 **29/29** pass.

### Added

- **`RenderCoalescer`** class — host-facing batch refresh API (`BeginRenderUpdate`, `EndRenderUpdate`, `RequestWidgetRefresh`, `PendingCount`, `LastFlushCount`). Use **`New RenderCoalescer`** from external projects.
- **`modRenderCoalescer`** — internal batch widget refresh engine; **`RequestWidgetRefresh`** dedupes by widget pointer.

### Changed

- **`StyleManager.ApplyStyle`** — defers widget refresh via **`RequestWidgetRefresh`** (respects any outer **`BeginRenderUpdate`** batch); no separate inner batch wrapper.
- **`Button`** — **`Selected`** / **`Content`** DP changes use **`RequestWidgetRefresh`** instead of immediate **`W.Refresh`**.

### Test

- **P6d-COAL** in `.Tests/Phase0` (suite → **29** tests).

---

## [2.14.0] — 2026-06-21 — Phase 6c (ControlTemplate — validated)

Tag: **`v2.14.0-wpf-alignment-p6c`** · Phase0 **28/28** pass.

### Added

- **`ControlTemplate`** — `TargetType` + visual tree (`Children`); **`Clone`** for resource merge.
- **`Style.Template`** — assigned from code or XAML `<Setter Property="Template">` / nested **`ControlTemplate`**.
- **`modControlTemplateEngine`** — applies template root to control chrome after style setters/triggers (Phase 6c: **Button** + **Border** root → **`CornerRadius`**, **`BackColor`**).
- **XAML** — standalone **`ControlTemplate`** load via **`XAMLReader.LoadElement`**.

### Changed

- **`StyleManager.ApplyStyle`** — invokes control template application after triggers.

### Test

- **P6c-TMPL** in `.Tests/Phase0` (suite → **28** tests).

---

## [2.13.0] — 2026-06-21 — Phase 6b (Style.Triggers — validated)

Tag: **`v2.13.0-wpf-alignment-p6b`** · Phase0 **27/27** pass.

### Added

- **`PropertyTrigger`** — WPF-like **`Trigger Property="..." Value="..."`** with child setters.
- **`Style.Triggers`** collection — **`AddTrigger`**, **`TriggerCount`**, **`TriggerAt`**, **`ClearTriggers`**.
- **`modStyleTriggerEngine`** — evaluates triggers after base style setters; **`BasedOn`** chain triggers first.
- **`Button.IsMouseOver`** — readable/writable state property for trigger evaluation (mouse handlers update it).
- **XAML** — **`Style.Triggers` / `Trigger`** parsing in **`XAMLStyleReader`**; **`Setter Name=`** alias fixed alongside **`Property=`**.

### Changed

- **`StyleManager.ApplyStyle`** — applies active trigger setters after style setters.

### Test

- **P6b-TRIG** in `.Tests/Phase0` (suite → **27** tests).

---

## [2.12.0] — 2026-06-21 — Phase 6a (Button Content — validated)

Tag: **`v2.12.0-wpf-alignment-p6a`** · Phase0 **26/26** pass.

### Added

- **`Button.Content`** dependency property — string caption drawn in **`W_Paint`** (centered) using widget font style setters.
- **XAML `Content="..."`** and **`Content="{Binding ...}"`** on `Button`.
- **`Text` → `Content` alias** when `Content` DP exists (strict XAML accepts legacy `Text="..."` on Button).

### Changed

- Golden / layout Phase0 XAML samples use **`Content="OK"`** instead of widget-fallback **`Text`**.

### Precedence

- If the button has **visual children** (e.g. nested `TextBlock`), **children win** — string `Content` is not painted.
- If there are **no children** and `Content` is a non-empty string, caption is painted on the button chrome.

### Test

- **P6a-CONTENT** in `.Tests/Phase0` (suite → **26** tests).

---

## [2.11.0] — 2026-06-21 — Phase 5c (Row hierarchy — validated)

Tag: **`v2.11.0-wpf-alignment-p5c`** · Phase0 **25/25** pass.

### Added

- **`ListViewBase.QueryRowLevel`** event — host sets **`Level`** (0 = parent, 1+ = child) per flat row index.
- **`ListViewBase.RowIndentPerLevel`** (default **16**) · **`MeasuredRowLevel`** · **`MeasuredRowIndent`**.
- **`ListView.QueryRowLevel`** — bubbled from engine.

### Changed

- Row paint and **`MeasureRow`** width account for hierarchy indent (single-column / owner-draw path).

### Test

- **P5c-HIER** in `.Tests/Phase0` (suite → **25** tests).

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
