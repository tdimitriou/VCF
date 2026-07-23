# VCF — dependency property registry (current + target)

**Companion:** [VCF_CLASS_REFERENCE.md](./VCF_CLASS_REFERENCE.md) · [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md)  
**Last updated:** 2026-07-23  

---

## 1. How properties work today

### 1.1 Registration pattern

Each control in `Class_Initialize`:

```vb
With NewDependencyProperties(Me)
    Call .Register("DataContext", vbObject, , , , NewDependencyPropertyMetadata(False, False, True))
    ' ... more Register calls
End With
```

**Metadata flags:**

| Flag | Meaning |
|------|---------|
| `IsInheritable` | Effective value may pull from parent via `DependencyProperty.GetValue` (Phase 8b); ancestor changes notify descendant bindings |
| `AffectsRender` | Control should `W.Refresh` on change |
| `AffectsMeasure` | **Defined but unused** — no InvalidateMeasure pipeline |
| `BindingMode` | Default for `{Binding}` when Mode=Default |
| `DefaultValue` / `HasDefaultValue` | Metadata default below inherit (**3.2.0**); prefer over init `SetCurrentValue` for framework defaults |

### 1.2 Non-DP properties (dual storage problem)

These are set via XAML `CallByName` / fields — **outside** binding/style precedence:

| Property | Controls | Storage |
|----------|----------|---------|
| `DesignLeft`, `DesignTop`, `DesignWidth`, `DesignHeight` | All `IUIElement` | Private fields |
| `CornerRadius` | Border, Button | Field + DP (Border) / field only (Button) |
| `GradientBackground` | Button | Field |
| `ClickMode`, `BorderWidth` | Button | Fields |
| `ImageKey`, `KeepAspectRatio` | Image | Fields |
| `Padding`, `Rows`, `Columns` | UniformGrid | Fields (+ Padding DP) |
| `Background`, row colors | ListView | Fields |
| `ItemTemplate` | ListView | Field (IItemsControl stub) |

**Rewrite rule:** Migrate to DPs on `FrameworkElement` or remove in favor of templates.

### 1.3 Attached properties (RegisterAttached — Grid / DockPanel)

**3.4.0:** `DependencyPropertyRegistry.RegisterAttached` metadata + typed **`Grid.Get*` / `Set*`** for:

- `Grid.Row` (default 0)
- `Grid.Column` (default 0)
- `Grid.RowSpan` (default 1)
- `Grid.ColumnSpan` (default 1)

**3.17.0:** **`DockPanel.Dock`** (default `Left` / 0) + **`DockPanel.GetDock` / `SetDock`**.

**Storage:** per-element `DependencyProperties` (`Grid.Row`, `DockPanel.Dock`, …) via lazy **`EnsureAttachedProperty`** on Set/XAML (**3.16.0**); nested `AttachedProperties("Grid"|"DockPanel")` dict kept as shim for writers/legacy.

**Done (3.16.0):** migrate `Grid.*` into the target DP bag; layout/`Get*` read DP (ClearValue → metadata default); dict shim retained. Do **not** eager-register attached DPs on every instance.

**UniformGrid** reads spans via `GetGridAttachedLong`.

---

## 2. Current registry by type

### 2.1 Shared across most controls

| Property | Type | Inherit | AffectsRender | Default binding | Types |
|----------|------|---------|---------------|-----------------|-------|
| `DataContext` | Object | Yes | No | OneWay | Button, Border, Panel, Window, UC, Scene, TextBlock, TextBox, ListView, UniformGrid, Image, WindowsFormsHost, UnboundListView |
| `Style` | Style | No | Yes | OneWay | All styled controls above |

### 2.2 Button

| Property | Type | Inherit | AffectsRender | Notes |
|----------|------|---------|---------------|-------|
| `Visible` | Boolean | No | — | Not WPF Visibility enum |
| `Margin` | Thickness | No | — | |
| `Command` | ICommand | No | — | |
| `CommandParameter` | Variant | No | — | |
| `Selected` | Boolean | Yes | Yes | |
| `BackColor` | Long | Yes | Yes | Visual inheritance |
| `BorderColor` | Long | Yes | Yes | |
| `ToolTip` | String | No | Yes | |

**Non-DP:** `ClickMode`, `BorderWidth`, `GradientBackground`, `CornerRadius`, all `Design*`

### 2.3 TextBlock

| Property | Type | AffectsRender |
|----------|------|---------------|
| `ForeColor` | Long | Yes |
| `FontName` | String | Yes |
| `FontSize` | Single | Yes |
| `FontBold`, `FontItalic`, `FontUnderline`, `FontStrikeThrough` | Boolean | Yes |
| `ScaleFont` | Boolean | Yes |
| `HorizontalAlignment`, `VerticalAlignment` | Long | Yes |
| `Text` | String | Yes |

### 2.4 TextBox

| Property | Type | AffectsRender | BindingMode |
|----------|------|---------------|---------------|
| `Text` | String | Yes | **TwoWay** |
| `Alignment` | Long | Yes | OneWay |
| `VCenter` | Boolean | Yes | OneWay |
| `PasswordChar` | String | Yes | OneWay |
| `CueBanner` | String | Yes | OneWay |

### 2.5 Border

| Property | Type | Notes |
|----------|------|-------|
| `ShowGridLines` | Boolean | Legacy debug? |
| `CornerRadius` | CornerRadius UDT | |

### 2.6 Panel, Window, Scene, WindowsFormsHost, UniformGrid

| Property | Types | Notes |
|----------|-------|-------|
| `ShowGridLines` | Panel, Border, Window, WFH, UniformGrid | |
| `BorderStyle` | Window | |
| `Visible` | UniformGrid | Boolean, not Visibility |
| `Padding` | UniformGrid | Thickness object |

### 2.7 ListView

| Property | Type | Notes |
|----------|------|-------|
| `ItemsSource` | Object | Must be ObservableCollection (bug B4) |

**Missing DPs (target Selector):** `SelectedItem`, `SelectedIndex`, `SelectedValue`, `SelectedValuePath`, `ItemTemplate` (proper), `ItemsPanel`

### 2.8 Image

| Property | Type |
|----------|------|
| `DataContext` | Object |

**Non-DP:** `ImageKey`, `KeepAspectRatio`

---

## 3. Target registry — FrameworkElement (all visual types)

Register **once** in `DependencyPropertyRegistry` module:

### 3.1 Layout (replaces Design*)

| Property | Type | AffectsMeasure | AffectsRender | Default |
|----------|------|----------------|---------------|---------|
| `Width` | Single | Yes | No | NaN/auto |
| `Height` | Single | Yes | No | NaN/auto |
| `MinWidth`, `MinHeight` | Double | Yes | No | default 0 |
| `MaxWidth`, `MaxHeight` | Double | Yes | No | default **0 = unbounded**; positive = ceiling |
| `Margin` | Thickness | Yes | No | 0 |
| `Padding` | Thickness | Yes | No | 0 |
| `HorizontalAlignment` | Enum | Yes | No | Stretch |
| `VerticalAlignment` | Enum | Yes | No | Stretch |
| `ActualWidth`, `ActualHeight` | Single | No | No | Read-only, set by Arrange |

### 3.2 Tree & context

| Property | Type | Inherit | Notes |
|----------|------|---------|-------|
| `DataContext` | Object | Yes | Rebind all BindingExpressions on change |
| `Name` | String | No | x:Name registry |
| `Visibility` | Visibility enum | No | Visible/Hidden/Collapsed — affects measure when Collapsed |
| `IsEnabled` | Boolean | No | Phase 6+ |

### 3.3 Resources & styling

| Property | Type | Notes |
|----------|------|-------|
| `Style` | Style | SetCurrentValue from StyleManager |
| `Resources` | ResourceDictionary | Local merged dict |

### 3.4 Attached (Grid)

| Attached | Type | Used on |
|----------|------|---------|
| `Grid.Row`, `Grid.Column` | Integer | Any FrameworkElement in Grid |
| `Grid.RowSpan`, `Grid.ColumnSpan` | Integer | Any FrameworkElement |
| `Grid.RowDefinitions`, `Grid.ColumnDefinitions` | Collection | Grid only |

---

## 4. Target registry — by control type (additions)

### Button (Control)

| Property | Type | Notes |
|----------|------|-------|
| `Content` | Object/String | Phase 5 — replaces nested TextBlock for caption |
| `Command`, `CommandParameter` | | Keep |
| `IsPressed` | Boolean | Internal |
| `Background`, `BorderBrush`, `BorderThickness` | | Template phase |
| Remove | `Visible` bool | → Visibility |
| Remove | `Selected` | → use IsPressed or visual state |

### TextBlock

Keep font/text DPs; add `TextWrapping`, `TextTrimming` (P2).

### TextBox

Keep TwoWay `Text`; add `UpdateSourceDelay` metadata (§2.9 alignment).

### Border (Decorator)

| Property | Type |
|----------|------|
| `Child` | FrameworkElement (single) |
| `Background`, `BorderBrush`, `BorderThickness`, `CornerRadius` | |

Remove `ShowGridLines`.

### ListView (Selector)

| Property | Type |
|----------|------|
| `ItemsSource` | IEnumerable |
| `ItemTemplate` | DataTemplate |
| `ItemsPanel` | ItemsPanelTemplate |
| `SelectedItem`, `SelectedIndex` | |
| `SelectedValue`, `SelectedValuePath` | |
| `DisplayMemberPath` | String (optional P2) |

### ContentControl

| Property | Type |
|----------|------|
| `Content` | Object |
| `ContentTemplate` | DataTemplate |

### ItemsControl

| Property | Type |
|----------|------|
| `ItemsSource`, `ItemTemplate`, `ItemsPanel` | |
| `ItemContainerStyle` | Style (P2) |

---

## 5. Value precedence (locked — 3.2.0)

```text
Local value (SetValue)           ← XAML, CLR, Binding, TemplateBinding
  ↓ if unset
Current (SetCurrentValue)        ← style setters, triggers, some init seeds
  ↓ if unset
Inherited (IsInheritable)        ← lazy parent walk
  ↓ if unset
Metadata DefaultValue            ← DependencyPropertyMetadata
```

**API:** `ClearValue` drops local only; `ReadLocalValue` inspects local slot; binding Detach clears local.

**Trigger interim:** re-`ApplyStyle` on condition change (active triggers only).  

**Roadmap (deferred):** full WPF-aligned layer stack — [VCF_DP_PRECEDENCE_ROADMAP.md](./VCF_DP_PRECEDENCE_ROADMAP.md). Do not start until queued.

---

## 6. Metadata callbacks (target)

Each registered property may specify:

```text
PropertyChangedCallback(element, oldValue, newValue)
  → sync Cairo widget
  → InvalidateMeasure if AffectsMeasure
  → InvalidateVisual if AffectsRender
  → notify BindingExpression listeners

CoerceValueCallback (P3)
ValidateValueCallback (P3)
DefaultValue (replace Class_Initialize SetCurrentValue blocks)
```

---

## 7. Migration: Design* → layout DPs

| VCF today (XAML) | Target | Notes |
|------------------|--------|-------|
| `DesignLeft="10"` | `Margin` or Canvas.Left | Prefer Margin in panels |
| `DesignTop="20"` | same | |
| `DesignWidth="100"` | `Width="100"` | |
| `DesignHeight="40"` | `Height="40"` | |
| Numeric alignment `HorizontalAlignment="1"` | Named enum | Document mapping |

**One-release shim (optional):** XAMLReader maps Design* → DPs with deprecation warning — not permanent dual API.

---

## 8. POS XAML audit hotspots

Properties most used in `pos-v1/UI/Resources/XAML/` (grep recommended during migration):

- `DesignLeft/Top/Width/Height` — **all sales screens**
- `Grid.ColumnSpan` / `Grid.RowSpan` on UniformGrid children
- `{Binding Path=… Mode=TwoWay}` on TextBox
- `{StaticResource …}` / `{ThemeResource …}` on Button styles
- `res:Screens\…` includes — → MergedDictionaries

---

*Registry implementation owner: Phase 1. Update this doc when `VCF_PROPERTY_REGISTRY.md` moves to VCF repo `doc/`.*
