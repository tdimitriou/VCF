# VCF — ListView architecture specification

**Companion:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) · Alignment §2.14, §6.1  
**Last updated:** 2026-06-19  

---

## 1. Purpose

Complete specification of the **ListView stack** — current three-layer model, known defects, WPF **Selector** target, and **InvoiceGrid** (POS) requirements — for VCF team rewrite (Phases 4–5).

---

## 2. Layer model (today)

```text
┌─────────────────────────────────────────────────────────┐
│  ListView (755 LOC)          UnboundListView (448 LOC)   │
│  IItemsControl · bound rows  Owner-draw · no ItemsSource │
├─────────────────────────────────────────────────────────┤
│  ListViewBase (1178 LOC) — Cairo/vbWidgets list engine   │
│  cWidgetBase · scrollbars · columns · selection · paint  │
├─────────────────────────────────────────────────────────┤
│  vbRichClient5 cWidgetBase / Cairo context               │
└─────────────────────────────────────────────────────────┘
```

**Not** the vbWidgets DLL standalone List control — this is an **owner-draw Cairo engine** (cwWidget / Colin Edwards lineage) vendored into VCF as `ListViewBase`.

---

## 3. ListViewBase — engine reference

### 3.1 Creatable & exposure

- **`Creatable = False`** — only constructed by ListView/UnboundListView
- **`VB_Exposed = True`** — accessible as `ListView.Base` for advanced POS code

### 3.2 Widget configuration

```vb
Set W = Cairo.WidgetBase
W.RuntimePropertiesCommaSeparated = "VisibleRows,ListCount,ScrollIndex,..."
Public WithEvents HScrollBar, VScrollBar As VCF.ScrollBar
Public WithEvents W As cWidgetBase
```

### 3.3 Events (complete)

| Event | When |
|-------|------|
| `OwnerDrawRowSelector` | Row gutter paint |
| `OwnerDrawHeader` | Column header paint |
| `OwnerDrawItem` | **Per-row** owner-draw — bound ListView hooks here |
| `Click`, `DblClick`, `HeaderClick(ColIdx, SortState)` | Input |
| `DeleteKeyPressed` | Keyboard |
| `ScrollIndexChange`, `HScrollChange` | Scroll |
| `MouseUpClick`, `MouseMoveOnListItem` | Mouse |
| `SelectedAll` | Ctrl+A |
| `DimensionsAdjusted` | Layout |
| `ListIndexChanged` | Selection index |

### 3.4 Selection model (today)

| API | Behavior |
|-----|----------|
| `ListIndex` | Single focus row (-1 none) |
| `MultiSelect` | None / Simple / Extended (VB constants) |
| `Selected(Index)` | Per-row bool |
| `GetSelections()` | Array of selected indices |
| `ClearSelections` | Reset |
| `ShowSelection`, `DeselectOutsideClick` | UX flags |

**Not WPF:** No `SelectedItem` object — index only. POS uses `ListIndex` + ViewModel hacks.

### 3.5 Column model

| API | Role |
|-----|------|
| `ColumnCount`, `ColumnIndex` | Grid columns |
| `ColumnWidth(Idx)`, `ColumnDefaultWidth`, `ColWidth` | Sizing |
| `ColMapIndex(Idx)` | Display order map |
| `MoveColumnToNewIndex` | Reorder |
| `VisibleCols`, `AllowColResize` | Header drag |
| `HeaderClick` + `ColumnSortState` enum | Sort indicator (Asc/Desc/None) |
| `ResetSortStates` | Clear sort |

### 3.6 Scroll & layout

| API | Role |
|-----|------|
| `ListCount` | Row count (bound: ItemsSource.Count) |
| `ScrollIndex` | First visible row |
| `VisibleRows` | Computed visible row count |
| `RowHeight` | **Fixed** row height today |
| `EnsureVisibleSelection` | Scroll to selected |
| `AdjustDimensions` | Recalc on resize |
| `DrawWidth`, `DrawHeight` | Client area |

### 3.7 Paint pipeline (simplified)

```text
W_Paint
  → header + column dividers
  → for each visible row index:
       if OwnerDrawItem event has handlers → fire (ListView draws template here)
       else default row fill + text
  → scrollbars (HScrollBar/VScrollBar widgets)
```

### 3.8 Rewrite targets for engine

| Feature | Today | Target |
|---------|-------|--------|
| Row height | Fixed `RowHeight` | **`MeasureRow(index)`** — variable height |
| Hierarchy | Flat list | **Parent/child rows** (indent, collapse) — InvoiceGrid |
| Virtualization | Scroll only | **Recycle** templates (P2) — optional |
| Sort | Header state | Wire to ICollectionView Sort (P2) |

---

## 4. ListView — bound mode

### 4.1 Construction

- Creates **`ListViewBase`** as `m_Base`
- **`ItemTemplates`** — `List` cache of cloned templates per row/type
- DPs: `DataContext`, `ItemsSource`, `Style`
- **`CanGetFocus = False`** on inner widget — focus policy TBD

### 4.2 ItemsSource contract (today)

```vb
' Must be ObservableCollection — else binding silently fails (B4)
Set ItemsSource = someObservableCollection
```

**Collection change handler:**

- Add/Remove/Replace — refresh row count, invalidate
- **Move — NOT handled (B2)**
- Reset — full refresh

**Uses:** `CollectionViewSource.GetDefaultView` / `ListCollectionView` in some paths — **static init bug (B1)** breaks second view.

### 4.3 ItemTemplate

**IItemsControl stubs (B3):**

```vb
Private Property Set IItemsControl_ItemTemplate(ByVal RHS As DataTemplate)
    ' empty
End Property
```

**Actual template:** Set via public property or XAML — wired partially outside interface.

**Render flow:**

```text
ItemsSource[i] → row DataContext
  → FindDataTemplate / CreateDataTemplate (clone DataTemplate.Children)
  → For each template child (TextBlock, etc.):
       set DataContext to row item
       apply bindings
  → OwnerDrawItem: call DrawOn / widget paint at row rect
```

**Cache:** `ItemTemplates` list — clone per row or reuse — memory hotspot on large lists.

### 4.4 PropertyChangedCallback

`ListViewPropertyChangedHandler` forwards INPC from row items → `PropertyChangedCallback` → refresh row.

**Target:** BindingExpression per template binding with row DataContext.

### 4.5 Public API (beyond scaffold)

| Member | Notes |
|--------|-------|
| `Base` | ListViewBase engine |
| `Resources` | Local resource lookup for templates |
| `Background`, `SelectedBackground`, `SelectedForeground` | Fields |
| `RowHeight` | Delegates to Base |
| `ItemsSource`, `ItemTemplate` | |
| `Name` | |

### 4.6 Known bugs (mandatory fix)

| ID | Issue |
|----|-------|
| B1 | ListCollectionView static init |
| B2 | Move not handled |
| B3 | IItemsControl ItemTemplate stubs |
| B4 | Non-ObservableCollection ItemsSource |
| B6 | DataContext change doesn't rebind template bindings |

---

## 5. UnboundListView — owner-draw mode

### 5.1 Role

- Same visual chrome as ListView
- **No ItemsSource** — POS code handles `OwnerDrawItem` on `Base`
- **`Refresh()`** — force repaint
- Used in stubs (`OrderItemsView.xml` → `<UnboundListView/>`)

### 5.2 Events forwarded

Forwards 14+ events from `Base` (Click, OwnerDraw*, Scroll*, etc.) to outer control.

### 5.3 Rewrite

**Remove type** — merge into single `ListView`:

```text
If ItemsSource Is Nothing → owner-draw mode (today's UnboundListView)
Else → bound mode
```

Or explicit **`IsItemsHostEnabled`** DP — prefer automatic from ItemsSource.

---

## 6. DataTemplate (ListView context)

**Current class (28 LOC):**

- Fields: `DataType`, `Key`, `Name`
- `Children` → `UIElementCollection`

**XAML load:** Template tree stored as child elements (TextBlock, etc.)

**Target:**

- Full inflation API: `LoadContent() As FrameworkElement`
- **ResourceReference** clone for grid cells (§2.16)
- ItemsControl item generation

---

## 7. WPF Selector target

### 7.1 Inheritance (target)

```text
FrameworkElement
  → ItemsControl (ItemsSource, ItemsPanel, ItemTemplate)
    → Selector (+ selection DPs)
      → ListView (ListViewBase engine as ItemsHost internal)
```

### 7.2 Selection DPs (Phase 4–5)

| DP | Maps from today |
|----|-----------------|
| `SelectedItem` | `ListIndex` + ItemsSource[index] |
| `SelectedIndex` | `ListIndex` |
| `SelectedValue` | Column binding path |
| `SelectedValuePath` | Property name on item |
| `SelectedItems` | MultiSelect (P2) |

**XAML:**

```xml
<ListView ItemsSource="{Binding Orders}"
          SelectedItem="{Binding SelectedOrder, Mode=TwoWay}"
          ItemTemplate="{StaticResource OrderLineTemplate}"/>
```

Replaces POS `ListIndex` Let + `@Selected` dialog hacks (§2.13).

### 7.3 Item container

WPF `ListViewItem` equivalent — **optional Phase 5**; until then template root is item visual.

---

## 8. InvoiceGrid (POS) — requirements

**Not a VCF control today** — Codejock ReportControl on `FormMain`.

### 8.1 Required behavior

| Feature | Detail |
|---------|--------|
| Read-only | No in-cell edit |
| Multi-column | ListViewBase column model |
| **Variable row height** | Parent row ~40px, child ~20px — **MeasureRow** |
| **Hierarchy** | Parent/child order lines |
| Selection | `SelectedItem` / row focus |
| Performance | Large order line counts |

### 8.2 Codejock reference

POS `InvoiceGridHelper.cls` — `MeasureRow` logic documents height rules.

### 8.3 VCF target

**Evolve ListViewBase** — not separate VirtualizingListView:

1. `MeasureRow(index As Long) As Long`
2. Indent / tree expanders (P2)
3. Column templates via DataTemplate (multiple TextBlocks per column)
4. Stub `OrderItemsView.xml` migrates from `<UnboundListView/>` to bound `<ListView ItemsSource="{Binding Lines}"/>`

---

## 9. POS usage map

| Artifact | Usage |
|----------|-------|
| `OrderItemsView.xml` | `<UnboundListView/>` stub |
| `MenuItemsGridButton.xml` | Dense grid — 6 bindings/cell (framework + ItemsControl target) |
| Various sales views | Bound ListView / manual lists |
| ViewModels | `ObservableCollection` of row VMs |

---

## 10. ScrollBar integration

`ListViewBase` owns **`HScrollBar`** and **`VScrollBar`** VCF controls as child widgets.

**Rewrite:** ScrollBar stays; link to **`ScrollViewer`** pattern (P2) for nested content — optional.

---

## 11. Styles

**Built-in:** `Styles/ListView.xml`, `Styles/UnboundListView.xml`

**Apply path:** `GetBaseStyle` → `StyleManager.ApplyStyle` on DPs (Background, etc.)

---

## 12. Test plan

| Test | Validates |
|------|-----------|
| Bind 1000 rows ObservableCollection | Perf + memory |
| Move item in collection | B2 |
| Two ListViews two sources | B1 |
| Swap DataContext on ListView | B6 rebind |
| ItemTemplate with 3 TextBlocks + bindings | Template clone |
| Owner-draw mode ItemsSource=Nothing | Unbound merge |
| Variable MeasureRow golden | InvoiceGrid prep |
| SelectedItem TwoWay binding | Selector |

**Golden files:** `.Tests/Test0/DataTemplate1.xml`, POS order line templates (when added).

---

## 13. Phase breakdown

| Phase | Deliverable |
|-------|-------------|
| **4a** | Fix B1, B2, B4; BindingExpression in templates |
| **4b** | ItemsControl + basic Selector DPs |
| **5a** | Merge UnboundListView; IItemsControl complete |
| **5b** | ListViewBase MeasureRow + variable height |
| **5c** | Hierarchy + InvoiceGrid parity tests |
| **6** | Virtualization (if needed after 5c metrics) |

---

## 14. Anti-patterns to remove

| Pattern | Replacement |
|---------|-------------|
| `ListView.Base.ListIndex = i` in VM | `SelectedIndex` / `SelectedItem` binding |
| Manual `Refresh` after collection hack | INCC Move + proper notifications |
| Per-cell `ContentControl` | ResourceReference / ItemsControl |
| Second ListView type | Single ListView |
| Template cache leak on navigate | Detach template bindings |

---

*Engine owner: assign senior control developer familiar with Cairo/vbWidgets. Cross-ref [VCF_CLASS_REFERENCE.md § ListViewBase](./VCF_CLASS_REFERENCE.md).*
