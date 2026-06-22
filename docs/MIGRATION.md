# Demac.VCF — migration guide

**Audience:** POS (DeNovo) and internal Demac apps upgrading `Demac.VCF.dll`.

**Handoff baseline:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md)

---

## General upgrade process

1. Read **[BREAKING_CHANGES.md](./BREAKING_CHANGES.md)** for your target version.
2. Build and register **`Demac.VCF.dll`** (`regsvr32 bin\Demac.VCF.dll`).
3. **Recompile** the app EXE (required — VCF uses **Project Compatibility**; typelib GUID is stable but member layout changes until Binary Compatibility is enabled for production pins).
4. Run **`.Tests/Phase0`** and app smoke tests.
5. Apply XAML transforms (per phase below).
6. Apply VB6 code transforms (per phase below).

See [VCF_TEAM_HANDOFF_GUIDE.md §8.1](./VCF_TEAM_HANDOFF_GUIDE.md) for No / Project / Binary compatibility during the rewrite.

**POS smoke checklist:** [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md)  
**Bulk XAML:** [tools/xaml-migrate/README.md](../tools/xaml-migrate/README.md) · [XAML_MIGRATION_PROMPTS.md](./XAML_MIGRATION_PROMPTS.md)  
**DeNovo consumer policy:** denovo monorepo → `docs/migration/DENOVO_VCF_MIGRATION_POLICY.md`  
**POS runtime feedback (2026-06-19):** [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md) — layout shim **`v2.18.0-wpf-alignment-p7c-layout-shim`**

---

## Upgrading to 2.18.0 (POS layout shim — migrated XAML)

**When:** DeNovo applied mechanical XAML migration (`Margin` on `TextBlock`, etc.) and needs runtime layout on pin ≥ 2.15.0.

**Tag:** **`v2.18.0-wpf-alignment-p7c-layout-shim`**

| Item | Action |
|------|--------|
| DLL | Rebuild/register VCF; copy to DeNovo lib path |
| XAML script | Re-run `Invoke-VcfXamlMigration.ps1` (legacy types skip `Margin` transform) or keep migrated XAML + use shim DLL |
| Themes | `ActiveThemeName` non-empty when using `{DynamicResource}` in `MyApp.xml` |
| Verify | Phase0 **31/31**; DeNovo smoke §3 |

---

## Upgrading to 2.15.0 (Phases 0–6 complete — framework baseline)

### DLL pin

**Tag:** **`v2.15.0-wpf-alignment-p6d`** (framework) · **`v2.16.0-wpf-alignment-p7a`** adds migration docs + POS shell smoke test (same DLL as 2.15.0).

### Summary of changes since 2.0.0

| Phase | Tag (approx) | Required POS change | Optional / incremental |
|-------|--------------|---------------------|----------------------|
| 0 | 2.0.0 | None | `StrictXamlLoad`, TypeRegistry |
| 1 | 2.1–2.2 | None | `Width`/`Height` instead of `Design*` |
| 2 | 2.3 | None | StackPanel, Grid for new screens |
| 3 | 2.4 | None | ResourceDictionary, `{DynamicResource}` |
| 4 | 2.5–2.8 | None | Binding rebind automatic; ItemsControl/Selector |
| 5 | 2.9–2.11 | **`UnboundListView` → `ListView`** | MeasureRow, QueryRowLevel for hierarchy |
| 6 | 2.12–2.15 | **`Button.Text` → `Content`** (shim remains) | PropertyTrigger, ControlTemplate, RenderCoalescer batch |

Details per release: [BREAKING_CHANGES.md](./BREAKING_CHANGES.md).

### XAML transforms (apply over time)

| Before | After |
|--------|-------|
| `DesignWidth` / `DesignHeight` | `Width` / `Height` (shim accepts both) |
| `DesignLeft` / `DesignTop` on **`TextBlock`**, **`Image`**, other legacy types | **Keep `Design*`** or use DLL with [POS layout shim](./POS_RUNTIME_FEEDBACK.md) — **`Margin` is ignored** on types without layout DPs |
| `<UnboundListView …/>` | `<ListView …/>` without `ItemsSource` |
| `{ThemeResource Key=…}` | `{DynamicResource …}` |
| `ThemesManager` `ActiveThemeName=""` with `{DynamicResource}` styles | Set **`ActiveThemeName`** to a valid theme key (e.g. `Default`) before styles apply |
| `<res:Path\Fragment/>` | Merged `ResourceDictionary` + `{StaticResource …}` |
| `Button` caption via `Text` | `Content` (layout engine aliases `Text` → `Content`) |
| `Scene BackColor="…"` | Style or widget API (not a Scene DP today; strict XAML rejects it) |
| Dialog `@` fragment templates | `DataTemplate` + `{Binding …}` (Phase 7 migration) |

### VB6 code

| Before | After |
|--------|-------|
| `ListView` owner-draw only via `UnboundListView` | `ListView` (same API surface) |
| Manual binding recreate on DataContext | Automatic rebind (Phase 4) |
| Batch layout refresh storms | Optional `New RenderCoalescer` + `BeginRenderUpdate` / `EndRenderUpdate` |

### Verification

- [ ] `.Tests/Phase0` — **30/30** pass (includes **P7a-SMOKE** POS SalesOrder shell)
- [ ] [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) §3 manual flows
- [ ] DeNovo EXE recompiled against pinned DLL

---

## Upgrading to 2.3.0 (Phase 2 — Grid, StackPanel, ContentControl)

### DLL pin

Update reference to **`v2.3.0-wpf-alignment-p2`**.

### New XAML (preferred for new screens)

| Pattern | Example |
|---------|---------|
| Vertical stack | `<StackPanel Orientation="Vertical"><Button Height="40"/></StackPanel>` |
| Grid rows/cols | `<Grid><Grid.RowDefinitions><RowDefinition Height="*"/></Grid.RowDefinitions></Grid>` |
| Single swap region | `<ContentControl><UserControl/></ContentControl>` |
| Border decorator | `<Border><Panel/></Border>` or `Child` DP in code |

### Verification

- [ ] `.Tests/Phase0` — **11/11** pass (includes P2-STACK, P2-STACK-LAY, P2-GRID)
- [ ] POS screens using **UniformGrid** unchanged; Collapsed children no longer consume cells

---

## Upgrading to 2.2.0 (Phase 1b — Border, UserControl, Window, Button)

### DLL pin

Update reference to **`v2.2.0-wpf-alignment-p1b`**.

### XAML

`Width` / `Height` now resolve on **Border**, **UserControl**, and **Window** (in addition to Panel). **Button** supports layout DPs internally; public `Width`/`Height` available on Button.

### Verification

- [ ] `.Tests/Phase0` — **8/8** pass (includes P1-BORDER)
- [ ] POS smoke unchanged on screens using Border/UserControl shells

---

## Upgrading to 2.1.0 (Phase 1 — layout core)

### DLL pin

Update reference to **`v2.1.0-wpf-alignment-p1`** (or 2.1.x build).

### XAML (recommended, not required yet)

| Before | After (Panel and future migrated types) |
|--------|----------------------------------------|
| `DesignWidth="200"` | `Width="200"` |
| `DesignHeight="40"` | `Height="40"` |

**Shim:** XAML reader still accepts `DesignWidth`/`DesignHeight` on types with layout DPs (aliased automatically).

### VB6 code (Panel)

| Before | After |
|--------|-------|
| `panel.DesignWidth = 100` | `panel.Width = 100` (preferred) |
| `panel.DesignWidth = 100` | Still works — forwards to `Width` on Panel |

### Verification checklist

- [ ] `.Tests/Phase0` passes (includes P1-WIDTH, P1-VIS)
- [ ] POS screens using `Panel` layout unchanged (legacy scale layout default)
- [ ] Optional: new XAML uses `Width`/`Height` on Panel

---

## Upgrading to 2.0.0 (Phase 0)

### DLL pin

Update `Reference=...\Demac.VCF.dll#` to the **`v2.0.0-wpf-alignment-p0`** tag (or later 2.x build).

### No required XAML or VB6 changes

Phase 0 is foundation-only. Existing apps continue to work with default **`StrictXamlLoad = False`**.

### Optional — enable strict XAML in tests

```vb
' In test bootstrap (Sub Main) before loading XAML:
VCF.StrictXamlLoad = True
```

When enabled, fix any load errors reported via **`XamlLoadException`** before POS production cutover (strict mode becomes default in a later phase).

### Optional — register app types in TypeRegistry

Reduce reliance on giant `ObjectConstructor` `Select Case` over time:

```vb
' After SetCustomConstructor, or in app init:
StaticClasses.TypeRegistry.Register "ShellWindow", "MyApp.ShellWindow"
StaticClasses.TypeRegistry.Register "SalesOrderView", "MyApp.SalesOrderView"
```

Registered names are resolved in `CreateInstance` before `CreateObject` and before `CustomConstructor`.

### Verification checklist

- [ ] App starts; login screen loads
- [ ] `.Tests/Phase0` passes (strict mode smoke)
- [ ] Two `ListCollectionView` instances on same screen (B1 regression)
- [ ] No change in Task Manager memory vs previous DLL

---

## Phase 7 — POS migration (in progress)

| Slice | Tag | Deliverable |
|-------|-----|-------------|
| **7a** | `v2.16.0-wpf-alignment-p7a` | [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md), 2.15 pin guide above, **P7a-SMOKE** |
| **7b** | `v2.17.0-wpf-alignment-p7b` | [Invoke-VcfXamlMigration.ps1](../tools/xaml-migrate/Invoke-VcfXamlMigration.ps1), [XAML_MIGRATION_PROMPTS.md](./XAML_MIGRATION_PROMPTS.md) |
| **7c-layout** | `v2.18.0-wpf-alignment-p7c-layout-shim` | POS layout shim + **P7c-LAY** + script legacy-type skip ([POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md)) |
| 7c-dialog | TBD | DeNovo `@` template → DataTemplate migration |

---

*Maintained by VCF team with POS validation from DeNovo.*
