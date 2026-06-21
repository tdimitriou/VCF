# POS integration smoke — DeNovo checklist

**When:** After pinning a new `Demac.VCF.dll` tag (start with **`v2.15.0-wpf-alignment-p6d`**).  
**Where:** `pos-v1` (DeNovo) · reference VCF docs in this repo.

---

## 1. Pin and build

| Step | Action |
|------|--------|
| 1 | Read [BREAKING_CHANGES.md](./BREAKING_CHANGES.md) and [MIGRATION.md](./MIGRATION.md) for target tag |
| 2 | Copy `Demac.VCF\bin\Demac.VCF.dll` (or CI artifact) into DeNovo lib path |
| 3 | Update `DeNovo.vbp` reference if path/GUID changed |
| 4 | **Full recompile** DeNovo EXE (Project Compatibility — mandatory) |
| 5 | `regsvr32 Demac.VCF.dll` on test machines |

**Recommended first pin:** `v2.15.0-wpf-alignment-p6d` (Phases 0–6 complete).

---

## 2. VCF repo regression (before POS)

| Check | Command / location | Pass |
|-------|-------------------|------|
| Phase0 suite | `.Tests/Phase0` → `RunAll` | 29/29 (30/30 after Phase 7a) |
| Strict XAML | `VCF.StrictXamlLoad = True` in test bootstrap | B-STRICT-* pass |

---

## 3. POS manual smoke (required)

Run on a **dev DB** with typical configuration. Record build tag and date.

| # | Flow | What to verify |
|---|------|----------------|
| 1 | **Login** | Shell loads; no XAML load errors; keyboard focus OK |
| 2 | **Sales** | Sales order screen opens; left/right columns render; UniformGrid layout |
| 3 | **Line grid** | Add line item; grid updates; scroll if applicable |
| 4 | **Dialog** | Open one modal (e.g. payment / lookup); close without leak |
| 5 | **Commit path** | Save or finalize one transaction (sandbox) |
| 6 | **Navigation** | Switch view twice; bindings update (DataContext rebind) |
| 7 | **Memory** | Task Manager: process **&lt; 100 MB** idle after smoke (no video on second display) |

---

## 4. Known migration touchpoints (2.x)

| Area | POS location | Action |
|------|--------------|--------|
| `Design*` XAML | `UI/Resources/XAML/` | Migrate to `Width`/`Height`/`Margin` over time; shims still work |
| `UnboundListView` | Dialog grids | → `ListView` (see MIGRATION 2.9.0) |
| `{ThemeResource}` | `MyApp.xml` styles | → `{DynamicResource}` |
| `res:` fragments | XAML includes | → `ResourceDictionary` / `MergedDictionaries` |
| `Button.Text` | Caption setters | → `Content` (alias shim for `Text` in layout engine) |
| `@` dialog templates | MessageBox / DialogWindow | Phase 7 — migrate to `DataTemplate` + binding |

---

## 5. Sign-off

| Role | Name | Date | Tag | Result |
|------|------|------|-----|--------|
| VCF | | | | Phase0 green |
| DeNovo | | | | POS smoke §3 pass |

---

*Part of Phase 7 — POS migration support.*
