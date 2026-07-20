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

**Recommended first pin:** `v2.15.0-wpf-alignment-p6d` (Phases 0–6 only).

**Migrated POS XAML (Margin on TextBlock, etc.):** pin **`v2.18.0-wpf-alignment-p7c-layout-shim`** — see [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md).

---

## 2. VCF repo regression (before POS)

| Check | Command / location | Pass |
|-------|-------------------|------|
| Phase0 suite | `.Tests/Phase0` → `RunAll` | **33/33** (includes **B-RESZ**, **B-NAV**) |
| Strict XAML | `VCF.StrictXamlLoad = True` in test bootstrap | B-STRICT-* pass |
| DeNovoSmoke harness | `.Tests/DeNovoSmoke` (when scaffold lands) | Milestone screens per [README](../.Tests/DeNovoSmoke/README.md) |

**Phase 1 sign-off (UI/XAML):** **Complete** on tag **`v2.23.0-wpf-alignment-phase1-complete`** — Phase0 + DeNovoSmoke Splash → Login → MainMenu → Sales layout-only. **Not** full `DeNovo.exe`. See [DENOVO_HARNESS_PROPOSAL_RESPONSE.md](./DENOVO_HARNESS_PROPOSAL_RESPONSE.md).

**Phase 2a (VCF parallel):** Framework tags validated here (§2b) + Phase0 only — Data/Kernel may still be unfinished.

---

## 2b. DeNovoSmoke harness — Phase 1 / Phase 2a gate

Minimal exe: **VCF + vbRichClient5 only** (no KernelLib / Data / DB).

| # | Screen | Pass criteria |
|---|--------|---------------|
| 1 | **SplashView** | Loads; layout OK; theme resolves |
| 2 | **LoginView** / **LoginViewWpf** | Stub `DataContext`; bindings; focus |
| 3 | **MainMenuView** | Login → MainMenu; Logout; named lookup |
| 4 | **SalesOrderView** | Layout-only columns; stub lines; Back → MainMenu |

Pin (Phase 1 exit): **`v2.23.0-wpf-alignment-phase1-complete`**. Later Phase 2a tags: same harness regression.

**Phase 2b** (full POS): §3 below when Data + Kernel + UI are ready together.

---

## 3. POS manual smoke (Phase 2b — full DeNovo.exe)

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
| `Design*` XAML | `UI/Resources/XAML/` | Run [Invoke-VcfXamlMigration.ps1](../tools/xaml-migrate/Invoke-VcfXamlMigration.ps1) then [XAML_MIGRATION_PROMPTS.md](./XAML_MIGRATION_PROMPTS.md) |
| `UnboundListView` | Dialog grids | → `ListView` (see MIGRATION 2.9.0) |
| `{ThemeResource}` | `MyApp.xml` styles | → `{DynamicResource}` |
| `res:` fragments | XAML includes | → `ResourceDictionary` / `MergedDictionaries` |
| `Button.Text` | Caption setters | → `Content` (alias shim for `Text` in layout engine) |
| `@` dialog templates | MessageBox / DialogWindow | Phase 7 — migrate to `DataTemplate` + binding |

---

## 5. Sign-off

| Role | Name | Date | Tag | Result |
|------|------|------|-----|--------|
| VCF | | | | Phase0 + DeNovoSmoke milestone green |
| DeNovo | | | | Harness §2b pass (2a); then POS smoke §3 when Phase 2b kicks off |

---

*Part of Phase 7 — POS migration support.*
