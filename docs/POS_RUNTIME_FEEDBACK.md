# POS runtime feedback — DeNovo integration (2026-06-19)

**Audience:** VCF maintainers  
**Consumer:** DeNovo `pos-v1` · pin `v2.15.0-wpf-alignment-p6d`  
**Status:** **Integrated on `master`** — tag **`v2.18.0-wpf-alignment-p7c-layout-shim`** (validate Phase0 in IDE, then DeNovo re-pin).

**DeNovo handoff summary:** denovo monorepo → `docs/migration/VCF_TEAM_HANDOFF_POS_INTEGRATION.md`

---

## 1. DLL changes (VCF-owned)

| File | Description |
|------|-------------|
| `Modules/modLayoutEngine.bas` | `ApplyLegacyLayoutProperty` — when `SetProperty` cannot set `Margin`/`Width`/`Height`, map to `IUIElement.Design*` |
| `Classes/XAMLReader.cls` | Invoke shim from `SetProperty` after `CallByName` failure |
| `.Tests/Phase0` | **P7c-LAY** — `PosMigratedTextBlockLayout.xml` |
| `tools/xaml-migrate/` | Skip layout transform on legacy types; report migrated legacy tags |

---

## 2. Problem statement

### 2.1 Mechanical XAML migration vs layout engine

`Invoke-VcfXamlMigration.ps1` converts:

| Before | After |
|--------|-------|
| `DesignLeft` / `DesignTop` | `Margin="L,T,0,0"` |
| `DesignWidth` / `DesignHeight` | `Width` / `Height` |

**But** on pin `2.15.0`:

- `UserControl` / `Panel` default **`LegacyScaleLayout = True`** (`FrameworkElement.Initialize`).
- `ArrangeChildren` uses **`DesignLeft`/`DesignTop`/`DesignWidth`/`DesignHeight`** when legacy mode is on — see `FrameworkElement.cls` `LayoutRectFromDesign`.
- **`TextBlock`** has **no** `Margin`/`Width`/`Height` DPs — only private `m_Design*` fields. XAML `Margin` is dropped in `SetProperty` (`On Error Resume Next`).
- **`TextBlock`** defaults **`ScaleFont = True`**. `GetScaleFactor` uses root `Widget.Width / DesignWidth`. With children at (0,0) and wrong dimensions, fonts render at extreme scale.

**Observed on POS:** Splash labels piled at top-left; grey default theme background; `Run-time error 91` then `80010007` during `LoadUI`; process crash after dismissing VCF error dialog.

### 2.2 Migration script scope mismatch

From `tools/xaml-migrate/README.md`:

- Script **does** transform `Design*` → `Margin`/`Width`/`Height`.
- Script **does not** migrate `res:` includes (manual).
- Script **does not** validate that target types support layout DPs.

**Additional POS finding:** Some manual/script edits **stripped `res:` paths** (e.g. `res:Widgets\StatusBar` → `res:Widgets`). That is a **consumer XAML defect**, but the script should **report** ambiguous `res:Prefix` without path.

### 2.3 Themes

POS `MyApp.xml` had `ActiveThemeName=""` with styles using `{DynamicResource …}`. Empty active theme breaks dynamic resource resolution on pin `2.15.0`. **DeNovo** set `ActiveThemeName="Default"`; document as required for `{DynamicResource}` apps.

---

## 3. Shim design (draft in tree)

**Intent:** Transitional compatibility for **mechanically migrated POS XAML** until `TextBlock`/`Image`/etc. register layout DPs or `UserControl` switches to `LegacyScaleLayout=False` with full Margin arrange.

**Not a substitute for:** Migrating `TextBlock` to `FrameworkElement` layout DPs (see `BREAKING_CHANGES.md` Phase 1b+ roadmap).

**Limitations:**

- Maps **Margin left/top only** (not right/bottom thickness semantics).
- Does not fix parent `LegacyScaleLayout` — still uses `Design*` at arrange time (shim populates `Design*` from migrated XAML).
- `Button` etc. with real `Margin` DP may double-apply if `CallByName` succeeds — shim only runs on failure.

---

## 4. Recommended permanent fixes (VCF backlog)

| Option | Effort | Notes |
|--------|--------|-------|
| **A. Migration script** | Low | Skip `Margin` transform for types without `Margin` DP; keep `DesignLeft`/`DesignTop` on `TextBlock` |
| **B. Shim (current draft)** | Low | Keep until POS XAML regenerated |
| **C. `TextBlock` + layout DPs** | Medium | Register `Width`/`Height`/`Margin`; `LegacyScaleLayout=False` on parents |
| **D. `UserControl` default** | Medium | `LegacyScaleLayout=False` when all children use Margin — breaking for unmigrated screens |

**Recommendation:** **A + B** for immediate POS unblock; **C** for proper WPF alignment.

---

## 5. Migration doc updates needed

Add to `MIGRATION.md` (Upgrading to 2.15.0 section):

1. **Layout transform is type-dependent** — `Margin` only on controls with layout DPs (`Panel`, `Border`, `Button`, `Grid`, …).
2. **`TextBlock` / legacy elements** — keep `DesignLeft`/`DesignTop` **or** use shim DLL ≥ *TBD tag*.
3. **`ActiveThemeName`** — set non-empty default before `{DynamicResource}` styles in `Application.Resources`.
4. **`res:` includes** — preserve full path (`res:Widgets\StatusBar`); never shorten to `res:Widgets`.

Update `tools/xaml-migrate/README.md` transform table accordingly.

---

## 6. Test cases to add (Phase0 or POS shell)

| ID | Load XAML | Assert |
|----|-----------|--------|
| POS-LAYOUT-1 | Migrated splash (`Margin` on `TextBlock`) | Text blocks not at (0,0); font size ≈ 30px / 12px |
| POS-LAYOUT-2 | `res:Widgets\StatusBar` fragment | Resolves; no error 91 |
| POS-LAYOUT-3 | `MyApp.xml` + `ActiveThemeName="Default"` | `{DynamicResource}` styles resolve |

Existing **P7a-SMOKE** uses legacy `Design*` — add migrated-XAML variant.

---

## 7. DeNovo fixes (out of scope for VCF DLL)

These remain in `denovo` repo — listed for context only:

- Restored `res:` paths in QuickService / Login / MainMenu XAML  
- `SplashView.xml` — `Design*` + orange background + `ScaleFont="0"`  
- `AppManager.Terminate`, `LoadUI` error logging  
- POSWidgets rebuild, `VCF_2X_BUILD_ORDER.md`

---

## 8. Sign-off workflow

| Step | Owner |
|------|--------|
| Integrate §1 patches (or script-only fix) | VCF |
| Phase0 + new layout tests | VCF |
| Tag + `BREAKING_CHANGES.md` entry | VCF |
| Pin tag, recompile, `regsvr32`, smoke §3 | DeNovo |
| Update `DENOVO_VCF_MIGRATION_POLICY.md` pin row | DeNovo |

---

*Maintained by VCF team after handoff. DeNovo reports smoke results against tagged builds.*
