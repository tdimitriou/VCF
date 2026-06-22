# VCF XAML migration tool

PowerShell script for **mechanical** POS XAML transforms (Phases 1–6). Non-destructive review modes included.

**Docs:** [XAML_MIGRATION_PROMPTS.md](../../docs/XAML_MIGRATION_PROMPTS.md) · [MIGRATION.md](../../docs/MIGRATION.md) · [POS_RUNTIME_FEEDBACK.md](../../docs/POS_RUNTIME_FEEDBACK.md)

---

## Transforms

| Transform | Before | After |
|-----------|--------|-------|
| **Layout** | `DesignWidth` / `DesignHeight` | `Width` / `Height` |
| **Layout** | `DesignLeft` / `DesignTop` | `Margin="L,T,0,0"` |
| **ListView** | `<UnboundListView …/>`, `TargetType="UnboundListView"` | `ListView` |
| **ThemeResource** | `{ThemeResource Key=X}`, `{ThemeResource X}` | `{DynamicResource X}` |
| **ButtonText** | `<Button Text="…"/>` | `<Button Content="…"/>` |

### Layout transform exceptions (legacy types)

**Skipped** on types without layout DPs — keeps `DesignLeft` / `DesignTop` / `DesignWidth` / `DesignHeight`:

`TextBlock`, `Image`, `Scene`, `UniformGrid`, `TextBox`, `WindowsFormsHost`

**Report-only:** legacy tags that already have `Margin` / `Width` / `Height` (from an earlier script run) flag **“needs layout shim DLL or revert to Design*”**.

See [POS_RUNTIME_FEEDBACK.md](../../docs/POS_RUNTIME_FEEDBACK.md). Pin **`v2.18.0-wpf-alignment-p7c-layout-shim`** (or later) when POS XAML was fully migrated before this script fix.

---

**Report-only (no edit):** `Scene BackColor=`, `res:` includes (incl. shortened paths), `@` dialog fragments, `Button` + `TextBlock` child pairs.

---

## Usage

From repo root or this folder:

```powershell
# Self-test (fixtures)
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -SelfTest

# Preview one file
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\SalesOrder.xml -WhatIf

# Preview tree
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse -WhatIf

# Apply (commit or backup first)
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse

# Manual-review scan only
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse -ReportOnly

# Subset
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\MyApp.xml -Transform ThemeResource,ListView
```

Requires **PowerShell 5.1+** (Windows). Does not require VCF DLL.

---

## Workflow (DeNovo)

1. Branch + backup XAML folder.
2. `-ReportOnly` on `UI/Resources/XAML` — triage manual items.
3. `-WhatIf` then apply mechanical script.
4. Cursor prompts from [XAML_MIGRATION_PROMPTS.md](../../docs/XAML_MIGRATION_PROMPTS.md) for Button Content, Scene BackColor, `res:`.
5. Pin DLL with layout shim if XAML already has `Margin` on `TextBlock` (see POS_RUNTIME_FEEDBACK).
6. Recompile DeNovo; run [POS_INTEGRATION_SMOKE.md](../../docs/POS_INTEGRATION_SMOKE.md).

---

## Fixtures

| File | Purpose |
|------|---------|
| `fixtures/before-sample.xml` | Input sample |
| `fixtures/expected-sample.xml` | Expected output after all transforms |

---

## Limits

- **Attribute order** may change (semantically equivalent).
- **Margin conflicts** (existing `Margin` + `DesignLeft`) are skipped and reported.
- Does **not** flatten `Button` > `TextBlock` trees (use Cursor Prompt 3).
- Does **not** migrate `res:` or `@` templates (Phase 7c dialog work).

---

*Updated for layout shim — `v2.18.0-wpf-alignment-p7c-layout-shim`*
