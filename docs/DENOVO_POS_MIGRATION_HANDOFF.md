# DeNovo / POS — migration handoff (VCF 3.26 stable)

**Audience:** DeNovo team migrating POS onto the current VCF pin  
**VCF pin:** **`Demac.VCF.dll` 3.26.0** (Phase0 **90/90**, DeNovoSmoke green)  
**VCF status:** Feature development **paused** — bugfix only unless migration finds a blocker.

This package covers **POS plan steps 1–2** (XAML + re-wire). Steps 3–6 stay on the DeNovo roadmap after the pin is green.

---

## POS plan (full) vs this handoff

| Step | Work | In this handoff? |
|------|------|------------------|
| **1** | Migrate existing XAML definitions | **Yes** — script + prompts + layout guide |
| **2** | Re-wire POS to VCF 3.26 contract | **Yes** — wiring checklist below |
| **3** | InvoiceGrid → `ListView` | **Later** — ActiveX cannot host in `VCF.Window` |
| **4** | RootPanel → top-level `IWindow` | **Later** — blocked until InvoiceGrid gone |
| **5** | Move MainForm functionality to classes | **Later** |
| **6** | Retire legacy MainForm | **Later** |

**Why 3 before 4:** InvoiceGrid is a classic ActiveX control on the legacy `MainForm` (ZOrder over VCF). It cannot sit inside `VCF.Window`. RootPanel/`IWindow` cutover only after ListView fully replaces InvoiceGrid.

---

## What to pin and read

| Item | Location |
|------|----------|
| DLL | Rebuild/register **`Demac.VCF.dll` 3.26.0**; copy into DeNovo lib path |
| Breaking log | [BREAKING_CHANGES.md](./BREAKING_CHANGES.md) — especially **3.0.0** (Design*), **3.14.0** (no canvas scale), **3.26.0** (auto Relayout) |
| Master upgrade | [MIGRATION.md](./MIGRATION.md) — **Upgrading to 3.26.0** |
| Layout redesign | [POS_LAYOUT_MIGRATION.md](./POS_LAYOUT_MIGRATION.md) |
| Mechanical script | [tools/xaml-migrate/README.md](../tools/xaml-migrate/README.md) |
| Cursor prompts | [XAML_MIGRATION_PROMPTS.md](./XAML_MIGRATION_PROMPTS.md) |
| Window / Relayout | [WINDOW_LIFECYCLE.md](./WINDOW_LIFECYCLE.md) |
| Smoke | [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) · [`.Tests/DeNovoSmoke/README.md`](../.Tests/DeNovoSmoke/README.md) |
| ListView (for step 3 later) | [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md) |

**Harness reference XAML** (migrated patterns, not production DeNovo):  
`.Tests/DeNovoSmoke/Resources/XAML/Migrated/`

---

## Step 1 — XAML migration workflow

### 1.1 Branch and scan

```powershell
# From Demac.VCF repo (or copy script into DeNovo tools/)
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -SelfTest

# On POS tree (adjust path)
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse -ReportOnly
```

Triage **Manual review** lines (`res:`, Scene `BackColor`, `@` fragments, leftover `Design*`, Button+TextBlock).

### 1.2 Mechanical transforms

```powershell
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse -WhatIf
# backup / commit, then:
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse
```

| Transform | Result |
|-----------|--------|
| `DesignWidth` / `DesignHeight` | `Width` / `Height` |
| `DesignLeft` / `DesignTop` | `Margin="L,T,0,0"` |
| `UnboundListView` | `ListView` |
| `{ThemeResource …}` | `{DynamicResource …}` |
| `Button Text=` | `Content=` |

**Skipped (no layout DPs on 3.26):** `Scene`, `WindowsFormsHost`, `Image` — keep sizing via parent/`Move` or wrap; do **not** expect `Design*` at runtime (**3.0.0** removed public Design*).

### 1.3 Judgment / AI pass

Use [XAML_MIGRATION_PROMPTS.md](./XAML_MIGRATION_PROMPTS.md) (updated for 3.26):

- Button + single TextBlock → `Content=`
- Scene `BackColor=` → child / Style
- `res:` → ResourceDictionary / MergedDictionaries
- Absolute Margin trees that must **resize** → Grid/Stack per [POS_LAYOUT_MIGRATION.md](./POS_LAYOUT_MIGRATION.md)

Mechanical Margin on multi-child hosts is **fixed pixels** (no canvas scale after **3.14.0**). Screens that must track window size need panel redesign — not another script pass.

### 1.4 XAML acceptance

- [ ] No `Design*` remaining on types that register layout DPs  
- [ ] No `UnboundListView` / `Button Text=` / `{ThemeResource` in migrated trees  
- [ ] `StrictXamlLoad` does not reject migrated screens  
- [ ] Spot-check Login / Sales / MainMenu against Migrated harness samples  

---

## Step 2 — Re-wire POS to VCF 3.26

### Required

| Area | Action |
|------|--------|
| **DLL pin** | 3.26.0; **recompile** DeNovo EXE (Project Compatibility) |
| **Themes** | `ActiveThemeName` non-empty when using `{DynamicResource}` |
| **Constructors** | App `IObjectConstructor` / TypeRegistry for custom views; prefer built-in ResourceDictionary over `res:` over time |
| **DataContext** | No manual “recreate bindings” — **P4-DCTX** rebinds automatically |
| **Relayout** | Do **not** call `Window.RelayoutChildren` / `RebuildNamedItemsList` for Show, resize, or Visibility nav — framework owns them (**3.26.0**). Public APIs remain escape hatches only. |
| **Shutdown** | Do not manual `DetachBindingsTree` before `RemoveAll` — `Form_Unload` handles it |
| **Borderless shell** | `BorderStyle=0` via `SetValue` in `IWindow_InitializeComponent` when using VCF Window — [WINDOW_LIFECYCLE.md](./WINDOW_LIFECYCLE.md) |

### Still OK on legacy shell (until steps 3–4)

- RootPanel / child window on VB6 `MainForm`  
- InvoiceGrid ActiveX + ZOrder over VCF surface  
- Manual Form client sizing for the host Form (VCF `Window` Relayout applies when the VCF window itself resizes)

### Wiring acceptance

- [ ] App starts with pinned DLL; XAML loads under strict mode  
- [ ] Login → main navigation; bindings update on DataContext change  
- [ ] Theme brushes resolve  
- [ ] No Relayout calls required after Visibility view swap  
- [ ] [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) §3 as applicable under current shell  

---

## Out of scope until later (do not block step 1–2)

- InvoiceGrid → ListView parity (hierarchy, MeasureRow, owner-draw)  
- RootPanel → top-level `IWindow`  
- Retiring `MainForm`  
- New VCF controls (ComboBox, TabControl, ScrollViewer, …) — VCF backlog Phases 3+  

When step 2 is green, VCF will reopen feature work only for **blockers** found in migration or for step 3 ListView needs.

---

## VCF team contact / gates

| Gate | Owner |
|------|--------|
| Phase0 **90/90** | VCF — before any new pin |
| DeNovoSmoke | VCF — harness patterns |
| POS smoke on pin | DeNovo — after XAML + wire |
| Bugfix on 3.26 | VCF — layout/bind/strict regressions only |

---

*Prepared for DeNovo POS migration pause — Demac.VCF 3.26.0.*
