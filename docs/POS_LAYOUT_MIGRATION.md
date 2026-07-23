# POS layout migration — canvas scale retired (3.14.0)

**Audience:** DeNovo / POS teams consuming Demac.VCF  
**Related:** [POS_LAYOUT_RESIZE.md](./POS_LAYOUT_RESIZE.md) · [BREAKING_CHANGES.md](./BREAKING_CHANGES.md) **3.14.0** · harness samples under [`.Tests/DeNovoSmoke/Resources/XAML/Migrated/`](../.Tests/DeNovoSmoke/Resources/XAML/Migrated/)

---

## Contract (locked)

| Rule | Meaning |
|------|---------|
| **No canvas scale** | Absolute `Margin` / `Width` / `Height` on multi-child **UserControl**, **Border**, or **Panel** are **fixed pixels**. Resizing the host does **not** multiply them by `host/design`. |
| **Panels only** | Screens resize via **Grid** / **StackPanel** / **UniformGrid** / single-child **Border** (decorator fill). |
| **UC fills Window** | Window content still stretches the active **UserControl** to the client. |
| **UC single root** | A UserControl with **one** child (typically a root `Grid`) fills that child to the UC client. Prefer this pattern for every screen. |
| **Unset panel size** | Layout hosts default `Width`/`Height` **0** (content-driven Measure). Nested Grids in `Auto` rows must not rely on a 300px implicit size — prefer pixel chrome rows or explicit sizes. |

Reference implementations (VCF harness, not production DeNovo.vbp):

- `Migrated/Login/LoginViewWpf.xml` + `LoginPad.xml`
- `Migrated/MainMenu/MainMenuView.xml`
- `Migrated/SalesOrder/SalesOrderView.xml`
- `Migrated/Widgets/StatusBar.xml`

Prefer **design-sized** center cards (e.g. Login 574×358) inside star rows so chrome tracks the host without ballooning the pad. For `Content=` buttons, set **`FontSize` / `FontBold` / `ForeColor`** (LoginPad pattern) — string Content does not use `ButtonText12` TextBlock styles. UniformGrid children should **fill cells** (avoid MaxWidth that leaves empty pockets on wide hosts).

Legacy absolute trees remain under `Screens/...` for comparison until POS migrates.

---

## Checklist per screen

1. Wrap the UserControl in a **single root `Grid`** (no sibling absolute chrome on the UC).
2. Replace absolute column/row `Margin`+`Width` pairs with **`Grid.RowDefinitions` / `ColumnDefinitions`** (`Auto` / `*`).
3. Convert multi-child Borders (logo stacks, status strips) to **single-child Border → StackPanel/Grid**.
4. Keep **`Name=`** hooks and bindings used by VMs / smoke assertions.
5. Point resource keys at migrated XAML; resize Form client + `RelayoutChildren` (Shift+1/3 in DeNovoSmoke).
6. Confirm: columns still fill halves; star rows grow; no reliance on 0.5× Margin math.

---

## Before / after patterns

### Login — absolute UC siblings → root Grid

**Before (broken after 3.14.0):** logo, card, chrome buttons, StatusBar as **siblings** on the UserControl with design-canvas Margins.

**After:**

```text
UserControl
  Grid
    Row Auto — header (brand | ClockIn)
    Row *    — InnerBorder card (Grid: list | password+pad)
    Row Auto — Restart | Exit
    Row Auto — StatusBar
```

### Sales — absolute columns → star columns

**Before:** `LeftColumn` / `RightColumn` Borders with `Margin="16,72"` / `Margin="520,72"` and fixed `Width`/`Height`.

**After:**

```text
UserControl
  Grid
    Row Auto — title/subtitle
    Row *    — Grid *|* → LeftColumn | RightColumn
    Row Auto — chrome
    Row Auto — StatusBar
```

Left column: nested Grid `Auto` / `*` / `Auto` (title, ListView, total).  
Right column: title + `UniformGrid` for quick buttons.

### StatusBar

**Before:** multi-child Border with absolute `Margin` TextBlocks (only correct at design width).

**After:** single-child Border wrapping a **Grid** with `Auto`/`*` columns.

---

## DeNovoSmoke

| Key | Path |
|-----|------|
| Login (default) | `Migrated\Login\LoginViewWpf` (`modHarnessConfig.LoginViewResourceKey`) |
| Sales | `Migrated\SalesOrder\SalesOrderView` (`SalesOrderViewResourceKey`) |
| Legacy absolute Login | `Screens\Login\LoginView` when `USE_WPF_LOGIN_LAYOUT = False` |

**Resize check:** Login or Sales → **Shift+1 / Shift+3** — columns and star rows track the Form client; Immediate `[HARNESS-RESIZE]`.

Phase0 gate: **P7d-LAY-PANEL** (Grid star columns + ListView fill), not **P7d-LAY-RESIZE**.

---

## Scripts (later)

Ready-to-run XAML rewrite scripts for the full DeNovo.vbp tree can follow once these patterns stabilize. This cut documents the contract and ships the local Migrated samples only.

---

## Out of scope

- Migrating every production POS screen in the external DeNovo repo.
- Reintroducing `Design*` / `LegacyScaleLayout` / canvas-scale bridge.
