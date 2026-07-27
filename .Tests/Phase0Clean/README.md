# Phase0Clean — framework-owned teardown gate

**Purpose:** Same assertions as Phase0 where ported, but **no KeepAlive / no suite DetachBindingsTree / no Dispose**. Lifetime is owned by `Window.Unload` ? `Form_Unload`.

## Rules

1. Parent live controls under `Shell.Base.Children` (`Park`).
2. End every session with:
   - `Shell.Base.Unload`
   - `Set Shell = Nothing`
   - `Set AppHost = Nothing`
   - `VCF.ClearApplication`
3. Do **not** call `VCF.DetachBindingsTree` for cleanup (allowed only as an **assertion** that detach works).
4. `Phase0App.Class_Terminate` must **not** call `ClearApplication` (suite owns that).

## Seed suite (5)

| Id | What |
|----|------|
| P0-GOLDEN | GoldenPanel.xml load |
| P4-ONEWAY | Binding + Unload |
| P7c-PANEL | ItemsPanel UniformGrid (freeze path) |
| B-CHROME | First-frame BorderStyle |
| B-NAV | Visibility swap + Unload |

## How to run

1. Rebuild `Demac.VCF` (IDE or Make).
2. Open `Phase0CleanGroup.vbg` (or `Phase0Clean.vbp` vs compiled DLL).
3. F5 ? Immediate should show all PASS, then MsgBox; IDE must stay responsive for a second F5.

## Migrating remaining Phase0 tests

Copy a bench from `.Tests/Phase0/Modules/modPhase0Bench.bas`, then:

- Delete `KeepAlive` / Fail-path KeepAlive.
- Wrap with `BeginSession` / `Park` / `EndSession` (or per-window `Unload`).
- Replace `Win.Dispose` with `Win.Unload`.
- Remove suite `DetachBindingsTree` / `WidgetForms.RemoveAll` used only for cleanup (B-NAV leak **assertion** can stay).

Legacy `.Tests/Phase0` remains until the full port is green vs compiled VCF.