# VCF — window lifecycle (init, chrome, show, teardown)

**Audience:** VCF maintainers · DeNovo harness integrators  
**Companion:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) · [VCF_XAML_WPF_SUBSET.md](./VCF_XAML_WPF_SUBSET.md)  
**Last updated:** 2026-06-20  

---

## 1. Purpose

Documents the **ordered lifecycle** for `VCF.Window` + `IWindow` shells — especially **border chrome**, **XAML load timing**, and **shutdown** — validated in DeNovoSmoke (Phase 7d).

---

## 2. Initialization sequence

`NewWindow` → `Window.Initialize`:

```text
A. InitializeObject
   · Register DPs; SetCurrentValue BorderStyle = vbSizable (fallback)
   · Create(0) top-level HWND + Form.BorderStyle = 0  (borderless at create — avoids flash)
   · WidgetRoot, collections, UIElementBase
   (GetBaseStyle is NOT run here)

B. Application.Windows.Add(IWindow shell)

C. IWindow.InitializeComponent
   · Code: Win.DependencyProperties.SetValue "BorderStyle", 0  (DeNovo shell)
   · XAML: LoadSuperclassData → SetValue on root attributes (e.g. BorderStyle="2")
   · Local SetValue wins over style defaults (WPF semantics)

D. Set Style = GetBaseStyle
   · MyApp implicit <Style TargetType="Window"> → ApplyStyle
   · Style setters use SetCurrentValue — do not override local BorderStyle

E. SyncFormBorderStyle
   · Form.BorderStyle = GetValue("BorderStyle") if changed
   · Form is still hidden (Form.Show not called yet)
```

**Rule:** Host overrides (code or XAML root attributes) belong in **`InitializeComponent`**, which runs **before** implicit style apply.

---

## 3. BorderStyle precedence

| Priority | Source | Mechanism | Example |
|----------|--------|-----------|---------|
| 1 (highest) | Code in `InitializeComponent` | `SetValue` | `ShellWindow`: `BorderStyle = 0` |
| 1 | XAML root attribute | `SetValue` via `LoadSuperclassData` | `BorderTestWindow.xml`: `BorderStyle="2"` |
| 2 | Named window style | `Style` DP → `ApplyStyle` | `Style="{StaticResource DialogWindow}"` |
| 3 | MyApp implicit `Window` style | `GetBaseStyle` → `SetCurrentValue` | `MyApp.xml`: `BorderStyle=2` |
| 4 | Framework default | Metadata **`DefaultValue`** = `vbSizable` (**3.2.0**; not `SetCurrentValue`) | `vbSizable` fallback |

**DeNovo POS shell:** borderless via **`ShellWindow.IWindow_InitializeComponent`** (`SetValue 0`), overriding MyApp default `2`.

**Bordered dialog:** XAML `BorderStyle="2"` or implicit style only — no shell override.

---

## 4. Show sequence

```text
Win.Show (VCF.Window)
  1. SyncFormBorderStyle          ← safe: immediately before Form.Show
  2. Form.Show

Host may RelayoutChildren + WidgetRoot.Refresh after show (DeNovoSmoke pattern).
```

**Anti-pattern:** Do **not** call `SyncFormBorderStyle` or assign `Form.BorderStyle` from `CollectionChanged` / during `WRoot.Refresh` — caused form teardown, blank/hidden window, and IDE hangs in DeNovoSmoke testing.

---

## 5. IWindow lifetime

On `NewWindow`:

- `Application.Current.Base.Windows.Add ObjPtr(Shell), Shell` keeps the **app shell class** alive while the Cairo form exists.
- A local `Dim w As New ShellWindow` is **not** sufficient on its own.

**Consumer API:** `Window.Unload` then `Set Shell = Nothing` / `VCF.ClearApplication`.

`Window.Unload` sequence:

1. `LockRefresh` + `DetachBindingsTree` while Form/tree still live
2. `PrepareWindowChildren` — drop `WithEvents W` / Button timers while widgets exist
3. Transfer children to framework hold (no Release)
4. `Form.Unload` → unhook + `Windows.Remove`
5. `Set Form` / `WRoot = Nothing`
6. `DrainHold` — releases **only** a single held control (safe). Multiple ItemsControls are **deferred** (Release hangs vs compiled DLL; evidence in `%TEMP%\VCF_Unload.log`)

Do **not** call manual `DetachBindingsTree` or suite `Children.Clear` for cleanup — `Unload` owns it.

**Removed:** `Window.Dispose` (registry-only; raced / duplicated unload).

**Open:** safe Release of ItemsControl+Button graphs vs compiled ActiveX (IDE End may still freeze after MsgBox while deferred holds exist).

---

## 6. Shutdown (DeNovo / harness)

```text
Exit command or Form X (shell)
  → modHarnessAppManager.Shutdown / AppManager.Shutdown
  → StopTimers, StopClock
  → Cairo.WidgetForms.RemoveAll
  → Application.Run returns
  → modApp.ResetSession (IDE: clear globals for second F5)
```

**Secondary windows** (e.g. `BorderTestWindow`): close with X/Esc — **do not** call `RemoveAll`; only the shell should end the message loop.

---

## 7. DeNovoSmoke test hooks

| Action | Key / API | Validates |
|--------|-----------|-----------|
| Skip to Login | **Shift+L** | Navigation + Login load |
| Bordered window | **Shift+B** | `BorderTestWindow.xml` `BorderStyle=2` chrome |
| First-frame chrome gate | Phase0 **B-CHROME** (**3.25.0**) | `Form.BorderStyle` matches DP after `NewWindow` (0 + 2) |
| Border diag | **Shift+B** + Immediate (`Ctrl+G`) | `[BORDER-DIAG]` lines + VERDICT — `ENABLE_BORDER_CHROME_DIAG` in `modHarnessConfig.bas` |
| Exit | Exit button / shell X | Shutdown without error 91 |

See [`.Tests/DeNovoSmoke/README.md`](../.Tests/DeNovoSmoke/README.md).

---

## 8. Related backlog (not lifecycle — load performance)

Slow first paint / XAML load is **separate** from chrome lifecycle:

- Batch `LoadSuperclassData` with `LockRefresh` / `BeginRenderUpdate` — [VCF_WPF_ALIGNMENT_NOTES.md §2.7.4](./VCF_WPF_ALIGNMENT_NOTES.md)
- Lazy DP inheritance — §2.8
- Timed benchmarks **P7d-LOAD-*** — [VCF_PERFORMANCE_BENCHMARKS.md](./VCF_PERFORMANCE_BENCHMARKS.md)

---

## 9. Phase 2a.1 — auto layout after Visibility navigation

**Status:** **Done** (2026-07-21) — DeNovoSmoke Splash → Login → MainMenu → Sales; repeated nav + named lookup OK.

**Goal:** Visibility-based view swap on a **direct** `Window` child must not require the host to call `RelayoutChildren` / `RebuildNamedItemsList`.

| Piece | Status |
|-------|--------|
| `UserControl` Visibility → `Window.OnChildVisibilityChanged` (direct child only; skip when `LockRefresh`) | **Done** |
| DeNovoSmoke `HarnessNavigation` no longer calls Relayout/Rebuild after swap | **Done** |
| Option C (Border scale propagate) | **Deferred** — Option B + LoginViewWpf cover Login; no open edge-case reports |
| B-RESZ / B-NAV | **Done** — Phase0 `Phase2aBench_NestedUniformGridResize` / `Phase2aBench_ViewNavLeak` |

**Still call Relayout manually:** first show / resize presets (`ShellWindow`, `modHarnessResize`) — not Visibility-driven.

**Gate:** Splash → Login → MainMenu → Sales in DeNovoSmoke without layout/named-lookup regressions.

---

*Maintained by VCF team · validated DeNovoSmoke Phase 7d.*
