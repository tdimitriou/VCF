# POS layout resize — legacy vs WPF (DeNovo harness finding)

**Date:** 2026-06-27 · **Updated:** 2026-07-23 (3.14.0 canvas-scale retired)  
**Status:** Bridge **retired** · Measure/Arrange is the only resize path  
**Context:** DeNovoSmoke — Login / Sales resize  
**Related:** [POS_LAYOUT_MIGRATION.md](./POS_LAYOUT_MIGRATION.md) · [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md) · [BREAKING_CHANGES.md](./BREAKING_CHANGES.md) **3.0.0** / **3.7.0** / **3.9.0**–**3.14.0** · [`.Tests/DeNovoSmoke/README.md`](../.Tests/DeNovoSmoke/README.md)

---

## 0. Locked contract (3.14.0)

| Layer | Role |
|-------|------|
| **Retired** | Design-canvas scale (`hostWidget / hostDP` on UC / multi-child Border / Panel) — **removed in 3.14.0** |
| **Only path** | `Grid` / `StackPanel` / `UniformGrid` / single-child **Border** decorator + Measure/Arrange |
| **UserControl** | Window fills UC; UC with a **single root** child fills that child to the client |
| **Absolute Margin/Width** | Fixed pixels on multi-child hosts — migrate screens per [POS_LAYOUT_MIGRATION.md](./POS_LAYOUT_MIGRATION.md) |

```text
Width/Height/Margin DPs
        └──► Grid / Stack / Border fill  ──► Measure/Arrange (3.9–3.14)
                    (canvas-scale bridge retired)
```

**Measure/Arrange epic progress:**

| Slice | Status |
|-------|--------|
| `ActualWidth` / `ActualHeight` + dirty flags + `InvalidateMeasure` | **3.9.0** |
| `StackPanel` Measure→Arrange via `MeasureLayout` / desired sizes | **3.9.0** / **P2-STACK-MEAS** |
| `Grid` MeasureLayout + Auto via `MeasureElementSize` | **3.10.0** / **P2-GRID-MEAS** |
| `Border` single-child decorator MeasureLayout | **3.11.0** / **P1-BORDER-MEAS** |
| `ContentControl` MeasureLayout (IUIElement Content) | **3.12.0** / **P6e-CC-MEAS** |
| `UniformGrid` MeasureLayout (max-cell × rows/cols) | **3.13.0** / **P2a-UG-MEAS** |
| Retire UC / multi-child Border / Panel canvas-scale | **3.14.0** / **P7d-LAY-PANEL** |
| Grid cell `HorizontalAlignment` / `VerticalAlignment` | **3.15.0** / **P2-GRID-ALIGN** |
| TextBlock: text align ≠ layout align; `Move` does not write Width/Height DPs | **3.15.x fix** (DeNovoSmoke MainMenu wrap) |
| `DockPanel` + `DockPanel.Dock` + `LastChildFill` | **3.17.0** / **P2-DOCK-XAML** / **P2-DOCK-LAY** |

**P7d-LAY-PANEL:** Grid star columns + nested ListView track host size on resize (not 0.5× Margin math).

---

## 1. Summary

POS XAML is authored on a **1024×768 design canvas** but must run on **800×600**, **1024×768**, **widescreen**, and other sizes. Production runs **borderless** (full display = layout host).

**Historical (pre-3.0.0) dual model:**

| Mode | Controls | Resize behavior |
|------|----------|-----------------|
| **Legacy scale** | `UserControl`, `Panel` (`LegacyScaleLayout = True`) | `Design*` × `host/design` |
| **WPF-style** | `Border`, `Grid`, … (`LegacyScaleLayout = False`) | fixed `Margin`/`Width`/`Height` pixels |

**3.0.0:** `Design*` / `LegacyScaleLayout` removed.  
**3.7.0:** canvas-scale **bridge** restored on DP-based absolute trees (same math, new property names) so DeNovo/Sales nested Borders resize again.

---

## 2. Evidence (DeNovoSmoke)

| View | What happens |
|------|----------------|
| **~1024×768 framed window** | Layout roughly correct; slight mismatch because XAML targets **client** 1024×768 while the harness window includes **title bar/chrome**. |
| **Maximized / wide** | Outer Border chrome scales; **inner** list + keypad do not fill the panel; corner buttons track the host. |

Screenshots (2026-06-27): windowed vs maximized Login — inner panel empty space with content clustered top-left.

**Production difference:** POS is **borderless** → at 1024×768 client there is no chrome gap; resize issue remains whenever host aspect ratio or size differs from design canvas.

---

## 3. How legacy resize works (design canvas)

```text
scaleX = hostWidth  / hostDesignWidth   (typically 1024)
scaleY = hostHeight / hostDesignHeight  (typically 768)

left   = DesignLeft   × scaleX
width  = DesignWidth  × scaleX
… same for Y
```

Implementation: `modLayoutEngine.LayoutRectFromDesign` · `FrameworkElement.ArrangeChildren` when `LegacyScaleLayout = True`.

**Note:** Independent X/Y scale (not uniform) is **legacy POS behavior**, not WPF default.

---

## 4. Why Border breaks nested scaling

`Border` sets `LegacyScaleLayout = False` in `Border.cls` (Phase 1b WPF alignment).

Nested POS XAML still uses **`DesignLeft/Top/Width/Height`** inside Borders. Those are arranged via `LayoutRectFromMargin` — **absolute pixels**, no scale factor — see `FrameworkElement.ArrangeChildren` else branch.

```text
LoginView (UserControl, legacy scale)
├── Button …                    ← scales ✓
├── Border InnerBorder …        ← outer box scales as UserControl child ✓
│   ├── Border ListWrapper      ← fixed Design* inside Border ✗
│   ├── TextBox, UniformGrid/LoginPad
└── Button …                    ← scales ✓
```

---

## 5. Will final placement match WPF specs?

**Short answer:** **Not with legacy `Design*` scaling alone.** WPF alignment and legacy POS resize are **different targets**.

| Target | Placement / sizing | Matches WPF? |
|--------|-------------------|--------------|
| **Legacy POS (pre–Phase 1b nested Border)** | All elements use proportional `Design*` scale | **No** — VB6/Cairo convention |
| **Current mixed engine (Phase 1–6)** | Split: legacy on shell, fixed pixels inside Border | **No** — regression on nested screens |
| **Interim fix (B/C below)** | Restore nested proportional scale for POS | **No** — restores POS parity, not WPF |
| **WPF north star (agreed program)** | `Margin`/`Width`/`Height` DPs, `Grid`/`StackPanel`, `LegacyScaleLayout = False`, Measure/Arrange | **Yes** — per [VCF_WPF_ALIGNMENT_NOTES.md](./VCF_WPF_ALIGNMENT_NOTES.md) §3, [VCF_XAML_WPF_SUBSET.md](./VCF_XAML_WPF_SUBSET.md) |

**WPF-final behavior (examples):**

- Login panel → `Grid` with star columns or fixed `*` tracks, not `DesignLeft` on every child.
- Resize → layout from **Margin**, **Grid row/column definitions**, **Stretch** / alignment — not `DesignLeft × scaleX`.
- Ultrawide → WPF uses grid/star/stretch policies; legacy POS stretches X and Y independently (may differ from WPF on the same XAML until migrated).

**Explicit non-goals:** pixel-identical WPF layout rounding, BAML, full animation — see alignment notes §2.

**Practical path for DeNovo:**

1. **Near term:** framework fix **B or C** so migrated POS XAML **behaves like old POS** on all resolutions (harness + production sign-off).
2. **Medium term:** migrate high-traffic screens (Login, Sales) inner panels to **Grid + Margin** XAML.
3. **Long term:** drop `Design*` and `LegacyScaleLayout` when POS tree is on WPF layout DPs.

---

## 6. VCF backlog options

| ID | Option | Effort | Outcome |
|----|--------|--------|---------|
| **B** | Multi-child `Border` + `DesignLeft`/`DesignTop` → `LegacyScaleLayout`; single-child always decorator fill | Low | **Done (2026-07-18)** — POS multi-child scale; LoginViewWpf ListView fills ListWrapper |
| **C** | Propagate parent scale into Border child arrange | Medium | POS parity without Design* detection |
| **D** | Migrate inner XAML to Grid/Margin; `LegacyScaleLayout = False` | High per screen | WPF-aligned; correct long-term |

**Recommendation (2026-07-23):** **3.14.0** retires the canvas-scale bridge. Absolute Margin trees must migrate to Grid/Stack — see [POS_LAYOUT_MIGRATION.md](./POS_LAYOUT_MIGRATION.md).

**Test:** **P7d-LAY-PANEL** (Phase0) — Grid `*` columns at 400×300 → 200×150; assert panel fill (not 0.5× absolute Margin).

### Deferred — WPF Margin / Padding defaults

**Status:** **Phase 2a** — content-control families + UniformGrid decision shipped.

| Family | Status |
|--------|--------|
| **ListView** | **Done** — `Margin` default 0; `Padding` default **4,1,4,1**; draw uses DP; **P2a-PAD** |
| **TextBox / Button** | **Done** — TextBox Margin=0 Padding=**1**; Button Padding=**1** (Aero2); draw/`InnerSpace` wired; **P2a-PAD-TB** |
| **UniformGrid** | **Done** — keep default **Padding=2** (cell inset); gated **P2a-PAD-UG**; apps may set `Padding="4"` |
| Layout hosts (Grid/Border/Panel) | Stay **Margin=0**; no content Padding unless WPF requires |

**Goal:** Align framework `Margin`/`Padding` defaults with WPF (prefer Win10-era metrics where themes differ), so content controls share one model instead of one-off paint insets.

---

## 7. Harness notes

| Item | Action |
|------|--------|
| Milestone 1 scope | XAML load, bind, theme, navigation — **fixed 1024×768 client** acceptable |
| **7e reference layout** | `.Tests/DeNovoSmoke/Resources/XAML/Screens/Login/LoginViewWpf.xml` — root **Grid**, inner panel **Grid** (list \| password+pad); toggle `modHarnessConfig.USE_WPF_LOGIN_LAYOUT` |
| Resize parity | Manual checklist in [`.Tests/DeNovoSmoke/README.md`](../.Tests/DeNovoSmoke/README.md) § resize validation |
| Harness shell | Borderless **1024×768** client (`ShellWindow`, `BorderStyle=0`) — [WINDOW_LIFECYCLE.md](../WINDOW_LIFECYCLE.md) |
| Border chrome test | **Shift+B** → `BorderTestWindow.xml` (`BorderStyle=2`) | Manual first-frame check |
| Load perf (7f) | Lazy Login + Immediate `[P7d-LOAD-*]`; XAML `LoadSuperclassData` batching | See [VCF_PERFORMANCE_BENCHMARKS.md](../VCF_PERFORMANCE_BENCHMARKS.md) §7f |
| Resize (7e) | **Shift+1/2/3** preset client sizes; optional `USE_SIZABLE_SHELL_BORDER` for drag | Borderless shell has no resize grips |
| Option B | `Border` auto-`LegacyScaleLayout` when children have Design* | Toggle `USE_WPF_LOGIN_LAYOUT=False` to exercise legacy Login inner scale |
| Re-sync XAML | `pos-v1/tools/Sync-DeNovoSmokeFixtures.ps1` (legacy `LoginView.xml` only) |

---

## 8. Sign-off matrix

| Check | Owner | Status |
|-------|-------|--------|
| Document finding | VCF | **Done** (this file) |
| DeNovo confirms on production borderless | DeNovo | Pending |
| VCF implements B/C + P7d-LAY-RESIZE | VCF | **B + P7d-LAY-RESIZE Done**; C optional |
| Login inner panel Grid migration | VCF harness **7e** (`LoginViewWpf.xml`) · DeNovo production XAML | **In progress** (harness reference) |
| Phase 7f load benchmarks + XAML batching | VCF | **Done** (Splash 35–41 ms, Login 448–670 ms) |

---

*Maintained by VCF team · DeNovo resize screenshots 2026-06-27.*
