# POS layout resize — legacy vs WPF (DeNovo harness finding)

**Date:** 2026-06-27  
**Status:** Known gap · tracked for VCF backlog  
**Context:** DeNovoSmoke milestone 1 — Login screen windowed vs maximized  
**Related:** [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md) · [`.Tests/DeNovoSmoke/README.md`](../.Tests/DeNovoSmoke/README.md)

---

## 1. Summary

POS XAML is authored on a **1024×768 design canvas** but must run on **800×600**, **1024×768**, **widescreen**, and other sizes. Production runs **borderless** (full display = layout host).

Phase 1–6 introduced a **dual layout model**:

| Mode | Controls | Resize behavior |
|------|----------|-----------------|
| **Legacy scale** | `UserControl`, `Panel` (default `LegacyScaleLayout = True`) | `Design*` × `host/design` scale factors |
| **WPF-style** | `Border`, `Grid`, `StackPanel`, … (`LegacyScaleLayout = False`) | `Margin` + `Width`/`Height` DPs in **fixed pixels** |

**Observed on Login (harness + expected in production when size ≠ design canvas):**

- **Direct children** of the screen (`UserControl`) — side buttons (Clock In/Out, Restart, Exit) — **scale/re-anchor** with the host.
- **Children inside `Border`** (user list, numpad, password field) — stay at **design-time pixel** positions; the Border frame may grow while content stays top-left at ~574×358 design coords.

This is **not** harness-only. The v2.18 layout **load** shim does not change **resize** inside `Border`.

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
| **B** | `LegacyScaleLayout = True` on `Border` when children use `Design*` (POS mode) | Low | POS parity; delays pure WPF on Border |
| **C** | Propagate parent scale into Border child arrange | Medium | POS parity without global Border legacy flag |
| **D** | Migrate inner XAML to Grid/Margin; `LegacyScaleLayout = False` | High per screen | WPF-aligned; correct long-term |

**Recommendation:** **C** (framework) + **D** (screen-by-screen) for Login/Sales.

**Test:** add **P7d-LAY-RESIZE** — load Login at 1024×768, 800×600, 1366×768; assert inner panel content scales or grid fills (criteria depend on chosen option).

---

## 7. Harness notes

| Item | Action |
|------|--------|
| Milestone 1 scope | XAML load, bind, theme, navigation — **fixed 1024×768 client** acceptable |
| Resize parity | **Out of milestone 1** until §6 resolved |
| Optional harness | Borderless window; client area exactly 1024×768 (match production) |
| Re-sync XAML | `pos-v1/tools/Sync-DeNovoSmokeFixtures.ps1` |

---

## 8. Sign-off matrix

| Check | Owner | Status |
|-------|-------|--------|
| Document finding | VCF | **Done** (this file) |
| DeNovo confirms on production borderless | DeNovo | Pending |
| VCF implements B/C + P7d-LAY-RESIZE | VCF | Pending |
| Login inner panel Grid migration | DeNovo | Backlog |

---

*Maintained by VCF team · DeNovo resize screenshots 2026-06-27.*
