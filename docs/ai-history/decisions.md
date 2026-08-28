# Major decisions and why

Ordered roughly by impact. Applies primarily to the **v2 WPF line** unless noted.

## Strategy & repo

| Decision | Why |
|----------|-----|
| **WPF-semantics north star** for VCF (not “WPF clone”) | POS developers should get familiar effective-value / layout / XAML behavior with minimal surprise; stay light on VB6/Cairo. |
| **Finish VCF fundamentals before POS rewire** | Breaking changes churn; migration docs + pin after the framework is coherent. DeNovo is a consumer, not part of VCF scope. |
| **Parallel track (Phase 2a VCF-local)** when POS/kernel not ready | Keep shipping gated VCF work + local Migrated XAML / DeNovoSmoke without blocking on full POS rebuild. |
| **Single repo, `v1/` + `v2/` trees** (Jul 2026) | Production must stay on a binary-compatible DLL; WPF line paused with teardown debt — dual trees avoid mixing. Remote: `tdimitriou/VCF` with branch `v1` + WPF branches. |
| **Pin production to June 20 2026 / Sep 2025-compatible `v1`** | Compiled WPF DLL + Phase0/DeNovo teardown freezes made v2 unusable for urgent POS work; June 20 sources rebuilt and ran DeNovo without dependent rebuild. **Registered DLL = v1.** |
| **Name WPF line `VCF2` / `Demac.VCF2` (or similar)** | Avoid ProgID/typelib clashes with registered production `Demac.VCF`. Folder is `v2/`; project rename may still be incomplete. |
| **Pause v2; resume later** | Great progress; few remaining glitches (esp. compiled ActiveX unload). Prefer stable POS over fighting teardown mid-crisis. |

## Compatibility

| Decision | Why |
|----------|-----|
| **Project Compatibility during rewrite**, not Binary Compatibility | Surface will change often (add/remove public members/classes). Binary Compat freezes typelib signatures and blocks progress. Binary freeze is a later finish-line policy item. |
| **Breaking changes OK if documented** | Each tagged minor: `BREAKING_CHANGES.md` + `MIGRATION.md` + Phase0 gates. |

## Layout / Design*

| Decision | Why |
|----------|-----|
| **Remove `Design*` and `LegacyScaleLayout` (3.0.0)** | Design-canvas scale is not WPF-aligned and is a poor pattern; Measure/Arrange + panels cover POS needs. |
| **Retire design-canvas scale bridge (3.14.0)** | After Design* removal, proportional host/design multiply on UC/Border/Panel was still non-WPF; migrate screens to Grid/Stack absolute pixels. |
| **Option C (parent-scale Border propagate) cancelled** | Superseded once Design* / scale bridge gone. |
| **Default padding/margins → WPF-like standards** | Document first; implement as lower priority than fundamentals (partially done for ListView/TextBox/etc.). |

## Dependency properties / styles

| Decision | Why |
|----------|-----|
| **Interim two-slot DP stack (3.2.0)** — Local > Current > Inherit > Metadata default | Full WPF layering is a multi-month epic; hover/style already work via re-`ApplyStyle`. |
| **Full DP precedence epic deferred** (~2–3 months when queued) | Documented in `VCF_DP_PRECEDENCE_ROADMAP.md`. Do not fake layers with trigger snapshot/restore. |
| **Trigger “restore” = re-`ApplyStyle`** (interim) | Matches two-slot contract; condition changes re-apply setters + active triggers. |
| **`Visible` bool kept as compat facade** | `False` → Collapsed; breaking enum-only removal not queued. |
| **Attached DP bag primary; nested dict shim retained** | RegisterAttached shipped; full shim removal partial / deferred. |

## Controls / content model

| Decision | Why |
|----------|-----|
| **VB6 has no class inheritance** — composition + `Implements` + FE host | “Button is ContentControl” means shared semantics / interfaces, not true subclassing. |
| **ListView dual presenters permanent** (bound + owner-draw) | Owner-draw is load-bearing for POS grids (InvoiceGrid path). Converging to per-row visual trees risks paint/perf regressions. |
| **Add ComboBox/TabControl/ScrollViewer after fundamentals** | Avoid building chrome on incomplete DP/layout/content model. Selector base shipped; concrete controls deferred (“soon” when resumed). |
| **ContentTemplate / lookless Button path pursued on v2** | WPF-aligned lookless tree preferred over forever self-paint; ContentTemplateSelector still deferred. |

## Window / lifecycle

| Decision | Why |
|----------|-----|
| **Keep `IWindow` shells in `Application.Windows`** | Unlike VB6 forms, VCF Window instances die when refs drop; collection keeps them alive for Cairo message loop. |
| **Borderless create + `SyncFormBorderStyle` before `Show`** | Avoid chrome flash; never sync border during CollectionChanged / mid-refresh. |
| **Framework owns unload/detach** (not host KeepAlive) | Host KeepAlive of DLL widgets deadlocks IDE vs compiled ActiveX; prefer `Window.Unload` + Cairo remove. |
| **Compiled DLL unload of ItemsControl graphs still open** | Evidence: hang in DrainHold / Disarm after Form.Unload. Suite may MsgBox-green then IDE End freezes. |

## Testing / quality bar

| Decision | Why |
|----------|-----|
| **Phase0 suite is the gate** for every slice | No “done” without green Phase0 (grew to ~89–90 gates on v2 tip). |
| **DeNovoSmoke is local harness**, not DeNovo.vbp edit | Validate POS-shaped XAML without touching production POS until ready. |
| **Prefer targeted fixes over mass Tier-1 error handlers** | Broad On Error GoTo + line-number scripts caused Err=0 fall-through, Attribute corruption, and noise; useful briefly for diagnosis, then restore. |
| **CRLF only for VB6 sources** | LF-only `.cls`/`.vbp` fail to load in VB6 IDE. |
| **Error logging pattern (when used):** `On Error GoTo Handler` + line numbers + log; never `TypeName(Me)` / self-refs in `Class_Terminate` handlers. |

## POS migration (consumer plan, agreed)

1. Migrate XAML definitions  
2. Re-wire POS to new VCF surface  
3. Replace InvoiceGrid with ListView (blocks pure `IWindow` root — ActiveX needs legacy VB Form host)  
4. RootPanel → top-level `IWindow`  
5. Move MainForm behavior into standard classes  

Hand off docs + migration script when VCF is finished enough; do not block VCF work on POS rewire.
