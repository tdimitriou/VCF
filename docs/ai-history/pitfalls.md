# Known pitfalls / do-not-regress

Hard-won from Phase0 and DeNovoSmoke debugging. Prefer these over rediscovering in chat.

## Layout & Design*

- **Do not resurrect Design* / LegacyScaleLayout / canvas-scale multiply** as the resize model. Consumers must migrate to panels + fixed pixels.
- **Do not run Option B and Option C scale paths together** — they fight (Option C cancelled anyway).
- Unset Width/Height defaulting to **0** (Auto-track) vs old 300 — Absolute trees that assumed 300 break; document/migrate.
- **ListView width** that ignores window resize is often **layout tree** (fixed Width / wrong star row), not a ListView paint bug — check host Grid first.
- Casting: `Child As Object` then `.Widget` fails — cast to **`IControl`** (or equivalent) before Widget access.

## Bindings & collections

- Binding detach / DC rebind can **freeze the IDE** if expressions stay hooked across terminate — Detach before release; Me-safe `Class_Terminate` (no self TypeName in error paths).
- `ObservableCollection` event sinks: procedure signatures must match; parentheses required on `Call OnCollectionChanged(...)`.
- ItemsSource historically required **ObservableCollection** — wrong type raises; don’t swallow in tests.
- Don’t double-fire **ListIndexChanged** and **SelectionChanged** without an intentional dual-raise contract.

## Templates / styles / DP

- **TemplateBinding / ContentPresenter slot** bugs often look like “HAlign wrong” — check ApplyStyle order and Current vs Local writes.
- Do not implement “trigger snapshot/restore” as if it completed full WPF precedence — contradicts 3.2.0 two-slot contract.
- **HoverColor = -1** had special “no Cairo default hover” meaning — don’t replace with a literal theme color without checking prior semantics.
- Buttons **without Command** may skip enablement paths that drive refresh — Enabled=False vs missing Command are different; hover/click paint depends on enabled state.
- Lookless / ContentTemplate work must not leave **orphan public debug members** on the COM surface.

## ListView

- **Do not converge dual presenters** into per-row visual trees without an explicit new decision — owner-draw is POS-critical.
- Bound vs owner-draw mode keyed off **`ItemsSource Is Nothing`** — don’t break that switch.
- Dense bind / clone paths caused **AV / IDE crash** when walking disposed clone trees — harden hotspot tests; avoid use-after-free in template clone loops.
- **Variable height (`FixedRowHeight = False`):** DeNovo hit **row metrics vs scrollbar desync** (content height / scroll range). Invalidate + rebuild metrics on data/width changes; verify `VScrollBar.Max` / paint offsets against `mTotalRowHeight` before signing off InvoiceGrid.

## Window / Cairo / ActiveX

- **KeepAlive of compiled DLL widgets after RunAll deadlocks the IDE** — never reintroduce suite-wide KeepAlive across the ActiveX boundary.
- **`SyncFormBorderStyle` during CollectionChanged / Refresh** caused blank windows, hide-on-click, IDE hangs.
- Window refs: local `Dim w As New Shell` is insufficient — **Application.Windows** must hold the shell.
- Shutdown: null-check in terminate/unregister paths; Binding terminate was a real crash source.
- **`API.ObjFromPtr` / weak parents:** never log+re-raise on expected freed pointers during teardown — silent Resume Next.
- **ItemsControl Release after `Form.Unload`** is unsafe for item visuals / emptied UniformGrid hosts vs compiled DLL — don’t “fix” by leaking forever without documenting; real fix is ordered Release while sinks/widgets live.
- Tier-1 automation hazards: line numbers on `Attribute … VB_UserMemId = -4` break `For Each`; missing `Exit Sub` before `Handler:` causes `Err.Raise 0`.

## VB6 / tooling

- **LF line endings** → class/project won’t load (“could not be loaded”).
- Reserved names: `Empty`, `Local` (use another property name), careful with `App` vs `VB.App` when hosts shadow `App`.
- `WithEvents` only on event sources — not arbitrary panels; not in standard modules.
- Interface event sink signatures must match exactly or compile fails with cryptic procedure-mismatch errors.
- Codesmart / IDE add-ins can muddy End/Stop crashes — bisect with add-ins unloaded when diagnosing compiled teardown.

## Process

- Don’t edit DeNovo production `.vbp` from VCF work unless explicitly requested — use DeNovoSmoke / Migrated copies.
- Don’t declare the WPF line “shipped to POS” while compiled unload hangs.
- After wild diagnostic branches: **restore tip → cherry-pick proven fixes only**.
