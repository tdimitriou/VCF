# Conventions established

## Engineering

- **WPF alignment over legacy convenience** when they conflict (Design*, canvas scale, naming). Document breaks; migrate consumers.
- **Composition + `Implements`**, not fake inheritance trees. Shared FE host / IContentControl / ISelector semantics.
- **Phase0 gate every slice** — fail the ship if the suite regresses. Prefer small tagged minors over mega-commits when releasing.
- **Project Compatibility** while the public surface is fluid; update test apps to the new typelib. Binary Compat only when freezing a release line.
- **VB6 source encoding:** Windows **CRLF**; prefer Windows-1252 for IDE load. Never leave LF-only `.cls`/`.bas`/`.vbp`.
- **Qualify ambiguous types:** `New VCF.TextBox`, `New VCF.Image`, `As VCF.Image` — VB6 built-ins and reserved names (`Empty`, `Local`, `Image`) bite otherwise.
- **`Set` for object properties** in tests (`Set Btn.Style = St`).
- **No `New` on non-creatable public bases** (`Window`, some views) — use factories / `NewWindow` / concrete shells.
- **Public COM surface:** Prefer framework-internal dirty work. Escape-hatch layout APIs may stay public temporarily but hosts should not need them for normal show/resize/Visibility.

## Layout & XAML

- Absolute layout = **fixed Margin/Width/Height pixels** after Design* retirement — not silent host/design scale.
- Prefer **Grid / StackPanel / DockPanel / Canvas / Border** composition for resize.
- Strict XAML: unknown types/properties raise (`XamlLoadException`); do not reintroduce CallByName for unknown attrs.
- Migrated sample screens live under harness **Migrated/** (or equivalent); keep originals as absolute references when useful.

## Dependency properties

- Writers today: **Local** (XAML, CLR Let, Binding, TemplateBinding) vs **Current** (style setters, triggers, many init seeds).
- Triggers deactivate via **re-ApplyStyle**, not a separate layer — until the precedence epic.
- Attached layout props: RegisterAttached + bag; invalidate measure/arrange on change.
- Metadata `DefaultValue` preferred over `SetCurrentValue` seeds where piloted (e.g. Window.BorderStyle).

## Window lifecycle

- Create borderless (`Create(0)`), apply chrome via DP, **`SyncFormBorderStyle` immediately before `Form.Show`**.
- Keep shells in **`Application.Windows`**.
- Teardown: **`Window.Unload`** owns detach; do not invent host KeepAlive of DLL widgets.
- Documented in `WINDOW_LIFECYCLE.md` (v2).

## Docs & release

- Living docs: `BREAKING_CHANGES.md`, `MIGRATION.md`, handoff guide, class/property registries.
- Tag format historically descriptive (`v3.xx.x-…`); bump MinorVer with BREAKING entry.
- Consumer handoffs: pin a DLL tag + checklist + script; do not silently edit DeNovo.vbp from VCF sessions unless asked.

## Testing

- Run Phase0 after framework changes; DeNovoSmoke after layout/chrome/window changes.
- Measure IDE vs compiled separately — **in-process green ≠ compiled ActiveX green**.
- After failed experiments: restore to last known good tip, re-apply **only** proven diffs.
