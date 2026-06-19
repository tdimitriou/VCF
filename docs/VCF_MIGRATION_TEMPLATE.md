# Demac.VCF — migration guide

**Template for VCF repo `docs/MIGRATION.md`**  
Copy to VCF repo and expand per release.

---

## Audience

POS (DeNovo) and internal Demac apps upgrading `Demac.VCF.dll`.

---

## General upgrade process

1. Read **`BREAKING_CHANGES.md`** for your target version.
2. Pin new DLL in app `.vbp` / references.
3. Run **`.Tests`** and app smoke tests.
4. Apply XAML transforms (below).
5. Apply VB6 code transforms (below).
6. AI-assisted bulk migration (DeNovo): Cursor with this doc + XAML folder context.

---

## Phase 1 — Layout (`Design*` → DPs)

### XAML transform

| Before | After |
|--------|-------|
| `DesignWidth="200"` | `Width="200"` |
| `DesignHeight="40"` | `Height="40"` |
| `DesignLeft="10" DesignTop="20"` | `Margin="10,20,0,0"` (adjust per layout panel) |

### VB6 code

| Before | After |
|--------|-------|
| `ctrl.DesignWidth = 100` | `ctrl.Width = 100` (when exposed as DP) |
| `MoveChild` scale assumptions | Remove; layout from Measure/Arrange |

### Optional one-release shim

If VCF ships reader alias `DesignWidth` → `Width` with `#pragma deprecation` log, use only during bulk XAML pass.

---

## Phase 3 — Resources

### `res:` includes

| Before | After |
|--------|-------|
| `<res:Screens\Sales\Menu.xml/>` | Merge into ResourceDictionary; `{StaticResource MenuTemplate}` |
| `ObjectConstructor` `Case "res...."` | Remove case; framework resolver |

### Theme brushes

| Before | After |
|--------|-------|
| `{ThemeResource Key=PrimaryBrush}` | `{DynamicResource PrimaryBrush}` |

---

## Phase 4 — Bindings

### DataContext

Remove manual `' TO-DO: Recreate the Bindings` — framework rebinds automatically after upgrade.

### ItemsSource

| Before | After |
|--------|-------|
| Plain `List` or array | `ObservableCollection` or register IEnumerable adapter |

---

## Phase 5 — ListView

| Before | After |
|--------|-------|
| `UnboundListView` | `ListView` (no ItemsSource) |
| `ListView.Base.ListIndex = i` | `SelectedIndex` / `SelectedItem` binding |
| Dialog `@Selected` hacks | `{Binding SelectedItem, Mode=TwoWay}` |

---

## ObjectConstructor shrink pattern

Keep in POS `ObjectConstructor`:

- View shells (`*View`, `*Window`)
- ViewModels, converters, app services

Move to VCF:

- All `VCF.*` control types (already CreateObject)
- `res:` fragment loading
- Markup extensions

---

## Verification checklist

- [ ] App starts; login screen loads
- [ ] Sales screen: grids, buttons, bindings
- [ ] Theme switch (if used)
- [ ] Modal dialogs
- [ ] No leak after 20 view navigations (Task Manager stable)
- [ ] Invoice / order line list (when migrated from Codejock)

---

## Release-specific sections

*(Add below per semver tag.)*

### [Unreleased]

Program start — see [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md).

---

*Maintained by VCF team with POS validation from DeNovo.*
