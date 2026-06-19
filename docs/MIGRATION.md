# Demac.VCF — migration guide

**Audience:** POS (DeNovo) and internal Demac apps upgrading `Demac.VCF.dll`.

**Handoff baseline:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md)

---

## General upgrade process

1. Read **[BREAKING_CHANGES.md](./BREAKING_CHANGES.md)** for your target version.
2. Pin new DLL in app `.vbp` / references.
3. Run **`.Tests/Phase0`** and app smoke tests.
4. Apply XAML transforms (per phase below).
5. Apply VB6 code transforms (per phase below).

---

## Upgrading to 2.0.0 (Phase 0)

### DLL pin

Update `Reference=...\Demac.VCF.dll#` to the **`v2.0.0-wpf-alignment-p0`** tag (or later 2.x build).

### No required XAML or VB6 changes

Phase 0 is foundation-only. Existing apps continue to work with default **`StrictXamlLoad = False`**.

### Optional — enable strict XAML in tests

```vb
' In test bootstrap (Sub Main) before loading XAML:
VCF.StrictXamlLoad = True
```

When enabled, fix any load errors reported via **`XamlLoadException`** before POS production cutover (strict mode becomes default in a later phase).

### Optional — register app types in TypeRegistry

Reduce reliance on giant `ObjectConstructor` `Select Case` over time:

```vb
' After SetCustomConstructor, or in app init:
StaticClasses.TypeRegistry.Register "ShellWindow", "MyApp.ShellWindow"
StaticClasses.TypeRegistry.Register "SalesOrderView", "MyApp.SalesOrderView"
```

Registered names are resolved in `CreateInstance` before `CreateObject` and before `CustomConstructor`.

### Verification checklist

- [ ] App starts; login screen loads
- [ ] `.Tests/Phase0` passes (strict mode smoke)
- [ ] Two `ListCollectionView` instances on same screen (B1 regression)
- [ ] No change in Task Manager memory vs previous DLL

---

## Future phases (preview)

| Phase | Topic | Doc section |
|------:|-------|-------------|
| 1 | `Design*` → layout DPs | See template below |
| 3 | `res:` → ResourceDictionary | See template below |
| 4 | Binding rebind on DataContext | Remove manual TODO comments |
| 5 | `UnboundListView` → `ListView` | See template below |

<details>
<summary>Phase 1 — Layout (`Design*` → DPs)</summary>

| Before | After |
|--------|-------|
| `DesignWidth="200"` | `Width="200"` |
| `DesignHeight="40"` | `Height="40"` |
| `DesignLeft="10" DesignTop="20"` | `Margin="10,20,0,0"` |

</details>

<details>
<summary>Phase 3 — Resources</summary>

| Before | After |
|--------|-------|
| `<res:Screens\Sales\Menu.xml/>` | Merged ResourceDictionary; `{StaticResource MenuTemplate}` |
| `{ThemeResource Key=PrimaryBrush}` | `{DynamicResource PrimaryBrush}` |

</details>

<details>
<summary>Phase 5 — ListView</summary>

| Before | After |
|--------|-------|
| `UnboundListView` | `ListView` (no ItemsSource) |
| Dialog `@Selected` hacks | `{Binding SelectedItem, Mode=TwoWay}` |

</details>

---

*Maintained by VCF team with POS validation from DeNovo.*
