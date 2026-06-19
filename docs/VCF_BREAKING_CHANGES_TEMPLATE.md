# Demac.VCF — breaking changes log

**Template for VCF repo `docs/BREAKING_CHANGES.md`**  
Copy to VCF repo and update on each breaking release.

---

## How to use

Each **major** or **breaking minor** release gets a section below. Link from `CHANGELOG.md` and tag notes.

---

## [Unreleased] — WPF alignment program

### API removals (planned)

| Removed | Replacement | Phase |
|---------|-------------|-------|
| Public `DesignLeft/Top/Width/Height` | `Width`, `Height`, `Margin` DPs | 1 |
| `UnboundListView` class | `ListView` with ItemsSource=Nothing | 5 |
| `NestedProperty` | Internal to `BindingExpression` | 4 |
| `ThemeResource` markup | `{DynamicResource}` | 3 |
| `OverlayWidget` | ControlTemplate / visual states | 1 |
| `CallByName` XAML property fallback | DP-only setters | 3 |

### Behavior changes (planned)

| Change | Migration |
|--------|-----------|
| x:Class load failure no longer falls back to root tag | Fix ObjectConstructor / type registration |
| ItemsSource requires INCC or IEnumerable adapter | Wrap arrays in ObservableCollection |
| Hidden vs Collapsed visibility semantics | Audit POS Visibility attributes |
| DataContext change rebinds all expressions | Remove manual binding recreate TODOs |

### File deletions (planned)

- `Classes/_Image.cls`, `_TextBlock.cls`, duplicate `MarkupExtensions.cls`
- `IDependencyPropertyCallbackListener.cls` stub
- `Modules/API.bas` duplicate

---

## Release template

```markdown
## [X.Y.Z] — YYYY-MM-DD

### Breaking

- **Area:** Description. **Migration:** one-line fix.

### Deprecated (remove in X+1)

- ...
```

---

*Maintained by VCF team. POS migration steps: see `MIGRATION.md`.*
