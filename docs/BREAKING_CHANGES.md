# Demac.VCF — breaking changes log

**Program:** WPF alignment (Phases 0–7)  
**Maintained by:** VCF team · POS migration steps in [MIGRATION.md](./MIGRATION.md)

---

## [2.0.0] — 2026-06-20 — Phase 0 (foundation)

### Breaking

- **Orphan source removed:** Unregistered duplicate files deleted from the repo (`_Image.cls`, `_TextBlock.cls`, duplicate `MarkupExtensions.cls`, stub `IDependencyPropertyCallbackListener.cls`, orphan `Modules/API.bas`). **Migration:** None — these were not in `Demac.VCF.vbp`.

### Added (non-breaking until strict mode enabled)

- **`XamlLoadException`** — structured XAML load errors (element, property, line context).
- **`TypeRegistry`** — register app types by name; used by `CreateInstance` before `CreateObject`.
- **`VCF.StrictXamlLoad`** — when `True`, malformed XML and unknown types raise `XamlLoadException` instead of returning `Nothing`. Default **`False`** for POS compatibility; enable in `.Tests/Phase0` and CI.

### Bug fixes

- **B1:** `ListCollectionView.Initialize` — static init flag replaced with per-instance initialization (second view no longer blocked).

### Deprecated (remove in Phase 1+)

- Public `DesignLeft/Top/Width/Height` → `Width`, `Height`, `Margin` DPs (Phase 1).
- `UnboundListView` → merged `ListView` (Phase 5).
- `ThemeResource` markup → `{DynamicResource}` (Phase 3).
- `CallByName` XAML property fallback → DP-only setters (Phase 3).

---

## [Unreleased] — planned (Phases 1–7)

See [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md) and [VCF_BREAKING_CHANGES_TEMPLATE.md](./VCF_BREAKING_CHANGES_TEMPLATE.md).

---

## Release template

```markdown
## [X.Y.Z] — YYYY-MM-DD

### Breaking

- **Area:** Description. **Migration:** one-line fix.

### Deprecated (remove in X+1)

- ...
```
