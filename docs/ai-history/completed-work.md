# Completed work

Status as of last productive v2 tip (~**3.25.x**, tag notes like `v3.25.1-border-test-esc`) and later v1 backports. Phase0 was **89/89** (later 90 with B-AUTO-LAY) on green project-group runs.

## Program phases (v2 framework)

| Phase | Delivered |
|-------|-----------|
| **0 — Foundation** | TypeRegistry, XamlLoadException, Phase0 harness, BREAKING/MIGRATION living docs |
| **1 — Layout core** | DP registry, FrameworkElement, Visibility, Measure/Arrange, Design* removed (3.0) |
| **2 — Panels** | Grid, StackPanel, Border decorator, ContentControl, UniformGrid, **DockPanel (3.17)**, **Canvas (3.21)** |
| **3 — Resources / strict XAML** | ResourceDictionary, Merged, DynamicResource, DP-only setters, unknown type/prop raise |
| **4 — Bindings / collections** | BindingExpression, DC rebind, Detach, OC, LCV, ItemsControl, Selector; RS/EN; UST/Delay; attached Path; caret; CanExecute |
| **5 — ListView** | Bound + owner-draw merged; MeasureRow; hierarchy indent; dual presenter locked |
| **6 — Templates / polish** | PropertyTrigger, ControlTemplate lookless, ContentTemplate, TemplateBinding, render coalesce |
| **8 — Inheritance** | Batch inherit (8a) + lazy GetValue parent walk (8b) |
| **2a VCF-local backlog** | Visibility relayout, padding defaults, themes, layout epic, chrome gate — cleared through **3.25.0 B-CHROME** |

## Notable shipped capabilities (v2)

- **Layout:** Public Measure/ArrangeOverride; Desired/Actual; Min/Max constraints; IsHitTestVisible (3.24); TextBlock/Image Visibility (3.23); attached layout invalidation (3.22)
- **DP:** Registry + bags; ClearValue; SetCurrentValue; metadata defaults (pilot); RegisterAttached for Grid/Dock/Canvas
- **Content:** IContentControl / ContentHost; live ContentPresenter; lookless Button path
- **Items:** ItemsControl + ItemsPanelTemplate (Stack/UniformGrid); Selector + SelectionChanged dual-raise (3.6)
- **Themes:** ActiveThemeName merge; System → Light/Dark
- **Window:** Create(0) + SyncFormBorderStyle; WINDOW_LIFECYCLE contract
- **Canvas-scale retirement (3.14):** Migrated Login/Sales/MainMenu samples under Migrated/; POS_LAYOUT_MIGRATION docs; P7d-LAY-PANEL gate
- **DeNovoSmoke harness** + sync scripts for POS-shaped fixtures
- **Migration tooling:** `tools/xaml-migrate`, prompts, DeNovo handoff package (~3.26 docs wave)

## v1 (stable) completed after dual-tree split

- June 20 2026 sources rebuilt, registered HKLM; DeNovo ran without dependent rebuild
- **Variable row height** for ListView (bound and/or unbound) **without binary break**; usage notes handed to DeNovo; tagged/pushed

## Explicitly cancelled / superseded

- Option C Border parent-scale  
- Design* / LegacyScaleLayout as a supported layout model  
- Dual ListView presenter “converge to visual-tree rows”  
- Optional VCF.Core utilities DLL split (never required)
