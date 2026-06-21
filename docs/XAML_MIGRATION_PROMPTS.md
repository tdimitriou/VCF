# XAML migration — Cursor prompts

**Audience:** DeNovo / POS team bulk-migrating `UI/Resources/XAML/` to VCF 2.15+ patterns.  
**Companion:** [tools/xaml-migrate/README.md](../tools/xaml-migrate/README.md) (PowerShell mechanical transforms) · [MIGRATION.md](./MIGRATION.md)

Run **mechanical transforms first**, then use these prompts in Cursor with the XAML folder (or single file) open.

---

## Setup (every session)

1. Pin **`v2.15.0-wpf-alignment-p6d`** or later (`v2.16.0+` migration docs).
2. Attach context: `@MIGRATION.md`, `@BREAKING_CHANGES.md`, `@VCF_XAML_WPF_SUBSET.md`.
3. Point Cursor at the file or folder under migration (e.g. `pos-v1/UI/Resources/XAML`).

---

## Prompt 1 — Scan folder (report only)

```
You are migrating Demac POS XAML to VCF 2.15 WPF-alignment patterns.

Read MIGRATION.md "XAML transforms" table. Scan every *.xml in this folder and produce a markdown report:

1. Count of DesignWidth/DesignHeight/DesignLeft/DesignTop (layout legacy)
2. UnboundListView occurrences (element + Style TargetType)
3. {ThemeResource …} markup extensions
4. res: includes or res: paths
5. Scene BackColor= attributes
6. Button with Text= attribute vs Button wrapping only a TextBlock child
7. @ fragment bindings (Phase 7c)

Do NOT edit files. Group by file path. Flag files safe for Invoke-VcfXamlMigration.ps1 vs needing manual work.
```

---

## Prompt 2 — Mechanical pass (after PowerShell script)

After running:

```powershell
.\tools\xaml-migrate\Invoke-VcfXamlMigration.ps1 -Path .\UI\Resources\XAML -Recurse -WhatIf
```

Use:

```
Review the git diff from Invoke-VcfXamlMigration.ps1 on this XAML folder.

Verify:
- DesignLeft/DesignTop became Margin="L,T,0,0" only where appropriate (not inside Grid cells that will move to Grid attached properties later)
- UnboundListView → ListView preserves Name, events, and owner-draw (no ItemsSource)
- {ThemeResource Key=X} and {ThemeResource X} → {DynamicResource X}
- Button Text= → Content= only on Button elements

Fix any incorrect mechanical edits. Do not change res: includes or Scene BackColor yet. Keep VB6 color literals (&H…, decimal) unchanged.
```

---

## Prompt 3 — Simple Button → Content

```
For each Button in this file that contains ONLY one TextBlock child (text caption, no Image sibling):

Replace with a single self-closing or empty Button using Content=:
- Move Text="{Binding …}" from TextBlock to Content="{Binding …}"
- Move static Text="…" to Content="…"
- Move FontSize/FontBold/ForeColor from TextBlock to Button when they are caption styling
- Remove the TextBlock child

Skip buttons with Image + TextBlock (mixed content) — list those for manual review.

Match existing indentation and attribute quoting. One file at a time.
```

---

## Prompt 4 — Scene BackColor (strict XAML)

```
Scene does not register BackColor as a dependency property. StrictXamlLoad rejects Scene BackColor=.

For each Scene with BackColor= in this file:
1. Remove BackColor from Scene XAML
2. Suggest equivalent: set background on root child (UniformGrid/Panel) OR a Scene Style setter in MyApp.xml
3. Document if color must stay on Scene via VB6 code after load

Do not invent new DPs. Prefer moving BackColor to the first visual child that already supports it.
```

---

## Prompt 5 — res: fragment (manual)

```
Migrate res: XAML includes toward VCF ResourceDictionary patterns (Phase 3).

For this file:
1. List each res: reference and target fragment path
2. Propose MergedDictionaries + Source= or inline ResourceDictionary keys
3. Replace element instantiation with {StaticResource Key} or loaded template reference
4. Note ObjectConstructor.cls cases that can be deleted after migration

Keep behavior identical. Flag dialog-specific @ templates for Phase 7c (DataTemplate), do not merge with screen fragments in one step.
```

---

## Prompt 6 — Single screen end-to-end

```
Migrate this screen XAML file to VCF 2.15 strict-safe subset:

1. Apply layout DP renames (Design* → Width/Height/Margin) where not already done
2. UnboundListView → ListView for owner-draw grids
3. ThemeResource → DynamicResource in setters
4. Simple Button+TextBlock → Content
5. Remove Scene BackColor or relocate to supported element
6. Leave res: and @ dialog templates unchanged but list them in a comment block at top: <!-- MANUAL: … -->

Output the full migrated file. List remaining manual items for QA.
```

---

## Prompt 7 — Styles dictionary (MyApp.xml)

```
Migrate Styles/UnboundListView and MyApp.xml theme references:

1. TargetType="UnboundListView" → TargetType="ListView"
2. {ThemeResource Key=…} and {ThemeResource …} → {DynamicResource …}
3. Do not rename brush keys in the theme dictionary
4. Preserve Setter Property names (BackColor not Background)

Show diff hunks only for Style and Setter lines changed.
```

---

## Verification after AI + script

- [ ] `Invoke-VcfXamlMigration.ps1 -SelfTest` pass (VCF repo)
- [ ] `.Tests/Phase0` 30/30 (VCF repo)
- [ ] DeNovo recompile + [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) §3
- [ ] Git diff review — no accidental `Design*` left on migrated files unless intentional shim period

---

*Phase 7b — POS migration support. Phase 7c covers `@` dialog → DataTemplate.*
