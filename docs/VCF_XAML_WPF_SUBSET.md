# VCF — XAML / WPF subset specification

**Companion:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) · Alignment [§2](./VCF_WPF_ALIGNMENT_NOTES.md)  
**Parser:** `XAMLReader.cls` + vbRichClient `SimpleDOM`  
**Last updated:** 2026-06-19  

---

## 1. Document purpose

Defines **every supported XAML construct** in VCF today, the **target WPF-aligned subset**, and **migration rules** for POS (`pos-v1/UI/Resources/XAML/`).

---

## 2. Document structure (XML)

### 2.1 Root elements

| Root tag | Loader | Phase | Notes |
|----------|--------|-------|-------|
| `<Application>` | `LoadApp` | Keep | Requires `x:Class` matching app type |
| `<Window>` | `Load` / `LoadSuperclassData` | Keep | Modal/modeless via code |
| `<UserControl>` | `Load` / `LoadSuperclassData` | Keep | Primary POS view root |
| `<Style>` | `XAMLStyleReader` | Keep | In resource dictionaries |
| `<DataTemplate>` | Resource load | Evolve | ListView + target ItemsControl |
| Any VCF type | `Load` (no x:Class) | Keep | Fragments, tests |

### 2.2 x:Class two-phase load

**Today:**

```text
1. XAMLReader.Load(xml)
   → If x:Class on root: NewCustomObject(className) ONLY
   → Returns shell; rest of document IGNORED

2. View.InitializeComponent / IWindow.InitializeComponent
   → LoadSuperclassData(Me, xml)  ' full tree into shell
```

**Target:**

- Step 1 **fails loud** if `NewCustomObject` returns Nothing (remove tag-name fallback — B5).
- Optional **`ViewBase`** module with generated `InitializeComponent`.

### 2.3 x:Name

- Registered via `NamingManager.GetNamedChildren`
- Code-behind: `Me.NamedChildren("name")` or view-specific accessors
- **Target:** Same; document lookup failures as `XamlLoadException`

---

## 3. Namespaces

| WPF | VCF today | Target |
|-----|-----------|--------|
| `xmlns="http://schemas.microsoft.com/..."` | **Not supported** | Optional default xmlns map (P3) |
| `xmlns:local="clr-namespace:..."` | **Not supported** | Via TypeRegistry + app registration |
| Default | **`VCF.TypeName`** via `CreateObject("VCF.Button")` | Keep |
| `res:` prefix | POS `ObjectConstructor` loads from flat dict | **Built-in ResourceDictionary** resolver |
| Custom types | `IObjectConstructor.CreateInstance` | **TypeRegistry** |

**Attribute prefix rules:**

- Unprefixed attribute → property on element type
- `x:Class`, `x:Name` → special
- `Grid.ColumnSpan` → attached property (UniformGrid today)

---

## 4. Elements (controls)

### 4.1 Supported today

| Element | Parent | Child content |
|---------|--------|---------------|
| `Window` | Application | Any visual |
| `UserControl` | Window/UC/Panel | Any visual |
| `Panel` | Containers | Multiple |
| `Border` | Containers | Single (informal) |
| `UniformGrid` | Containers | Multiple (cell order) |
| `Button` | Containers | Optional children (TextBlock for text) |
| `TextBlock` | Containers | None |
| `TextBox` | Containers | None |
| `Image` | Containers | None |
| `ListView` | Containers | None (ItemsSource binding) |
| `UnboundListView` | Containers | None (owner-draw) |
| `ScrollBar` | Rare | None |
| `WindowsFormsHost` | Containers | Legacy HWND |
| `Scene` | Window | Visual root |

### 4.2 Target additions

| Element | Phase | WPF reference |
|---------|-------|---------------|
| `Grid` | 2 | RowDefinitions, ColumnDefinitions, * |
| `StackPanel` | 2 | Orientation, Vertical/Horizontal |
| `ContentControl` | 2b | Single Content |
| `ItemsControl` | 4 | ItemsSource + ItemTemplate |
| `DockPanel` | P2 | Dock attached property |
| `ComboBox`, `TabControl` | 5+ | On Selector |

### 4.3 Removed / merged

| Element | Action |
|---------|--------|
| `UnboundListView` | Merge into `ListView` (mode or ItemsSource=Nothing) |

---

## 5. Attributes — layout

### 5.1 Today (VCF dialect)

| Attribute | Type | Applies to |
|-----------|------|------------|
| `DesignLeft`, `DesignTop` | Single (logical coords) | IUIElement |
| `DesignWidth`, `DesignHeight` | Single | IUIElement |
| `HorizontalAlignment`, `VerticalAlignment` | Long enum | TextBlock, some controls |
| `Grid.ColumnSpan`, `Grid.RowSpan` | Long | UniformGrid children |
| `Padding` | Thickness string/object | UniformGrid |
| `Rows`, `Columns` | Long | UniformGrid |
| `Margin` | Thickness | Button (DP) |

**Layout engine:** Parent `MoveChild` scales Design* by `actualWidth/DesignWidth` — **not WPF**.

### 5.2 Target

| Attribute | Replaces | Phase |
|-----------|----------|-------|
| `Width`, `Height` | DesignWidth/Height | 1 |
| `Margin` | DesignLeft/Top (partial) | 1 |
| `HorizontalAlignment`, `VerticalAlignment` | Same names, enum strings | 1 |
| `Grid.Row`, `Grid.Column` | Cell order hack | 2 |
| `Visibility` | Visible bool / missing | 1 |
| `MinWidth`, `MaxHeight`, … | — | 2 |

### 5.3 Legacy root scaling

Until all POS XAML migrated, support **`LayoutMode="ScaleToDesignSurface"`** on root Window/UC (one release) OR document Viewbox equivalent.

---

## 6. Attributes — visual & behavior

### 6.1 Common

| Attribute | Controls | DP? |
|-----------|----------|-----|
| `Name` / `x:Name` | All | Field |
| `DataContext` | All styled | Yes |
| `Style="{StaticResource Key}"` | All styled | Yes |
| `Visibility` | Panel, UC, TextBox, UnboundListView, WFH | Partial |
| `ShowGridLines` | Panel, Border, Window, … | Yes |

### 6.2 Button

| Attribute | Notes |
|-----------|-------|
| `Command`, `CommandParameter` | DPs |
| `ClickMode` | Field |
| `Selected` | DP (inherit) |
| `BorderWidth`, `GradientBackground`, `CornerRadius` | Fields |
| **Target** `Content="Text"` | Phase 5 |

### 6.3 TextBlock / TextBox / Image / ListView

See [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md).

---

## 7. Markup extensions

### 7.1 Supported today

| Extension | Syntax example | Implementation |
|-----------|----------------|----------------|
| **Binding** | `{Binding Path=Prop Mode=TwoWay Converter=MyConv StringFormat={}{0:C}}` | `BindingsManager` → `Binding` |
| **StaticResource** | `{StaticResource ResourceKey=ButtonStyle}` | `StaticResourceExtension` |
| **ThemeResource** | `{ThemeResource Key=BrushPrimary}` | `ThemeResource` → **deprecate** |
| **Self** | `{SelfBinding}` | `SelfBinding` |

**Binding subset:**

| Property | Supported |
|----------|-----------|
| `Path` | Yes (dot paths via NestedProperty) |
| `Mode` | TwoWay, OneWay, OneTime, OneWayToSource |
| `Source` | Explicit object; default = DataContext |
| `Converter` | Type name → CreateInstance |
| `ConverterParameter` | Yes |
| `StringFormat` | Limited |
| `RelativeSource` | **No** |
| `ElementName` | **No** |
| `UpdateSourceTrigger` | **No** (target: Default/LostFocus/PropertyChanged) |

### 7.2 Target

| Extension | Phase | Notes |
|-----------|-------|-------|
| `{DynamicResource Key}` | 3 | Replaces ThemeResource |
| `{StaticResource Key}` | 3 | Faster lookup |
| `{Binding}` | 4 | Via BindingExpression |
| `{x:Null}` | 2 | Explicit null |
| `{x:Type TypeName}` | P3 | Style targets |
| `ResourceReference` | 3 | Clone DataTemplate per cell — §2.16 |

### 7.3 Literal escaping

`MarkupExtensions.ParseLiterals` — brace escaping for nested markup in attributes. **Keep**; document grammar in tests.

---

## 8. Resources

### 8.1 Today (POS pattern)

```xml
<!-- MyApp.xml -->
<Application>
  <Application.Resources>
    <!-- Styles as nested XML or merged keys -->
  </Application.Resources>
</Application>
```

**Plus:** Flat dictionary `MyApp.XAMLResources("Screens\Sales\...")` loaded at startup.  
**Includes:** `<res:Screens\Fragment/>` → `ObjectConstructor.TryCreateObject`.

### 8.2 Target

```xml
<Application.Resources>
  <ResourceDictionary>
    <ResourceDictionary.MergedDictionaries>
      <ResourceDictionary Source="Themes/Dark.xaml"/>
      <ResourceDictionary Source="Screens/Sales/Templates.xaml"/>
    </ResourceDictionary.MergedDictionaries>
    <Style x:Key="ButtonSubmit" TargetType="Button">...</Style>
    <DataTemplate x:Key="MenuItemTemplate">...</DataTemplate>
  </ResourceDictionary>
</Application.Resources>
```

**Keys:** Explicit **`x:Key`** required; one-release filename fallback if missing (§2.16).

### 8.3 Command / cell templates (POS)

| Pattern today | Target |
|---------------|--------|
| `res:Commands\Pay.xml` inline | DataTemplate x:Key + StaticResource |
| Menu grid: clone template per cell | ResourceReference (LoadContent clone) |
| Long-term dense grid | ItemsControl + UniformGrid ItemsPanel |

---

## 9. Styles

### 9.1 Today

```xml
<Style TargetType="Button" Key="ButtonSubmit">
  <Setter Property="BackColor" Value="{ThemeResource Key=PrimaryBrush}"/>
  <Setter Property="BorderColor" Value="..."/>
</Style>
```

- Applied via `StyleManager.ApplyStyle` → `SetCurrentValue` on DPs + **CallByName on CLR props**
- `BasedOn` supported via `XAMLStyleReader`
- Theme change: `Style.ThemesManager_ThemeCkanged` (typo)

### 9.2 Target

- **Setter Property=** must be **registered DP** only (Phase 3)
- `{ThemeResource}` → `{DynamicResource}`
- Named semantic styles (`ButtonSubmit`, `ButtonCancel`) **kept** — not HTML class=
- **Triggers / ControlTemplate** — Phase 6

---

## 10. DataTemplate

### 10.1 Today (ListView only)

- Defined in XAML resources or inline
- `ListView.ItemTemplate` → clone children per row
- Row draw: owner-draw calls template elements' `DrawOn` / widget paint
- Bindings inside template use row item as DataContext

### 10.2 Target

- **ItemsControl** + **ItemTemplate** for menus, message box buttons
- **DataType** keying for implicit templates (P2)
- Drop POS `@`-fragment replacement loops (§2.12 alignment)

---

## 11. Property setting algorithm

### 11.1 Today (`XAMLReader.SetObjectProperties`)

```text
For each attribute:
  1. If registered DP → SetDependencyProperty
  2. Else If SetProperty (CallByName) succeeds → done
  3. Else If IControl.Widget → set widget property
  4. Else silently ignore or Debug.Print
```

### 11.2 Target (Phase 3)

```text
For each attribute:
  1. Resolve property via TypeRegistry (CLR, attached, DP)
  2. If DP → convert string → SetValue
  3. Else If attached → SetAttachedValue
  4. Else → XamlLoadException(element, attribute, reason)
```

**Remove:** steps 2–3 fallback chain (§2.18).

---

## 12. Type creation algorithm

### 12.1 Today

```text
CreateInstance(prefix, name):
  1. If CustomConstructor And prefix=res → CC.CreateInstance("res.path")
  2. If CustomConstructor → CC.CreateInstance(class)
  3. CreateObject(prefix.class)  ' default VCF.*
On Error Resume Next throughout
```

### 12.2 Target

```text
XamlTypeResolver.Resolve(prefix, localName, context):
  1. Built-in VCF types
  2. Registered app types (TypeRegistry)
  3. Resource templates (DataTemplate keys)
  4. Throw XamlLoadException if not found
```

---

## 13. Error handling

| Condition | Today | Target |
|-----------|-------|--------|
| Malformed XML | Exit Function / Sub | XamlLoadException |
| Unknown type | Nothing / Resume Next | XamlLoadException |
| x:Class mismatch | Debug.Print / exit | XamlLoadException |
| Unknown attribute | Silent / CallByName fail | XamlLoadException |
| Binding path error | Debug.Print | XamlLoadException or BindingFailed event |

**XamlLoadException fields:** Message, Line, Column, ElementName, PropertyName, InnerCode

---

## 14. POS migration checklist (per file)

- [ ] Replace `Design*` with `Width`/`Height`/`Margin`
- [ ] Replace `res:...` with merged dictionary + `{StaticResource}`
- [ ] Replace `{ThemeResource}` with `{DynamicResource}`
- [ ] Replace `Visibility="0"` / bool with `Visible`/`Hidden`/`Collapsed`
- [ ] Unify numeric alignments to named enums
- [ ] Move `@`-fragments to DataTemplate keys
- [ ] Update `ObjectConstructor` — remove `res:` cases as framework absorbs
- [ ] Verify bindings after DataContext rebind fix

---

## 15. Golden test documents

Use from `.Tests` and POS subset:

| File | Tests |
|------|-------|
| `.Tests/SampleApp/Resources/XAML/MyApp.xml` | App load, themes |
| `.Tests/Test2/Resources/XAML/SalesOrder.xml` | Bindings, grid |
| `.Tests/Test0/Resources/XAML/DataTemplate1.xml` | Templates |
| `pos-v1/UI/Resources/XAML/MyApp.xml` | Production theme |
| `pos-v1/UI/Resources/XAML/MenuItemsGridButton.xml` | Binding density |

---

## 16. WPF features explicitly out of scope

- 3D, Viewport3D
- FlowDocument, RichTextBox flow
- Animation storyboards (Phase 6+ triggers only)
- MultiBinding, PriorityBinding
- x:Shared=false resource semantics
- Pack URIs — use app-relative paths

---

*Implement parser changes in `XAMLReader` + new `XamlServices` per [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md).*
