# VCF — complete class reference (rewrite handoff)

**Companion:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) · [VCF_FRAMEWORK_REWRITE_SPEC.md](./VCF_FRAMEWORK_REWRITE_SPEC.md)  
**Source:** `Demac.VCF\Classes\` — 95 files on disk, **98 vbp registrations**  
**Last updated:** 2026-06-19  

---

## Legend

| Symbol | Meaning |
|--------|---------|
| **Std scaffold** | Controls implementing `IDependencyObject` + `IUIElement` + `IControl` also expose: `DependencyProperties`, `AttachedProperties`, `DataContext`, `DesignLeft/Top/Width/Height`, `Move(...)`, `Parent`, `Bindings`, `Style`, `Children`, `Widget`, `Widgets`, `Name` — omitted below when unchanged |
| **Action** | Rewrite verdict from master spec |
| **Lines** | Approximate LOC |

**Standard control private pattern:** `Class_Initialize` (Register DPs, create widget) → `GetBaseStyle` → `ApplyStyle` → `W_Paint` / `W_HitTest` → `Move` (Design* scale) → `DependencyProperties_DependencyPropertyChanged`.

---

## Core / Dependency Properties

### DependencyObjectBase (42 lines) — **Merge → FrameworkElement**

| Public / Friend | Notes |
|-----------------|-------|
| `DependencyProperties` (Get) | Creates bag in `Class_Initialize` |
| `Parent` (Get/Set) | Pointer-based tree |
| `Children` (Get/Set) | |

**Private:** `Class_Initialize` — `NewDependencyProperties(Me)`  
**Issues:** Minimal base; superseded by FrameworkElement composition.

---

### DependencyProperty (234 lines) — **Refactor**

**Implements:** `INotifyPropertyChanged`

| Member | Signature / type |
|--------|------------------|
| `Name`, `PropertyType`, `PropertyTypeName`, `ProgId`, `UnsetValue`, `Metadata` | Properties (Get) |
| `PropertyChangedEvent()` | → `PropertyChangedEvent` |
| **Friend** `Register(...)` | Static factory per property instance |
| **Friend** `GetValue`, `SetValue`, `SetCurrentValue`, `ClearValue` | |
| **Friend** `AddListener`, `RemoveListener`, `AddCallback`, `RemoveCallback` | Binding + DataContext |
| **Friend** `Parent` (Get) | Owning `DependencyProperties` |

**Private:** `Class_Initialize/Terminate`; `INotifyPropertyChanged_PropertyChangedEvent`  
**Issues:** `Debug.Print` on SetValue error (L143); per-instance property objects — move metadata to registry.

---

### DependencyProperties (101 lines) — **Refactor**

| Member | Notes |
|--------|-------|
| **Event** `DependencyPropertyChanged(Property, PreviousValue)` | |
| `Register(...)`, `GetValue`, `SetValue`, `SetCurrentValue`, `Exists`, `GetProperty` | |
| **Friend** `Target` (Get), `OnDependencyPropertyChanged`, `RegisteredProperties` (Get), `DependencyProperties(Target)` | |

---

### DependencyPropertyMetadata (23 lines) — **Keep**

| Public fields | Default |
|---------------|---------|
| `IsInheritable`, `AffectsRender`, `AffectsMeasure`, `BindingMode` | `BindingMode = OneWay` in Init |

**Target:** Add `PropertyChangedCallback`, `DefaultValue` fields.

---

### DependencyPropertiesStatic (65 lines) — **Refactor**

| Member | Role |
|--------|------|
| `PassPropertyValue(Children, Source)` | Push inheritance to children |
| `InheritPropertyValues(Target)` | Pull from parent |

**Private:** `PassPropertyValueToChild`, `InheritPropertyValue`  
**Target:** Lazy inheritance; batch DataContext propagation (§2.8 alignment).

---

### Thickness (19 lines) — **Keep**

Public fields: `Left`, `Top`, `Right`, `Bottom`. Factory: `NewThickness(ParamArray)`.

---

### UIElementBase (74 lines) — **Merge → FrameworkElement**

| Member | Role |
|--------|------|
| `AttachedProperties`, `Resources` (Get) | |
| `FindResource`, `TryFindResource` | Walk parent → Application |
| **Friend** `Initialize(Superclass)` | Superclass helper |

---

### OverlayWidget (55 lines) — **Remove**

| Fields | Notes |
|--------|-------|
| `Alpha`, `Color`, `CornerRadious`, `Visible` | Typo: Radious |
| `Widget`, `Widgets` (Get) | |

**Private:** `W_ContainerResize`, `W_HitTest` (always false), `W_Paint`  
**Used by:** Button press overlay — replace with template/trigger.

---

### Constructor (164 lines) — **Split → XamlServices**

| Member | Delegates to |
|--------|--------------|
| `SetCustomConstructor`, `GetCustomConstructor` | `modConstructors` |
| `NewCollectionChangedEventArgs`, `NewDependencyProperties`, `NewList`, `NewUIElementCollection` | |
| `NewWindow`, `NewUserControl`, `NewUIElementBase` | View shell wiring |
| `NewBinding`, `NewStyle`, `NewFunction`, `Variable`, `ArrayWrapper` | |
| `CreateInstance(Namespace, Class)` | Type resolution |

**Issues:** God object; `NewWindow` Catch swallows re-raise (L95 commented).

---

### Function (74 lines) — **Keep**

Delegate wrapper: `Object`, `Method`, `CallType`, `Parameter`, `Execute([Parameter])`, `GetHashCode`. Uses `CallByName`.

---

### Variable (93 lines) — **Keep**

Boxed variant for XAML: `Value`, type checks, `Between`, `CType`, `DirectCast`, `TryCast`, `Equals`.

---

## Binding

### Binding (353 lines) — **Replace → BindingExpression**

**Implements:** `IMarkupExtension`, `IDependencyPropertyCallbackListener`

| Member | Notes |
|--------|-------|
| **Enum** `BindingMode` | TwoWay, OneWay, OneTime, OneWayToSource, Default |
| `Mode`, `Source`, `Target`, `TargetProperty`, `Path`, `Converter`, `ConverterParameter`, `StringFormat` | |
| `ProvideValue()` | Markup |
| **Friend** `Initialize(...)`, `SrcDepObj` (Get/Set) | |

**Private:** Source/target `WithEvents`; `SetTargetPropertyValue`; `IDependencyPropertyCallbackListener_OnValueChanged`; `SetSourceProperty`  
**Issues:** `Debug.Print` on errors; triple event graph; no Detach.

---

### BindingsManager (156 lines) — **Refactor**

| Member | Role |
|--------|------|
| `CreateBindingFromMarkup(Target, TargetProperty, MarkupProperties)` | `{Binding Path=… Mode=…}` parser |

**Private:** `CreateBinding`, `ParseMarkupPropertiesString`  
**Note:** `Creatable=False`; singleton via `modStaticClasses`.

---

### NestedProperty (342 lines) — **Remove**

| Member | Role |
|--------|------|
| **Event** `ValueChanged` | |
| `Path`, `Source`, `Name`, `Child`, `Parent` | Property path chain |
| `GetValue`, `SetValue`, `Initialize`, `OnPropertyChanged` | |
| **Friend** `SetVal`, `OnValueChanged` | |

**Private:** `ChildNotifier_PropertyChanged`, `CreateChild`, `SetupChildSource`  
**Target:** Inline in BindingExpression path resolution.

---

### SelfBinding (27 lines) — **Keep**

**Implements:** `IMarkupExtension` — `{SelfBinding}` / `{Self}` → returns host object.

---

### PropertyChangedEvent (59 lines) — **Refactor**

Manual INPC bridge: `Register`, `Unregister`, `OnPropertyChanged`, `IsRegistered`. Uses `ObjPtr` for sender.

---

### ListViewPropertyChangedHandler (36 lines) — **Remove**

**Friend** `Init(Notifier, LV)` → forwards to `ListView.PropertyChangedCallback`. Fold into BindingExpression.

---

## Collections

### List (244 lines) — **Keep (internal)**

**Implements:** `IEnumerable`  
Full list API: `Add`, `AddRange`, `Insert`, `InsertRange`, `Remove`, `RemoveRange`, `RemoveAt`, `Clear`, `Contains`, `IndexOf`, `GetRange`, `ToArray`, `GetEnumerator`, `Count`, `Item`.

**Target:** Internal batches only; not for INCC `NewItems`/`OldItems`.

---

### ObservableCollection (381 lines) — **Refactor**

**Implements:** `IEnumerable`, `INotifyCollectionChanged`  
Same list API + **Event** `CollectionChanged`, `CollectionChangedEvent()`, `GetHashCode`.

**Issues:** Allocates `New List(item)` per change — lightweight args + optional `BeginUpdate`/`EndUpdate`.

---

### ObservableDictionary (286 lines) — **Evolve → ResourceDictionary**

Keyed collection: `Add`, `Remove`, `RemoveKey`, `RemoveAt`, `Clear`, `Contains`, `ContainsKey`, `IndexOf`, `IndexOfKey`, `KeyOfIndex`, `Item`, `ItemAt`, `GetEnumerator`, **Event** `CollectionChanged`.

**Target:** `MergedDictionaries`, `Source`, `TryGetResource`.

---

### UIElementCollection (148 lines) — **Keep**

Wraps inner `ObservableCollection`; **Event** `CollectionChanged`; `Initialize(Parent)`.

---

### CollectionChangedEventArgs (65 lines) — **Refactor**

**Enum** `CollectionChangedAction`: Add, Remove, Replace, Move, Reset  
Props: `Action`, `NewItems`, `NewStartingIndex`, `OldItems`, `OldStartingIndex`  
**Friend** `Initialize`

**Target:** Optional single-item fast path without List allocation.

---

### CollectionChangedEvent (60 lines) — **Keep**

Same pattern as `PropertyChangedEvent` for collection notifications.

---

### CollectionViewSource (45 lines) — **Keep**

`GetDefaultView(Collection)` — caches views by collection hash. **Friend** `DestroyDefaultView`, `GetView`.

---

### ListCollectionView (207 lines) — **Fix + extend**

**Events:** `CollectionChanged`, `CurrentChanged`  
Collection surface + `Source`, `CurrentPosition`, `CurrentItem`, `MoveCurrentTo*`  
**Friend** `OnSourceCollectionChanged`

**Critical bug B1:** Static `bIsInitialized` in `Initialize` — **only first view initializes**.  
**Bug B2:** Move action not handled in source change handler.

---

### DataTemplate (28 lines) — **Evolve**

Fields: `DataType`, `Key`, `Name`. Prop: `Children` → `UIElementCollection`.  
**Target:** Full tree inflation for ItemsControl / ResourceReference clone.

---

## Controls

### Window (552 lines) — **Refactor**

**Implements:** `IDependencyObject`, `IUIElement`, `IControl`

| Beyond scaffold | Notes |
|-----------------|-------|
| **Event** `WindowProc(...)` | Native hook |
| `DialogResult`, `NamedChildren` | |
| `Show([Modal], [Owner], [Focused])`, `ShowDialog`, `Dispose`, `SetFocus` | |
| **Friend** `Initialize`, `OnChildElementsChanged` | |

**DPs registered:** DataContext, ShowGridLines, BorderStyle, Style  
**Non-DP layout:** DesignLeft/Top/Width/Height fields; `MoveChild` scale cascade  
**Issues:** DataContext rebind TODO; Design* Let triggers resize.

---

### Panel (376 lines) — **Refactor**

+ `Visibility` (Get/Let). DPs: DataContext, ShowGridLines, Style.

---

### Border (414 lines) — **Refactor → decorator**

+ `CornerRadius` (CornerRadius UDT). DPs: DataContext, ShowGridLines, CornerRadius, Style.  
**No Visibility wrapper.** Target: single `Child` DP.

---

### Button (712 lines) — **Refactor**

| Beyond scaffold | Storage |
|-----------------|---------|
| `Selected`, `Margin`, `Command`, `CommandParameter`, `ClickMode`, `BorderWidth`, `GradientBackground`, `CornerRadius` | Mix: DPs + fields |

**DPs:** DataContext, Visible, Margin, Command, CommandParameter, Selected, BackColor, BorderColor, ToolTip, Style  
**Private:** `OnClick`, `TmrRelease_Timer`, `ApplyStyle`, `OverlayWidget`  
**Target:** Content DP; remove OverlayWidget; FrameworkElement.

---

### UniformGrid (515 lines) — **Refactor (compat)**

+ `Padding`, `Rows`, `Columns`. DPs: DataContext, ShowGridLines, Visible, Padding, Style.  
**Attached:** `Grid.ColumnSpan`, `Grid.RowSpan` via AttachedProperties dictionary.  
**Target:** Keep for POS migration; add real Grid.

---

### UserControl (391 lines) — **Refactor**

+ `Visibility`. **Friend** `Initialize(Superclass)`. DPs: DataContext, Style.

---

### Scene (366 lines) — **Refactor**

Visual root; DPs: DataContext, Style. No Visibility.

---

### TextBlock (630 lines) — **Refactor**

**Implements:** + `ICloneable`

| Beyond scaffold | |
|-----------------|--|
| `Text`, `ForeColor`, `FontName/Size/Bold/Italic/Underline/StrikeThrough`, `HorizontalAlignment`, `VerticalAlignment`, `ScaleFont` | DPs |
| `Clone()` | Public |

**Private:** `DrawOn`, `W_Paint`, `W_HitTest`, `ApplyStyle`, `ICloneable_Clone`  
**Note:** `DrawOn` is Private; `_TextBlock` orphan had Public `DrawOn` for IVisualChild.

---

### Image (328 lines) — **Refactor**

+ `ImageKey`, `KeepAspectRatio`. DPs: DataContext only.  
**Private:** `W_Paint`, `W_HitTest`

---

### TextBoxBase (1573 lines) — **Refactor (engine)**

**Not** a VCF control interface — Cairo/vbWidgets text engine.

**Events:** Click, Scroll, Change, OwnerDrawBackGround, KeyDown/Press/Up, Validate, MaxLengthViolation, BeforePaste, BeforeSelChange, SelChanged

**Public properties (complete):**

| Property | Type |
|----------|------|
| `Widget`, `Widgets` | cWidgetBase, cWidgets |
| `Text` | String |
| `CueBannerText` | String |
| `Locked`, `UndoDepth`, `InnerSpace` | |
| `Alignment`, `VCenter`, `PasswordChar`, `MaxLength` | |
| `SelStart`, `SelLength`, `SelText` | |
| `AutoSelectAll`, `HideSelection`, `Border` | |
| `TextShadowOffsetX/Y`, `TextShadowColor` | |
| `MultiLine`, `ScrollBars` | |
| `TopRow`, `RowHeight` | |
| `SelectAll()`, `EnsureVisible()` | |
| `RowCount()`, `VisibleRows()` | |

**Private (key):** Caret, selection, scrollbars, keyboard/mouse, `CalcRows`, `CalcCoords`, `SetSelection`, `AdjustDimensions`, full paint pipeline.

**Target:** Keep widget core; expose via TextBox DPs only; no duplicate public surface.

---

### TextBox (486 lines) — **Refactor (thin wrapper)**

**Implements:** `IDependencyObject`, `IUIElement`, `IControl`  
Forwards TextBoxBase events. + `Border`, `Focused`, `Visibility`, `SetFocus`.

**DPs:** DataContext, Text (TwoWay), Alignment, VCenter, PasswordChar, CueBanner, Style  
**Private:** `m_Base_*` event forwarders; `DependencyProperties_DependencyPropertyChanged`

---

### ScrollBar (265 lines) — **Keep**

**Events:** Change, Scroll, MouseWheel  
**Props:** `Vertical`, `Value`, `Min`, `Max`, `SmallChange`, `LargeChange`, `BottomRightEdgeWidth`, `Widget`, `Widgets`  
**Note:** Setting `.Value` in code does not fire Change (intentional, L34).

---

### ListViewBase (1178 lines) — **Refactor (engine)**

**Events:** OwnerDrawItem, OwnerDrawSubItem, OwnerDrawHeader, Click, HeaderClick, DblClick, DeleteKeyPressed, Scroll*, Mouse*, SelectedAll, DimensionsAdjusted, ListIndexChanged

**Public properties (complete):**

| Group | Members |
|-------|---------|
| Widget | `Widget`, `Widgets` |
| Scroll / index | `VisibleRows`, `ListCount`, `ScrollIndex`, `ListIndex`, `HoverIndex`, `ScrollerSize` |
| Selection | `MultiSelect`, `Selected(Index)`, `ClearSelections`, `GetSelections`, `ShowSelection`, `DeselectOutsideClick` |
| Layout | `RowHeight`, `ColWidth`, `HeaderHeight`, `RowSelectorWidth`, `ColumnIndex`, `ColumnCount`, `ColumnDefaultWidth`, `ColumnWidth(Idx)`, `ColMapIndex`, `VisibleCols`, `DrawWidth`, `DrawHeight` |
| Behavior | `AllowDrag`, `AllowColResize`, `AllowRowResize`, `ShowHoverBar`, `ImageView` |
| Colors | `RowColor`, `AlternateRowColor` |
| Methods | `EnsureVisibleSelection`, `GetListIndexFromMouseY/XY`, `GetColumnIndexFromMouseX`, `ResetSortStates`, `MoveColumnToNewIndex`, `AdjustDimensions` |

**Private (key):** Full Cairo list/grid paint, header resize, keyboard nav, sort state, scroll math.

**Target:** Variable row height (`MeasureRow`), column model, hierarchy for InvoiceGrid.

---

### ListView (755 lines) — **Merge → Selector**

**Implements:** + `IItemsControl`

| Beyond scaffold | Notes |
|-----------------|-------|
| `Base` (ListViewBase), `Resources`, `Background`, `SelectedBackground/Foreground`, `RowHeight`, `ItemsSource`, `ItemTemplate`, `Name` | |

**DPs:** DataContext, ItemsSource, Style  
**Private:** Template clone/cache, `FindDataTemplate`, `CreateDataTemplate`, bound draw, ItemsSource change  
**Bugs:** B3 stubs, B4 ObservableCollection-only, B2 Move, B6 DataContext rebind

---

### UnboundListView (448 lines) — **Remove (merge)**

Same visual API minus ItemsSource; + `Refresh()`. Delegates to `Base`.  
14 forwarded events from ListViewBase family.

---

### WindowsFormsHost (391 lines) — **Keep**

+ `AutomaticallyUnloadContent`. **`Children` → `ObservableCollection`** (not UIElementCollection).  
Hosts legacy VB6 Form/Control in Cairo tree.

---

### Orphan: _Image.cls (292 lines) — **Delete**

`Image_OLD`; **Implements** `IVisualChild`; Public `DrawOn`.

---

### Orphan: _TextBlock.cls (529 lines) — **Delete**

`TextBlock_BAK`; Public `DrawOn`; missing Style/Children parity.

---

## XAML

### XAMLReader (581 lines) — **Refactor**

| Public | Behavior |
|--------|----------|
| `Load(XML)` | x:Class shell or NewObject(root) |
| `LoadSuperclassData(Superclass, XML)` | Window/UC body |
| `LoadApp(Superclass, XML)` | Application root |

**Private (key):** `NewObject`, `CreateInstance`, `SetObjectProperties`, `SetDependencyProperties`, `SetProperty` (CallByName fallback), `SetObjData`, `LoadResources`, `LoadAppData`

**Issues B5:** Silent exit on malformed XML; dangerous x:Class fallback; `Debug.Print "Invalid Type"`.

---

### XAMLStyleReader (87 lines) — **Keep**

`LoadStyle(ResourceDictionary, Node)` → `CreateStyle` with Setter nodes.

---

### MakupExtensions.cls (217 lines) — **Refactor**

**VB_Name = MarkupExtensions** (filename typo).  
`RestoreLiterals`, `ParseLiterals`, `GetMarkupValue`, `SetProperties`, `GetExtensionValue`, `CreateExtObject`.

**Orphan duplicate:** `MarkupExtensions.cls` (193 lines) — **delete**.

---

### XAMLDependencyPropertyManager (72 lines) — **Merge**

`GetPropertyValueFromString(Prop, Value)` — UDF types, object refs via CreateInstance.

---

### XAMLImagePropertyManager (44 lines) — **Keep / fix B11**

`LoadImage(Path)` — file load works; **`LoadImageFromResource` empty stub**.

---

### XAMLThicknessConstructor (36 lines) — **Keep**

`NewThickness([Args As String])` — comma-separated parse.

---

### StaticResourceExtension (180 lines) — **Keep**

**Implements:** `IMarkupExtension`  
`ResourceKey`, `Target`, `ProvideValue`, `InitializeFromMarkup`  
**Creatable=False**

---

### ThemeResource (63 lines) — **Deprecate**

**Implements:** `IMarkupExtension` — `{ThemeResource Key=…}` → shim to DynamicResource.

---

## Style / Theme

### Style (254 lines) — **Keep**

**Implements:** `IEnumerable`  
**Event:** `StyleChanged`  
`Initialize`, `GetSetter`, `SetSetter`, `RemoveSetter`, `Clear`, `TargetType`, `Key`, `BasedOn`, `Count`, `MarkupValue`  
**Friend:** `AddChild`, `RemoveChild`, `BasedOnStyleChanged`  
**Issues:** `ThemesManager_ThemeCkanged` handler name typo (B7).

---

### StyleManager (153 lines) — **Refactor**

`ApplyStyle(Style, Target)` → `SetDependencyProperties` + **`SetProperty` CallByName fallback** (remove in Phase 3).

---

### Setter (26 lines) — **Keep**

`Property` (Get/Let), `Value` field.

---

### SolidColorBrush (25 lines) — **Keep**

`Color` (Get/Let string → resolved Long via `Color.FromString`).

---

### ThemesManager (199 lines) — **Refactor**

**Implements:** `ObservableDictionary` delegation  
**Events:** `CollectionChanged`, **`ThemeCkanged`** (typo B7)  
`ActiveThemeName`, `ActiveTheme` + full dict API.

---

## Application

### Application (132 lines) — **Refactor**

**Event:** `Startup`  
`Windows`, `StartupURI`, `Resources`, `Run([Window])`, `FindResource`, `TryFindResource`  
**Friend:** `ThemesManager`, `Initialize`, `OnInitialized`, `OnStartup`  
**Target:** Resources = root ResourceDictionary.

---

### ApplicationStatic (52 lines) — **Keep**

`Current` (Get/Set), `Create(Superclass)` — singleton via ObjPtr.

---

## Utilities

### API (90 lines) — **Keep**

`ObjFromPtr`, `CopyVariable`, `CObj`, `CUnk`, `SetFocus`, `GetFocus`, `ArrayExists`, `pArrPtr`, `ArrayFind`

---

### StaticClasses (58 lines) — **Refactor**

Service locator: `CollectionViewSource`, `Application`, `API`, `INIParser`, `Mail`, `Object`, `StringConversion`, `NamingManager`, `StringProcessor`, `Color`, `Environment`

---

### ObjectStatic (194 lines) — **Keep**

`Equals`, `Greater`, `Less`, `EqualOrGreater`, `EqualOrLess`, `Between`

---

### Conversion (118 lines) — **Keep**

`CType`, `TryCast`, `DirectCast`, `Cast`, `CObj`, `CUnk`

---

### Color (145 lines) — **Keep**

`FromString`, `FromSystemColor`, `Invert`, `Multiply`, `FromHtml`, `ToHtml`, `ToRGB`, `FromRGB`, `Luminance`

---

### StringProcessor (65 lines) — **Split optional**

`Parse`, `Format`, `Split`

---

### StringConversion (87 lines) — **Split optional**

`LocaleID`, `ToProperCase`, `ToLowerCase`, `ToUpperCase` — Greek locale accent handling.

---

### Information (28 lines) — **Keep**

`IsNothing`, `IFNull` (typo in param name)

---

### INIParser (48 lines) — **Split → VCF.Core**

`ReadSetting`, `WriteSetting` — Win32 profile API.

---

### Mail (80 lines) — **Split → VCF.Core**

`Send(...)` — CDO; typos: Recepient, Attchments.

---

### ArrayWrapper (158 lines) — **Keep**

Array utilities for XAML/literals: `Initialize`, bounds, `Find`, `Join`, `EraseArray`.

---

### NamingManager (63 lines) — **Keep**

`GetNamedChildren(Parent)` — recursive x:Name registry.

---

### Environment (20 lines) — **Split optional**

`TickCount` — `GetTickCount` API.

---

### Interaction (18 lines) — **Keep**

`SendKeys(Text, [Wait])` — VBA wrapper.

---

## Interfaces (complete contracts)

| Interface | Members | Rewrite notes |
|-----------|---------|---------------|
| **IDependencyObject** | `DependencyProperties`, `Parent`, `Children` | + `ClearValue` |
| **IUIElement** | Design rect, `DataContext`, `Base`, `AttachedProperties`, `Parent`, `Move(...)` | Design* → layout DPs; ActualWidth/Height |
| **IControl** | `CornerRadius` type, `ClickMode`/`Visibility` enums, `Widget`, `Widgets`, `Children` | Widget → Friend |
| **IWindow** | `Base`, `InitializeComponent` | |
| **IUserControl** | Full UC + **`Move(ByRef ...)`** | Unify Move signature (B13) |
| **IApplication** | Resources, Base, StartupURI, Run, InitializeComponent, SetBase, Find/TryFindResource | |
| **IItemsControl** | `ItemsSource`, `ItemTemplate` | + ItemsPanel, ItemContainerStyle |
| **IVisualChild** | `DrawOn(CC, [ForeColor])` | Internal ListView path |
| **ICommand** | `Execute`, `CanExecute` | + `CanExecuteChanged` event |
| **IValueConverter** | `Convert`, `ConvertBack` | |
| **INotifyPropertyChanged** | `PropertyChangedEvent()` | |
| **INotifyCollectionChanged** | `CollectionChangedEvent()` | |
| **IMarkupExtension** | `ProvideValue()` | |
| **IObjectConstructor** | `CreateInstance(Classname)` | → TypeRegistry |
| **ICloneable** | `Clone()` | |
| **IDependencyPropertyCallbackListener** | `OnValueRequested(ByRef)`, `OnValueChanged` | Design doc in `IDependencyPropertyCallback.cls` (69 lines) |

**Orphan stub:** `IDependencyPropertyCallbackListener.cls` (16 lines) — wrong signature; **delete**.

---

## Async

### BackgroundWorker (235 lines) — **Split optional**

**Events:** ProgressChanged, WorkerEvent, RunWorkerCompleted  
`RunWorkerAsync(Task, [Args])`, Pause, ResumeWorker, CancelAsync  
Uses ThreadHost + modInternalSignals APC slots.

---

### InternalWorker (76 lines) — **Split**

**Event:** WorkerEvent  
`DoWork`, `CheckAndWaitIfPaused`, `ReportProgress`, `RaiseEventAsync`, `CancellationPending`

---

### IBackgroundTask (16 lines) — **Split**

`Execute(Bridge, ByRef Args()) As Variant`

---

### ErrorInfo (40 lines) — **Split**

`Number`, `Description`, `Source`; **Friend** `Populate`

---

## Modules (complete)

See [VCF_INFRASTRUCTURE.md § Modules](./VCF_INFRASTRUCTURE.md).

---

*End of class reference. Property details: [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md).*
