# VCF — non-visual infrastructure specification

**Companion:** [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) · Alignment §2.21  
**Last updated:** 2026-06-19  

---

## 1. Scope

Collections, collection views, change notification, resources, value types, static services, and modules — everything in `Demac.VCF.dll` that is **not** a visual control but supports MVVM and XAML.

---

## 2. Design principles (target)

| Principle | Implementation |
|-----------|----------------|
| **WPF parity** | INotifyCollectionChanged, ICollectionView, ResourceDictionary |
| **Lightweight** | No spurious allocations on hot paths |
| **Explicit lifecycle** | Views and bindings detach |
| **Optional split** | Mail/INI/Async → `Demac.VCF.Core` |

---

## 3. Collections

### 3.1 List

**Role:** Internal batch container for event args, tests, temporary groupings.

| API | Notes |
|-----|-------|
| Full IList-like surface | Add, Insert, Remove, Clear, IndexOf, Item, Count, GetEnumerator |
| **Not observable** | |

**Rewrite:** Keep; **forbid** use as `CollectionChangedEventArgs.NewItems` long-term.

---

### 3.2 ObservableCollection

**Implements:** `IEnumerable`, `INotifyCollectionChanged`

**Current issues:**

1. **`CollectionChangedEventArgs`** allocates `New List(item)` on every Add/Remove
2. No batch **`BeginUpdate`/`EndUpdate`**
3. **`GetHashCode`** exposed — unusual; document or remove

**Target API additions:**

```text
BeginUpdate()   ' suppress notifications
EndUpdate()     ' single Reset or batch notification
```

**Target notification:**

```text
' Fast path for single Add:
NewItems = Nothing  ' or single-item readonly wrapper
NewStartingIndex = index
' Avoid New List unless multi-item
```

**Tests:** 1000× Add < baseline ms; memory delta vs current.

---

### 3.3 ObservableDictionary → ResourceDictionary

**Current:** Key-value store with `CollectionChanged` event; used by `ThemesManager`, app resources.

**Target evolution:**

| Feature | Phase |
|---------|-------|
| `Item(key)`, `Add`, `Remove`, `ContainsKey` | Keep |
| **`MergedDictionaries`** collection | 3 |
| **`Source`** URI / path load | 3 |
| **`TryGetResource(key, out value)`** | 3 |
| Theme = swap active merged branch | 3 |

**Threading:** UI thread only (same as today).

---

### 3.4 UIElementCollection

Wraps `ObservableCollection` of child elements; raises `CollectionChanged`; **`Initialize(Parent)`** sets parent on add.

**Target:** Parent hook triggers `FrameworkElement.OnVisualChildAdded/Removed` → InvalidateMeasure.

---

### 3.5 CollectionChangedEventArgs

**Enum `CollectionChangedAction`:** Add, Remove, Replace, Move, Reset

| Field | Target |
|-------|--------|
| `NewItems`, `OldItems` | Optional; lightweight for single-item |
| `NewStartingIndex`, `OldStartingIndex` | Required for Move |

**Factory:** `NewCollectionChangedEventArgs` in `modConstructors` — consider pooling for ListView hot paths.

---

## 4. Collection views

### 4.1 CollectionViewSource

**Singleton** via `StaticClasses.CollectionViewSource`.

| Method | Behavior |
|--------|----------|
| `GetDefaultView(Collection)` | Cache keyed by collection identity |
| **Friend** `DestroyDefaultView` | Test/cleanup |
| **Friend** `GetView` | Internal |

**Target:** Cache invalidation when collection instance replaced.

---

### 4.2 ListCollectionView

**Critical bug B1:**

```vb
' Static flag — WRONG
Private Static bIsInitialized As Boolean
Public Sub Initialize(Source)
    If bIsInitialized Then Exit Sub  ' second view never initializes
```

**Fix:** Per-instance `m_Initialized` flag.

**Bug B2:** `OnSourceCollectionChanged` — no case for `CollectionChangedActionMove`.

**Current surface:**

- Pass-through collection ops on **view** (Add/Remove to source)
- `CurrentItem`, `CurrentPosition`, `MoveCurrentToFirst/Next/Previous/Last`
- **Events:** `CollectionChanged`, `CurrentChanged`

**Target (ICollectionView subset):**

| Feature | Priority |
|---------|----------|
| Move action | P0 |
| SortDescriptions | P2 |
| Filter | P2 |
| `IsCurrentBefore/After` | P3 |
| Sync with Selector.SelectedItem | P0 (Phase 4) |

---

## 5. Change events

### 5.1 PropertyChangedEvent

Manual INPC bridge when source object doesn't expose `PropertyChangedEvent()`:

- `Register` / `Unregister` listener
- `OnPropertyChanged` raises to subscribers
- Uses `ObjPtr` for sender identity

**Target:** Keep for legacy POS objects; prefer native INPC on ViewModels.

---

### 5.2 CollectionChangedEvent

Same pattern for `INotifyCollectionChanged`.

---

### 5.3 DependencyPropertyChanged

**Event on `DependencyProperties`:** `DependencyPropertyChanged(Property, PreviousValue)`

**Target:** Fold into FrameworkElement central hub; coalesce multiple DP changes per frame.

---

## 6. Binding infrastructure

See [VCF_CLASS_REFERENCE.md § Binding](./VCF_CLASS_REFERENCE.md).

**Summary target:**

| Today | Target |
|-------|--------|
| `Binding` + `NestedProperty` | `BindingExpression` |
| `BindingsManager` | `BindingExpression.AttachFromMarkup` |
| Listener-pull on `GetValue` | Push for INPC; pull opt-in legacy |
| Stored in `control.Bindings` List | `BindingExpressionCollection` with Detach |

---

## 7. Markup & value types

| Type | Role | Action |
|------|------|--------|
| **Thickness** | Layout margins/padding | Keep immutable |
| **Color** | HTML/RGB utilities | Keep |
| **SolidColorBrush** | Style setter value | Keep |
| **Variable** | XAML boxed literals | Keep |
| **ArrayWrapper** | Param arrays in XAML | Keep |
| **Function** | ICommand delegate | Keep |
| **CornerRadius** (UDT) | Border/Button | Keep; register as DP type |

---

## 8. Static services

### 8.1 StaticClasses / modStaticClasses

| Singleton | Target location |
|-----------|-----------------|
| `CollectionViewSource` | VCF |
| `DependencyPropertiesStatic` | VCF |
| `BindingsManager` | VCF → merge XamlServices |
| `Application` | VCF |
| `API`, `Object`, `Conversion`, `Color` | VCF |
| `NamingManager` | VCF |
| `INIParser`, `Mail` | **VCF.Core** |
| `StringConversion`, `StringProcessor` | **VCF.Core** |
| `Environment` | VCF or Core |

**Internal:** `Conversion` module friend — keep private.

---

## 9. Modules (complete reference)

### 9.1 modConstructors.bas

| Symbol | Signature | Target |
|--------|-----------|--------|
| `CustomConstructor` | `IObjectConstructor` | → TypeRegistry |
| `NewCollectionChangedEventArgs` | `(Action, NewItems, NewStartingIndex, OldItems, OldStartingIndex)` | Lightweight factory |
| `NewObject` | `(Classname)` | XamlServices |
| `NewCustomObject` | `(Classname)` | TypeRegistry app types |
| `NewList` | `ParamArray Values` | Internal |
| `NewUIElementCollection` | `(Parent)` | Keep |
| `NewDependencyProperties` | `(Target)` | FrameworkElement init |
| `NewThickness` | `ParamArray Params` | Keep |
| `NewBinding` | Full binding signature | BindingExpression.Attach |
| `NewDependencyPropertyMetadata` | AffectsMeasure, AffectsRender, IsInheritable, BindingMode | Registry internal |
| `NewFunction` | Object, Method, CallType, Parameter | Keep |
| `CreateInstance` | Namespace, Class | XamlTypeResolver |
| `NewUIElementBase` | Superclass, Baseclass | FrameworkElement |
| `NewStyle` | TargetType, Key, ParentStyle | Keep |
| `Variable` | Optional Value | Keep |
| `ArrayWrapper` | Optional SourceArray | Keep |

---

### 9.2 modStaticClasses.bas

Module-level `New` singletons — see §8.1.

---

### 9.3 modVisibilityHelper.bas

| Sub | Role |
|-----|------|
| `SetVisibility(Widget, Value)` | Maps bool/enum to widget visible |

**Target:** Merge into `FrameworkElement.Visibility` DP changed callback.

---

### 9.4 modWindowsFormsHostHelper.bas

| Sub | Role |
|-----|------|
| `SetChild`, `ShowWindow` | HWND embedding |

**Target:** Keep unchanged.

---

### 9.5 modInformation.bas / Information.cls

Duplicate **`IsNothing`** helpers — consolidate to one.

---

### 9.6 modUDFConstructors.bas

| Fn | Role |
|----|------|
| `NewCornerRadius` | UDT factory for XAML |

---

### 9.7 modInternalSignals.bas

APC slot array for `BackgroundWorker` thread → UI marshaling.

**Target:** Move with AsyncKit split.

---

### 9.8 modAPI.bas

| Fn | Role |
|----|------|
| `DllPath()` | Path helper |

**Delete duplicate:** `Modules/API.bas` orphan.

---

### 9.9 modStyleWriter.bas

Commented out — remove or designer-only revival.

---

## 10. Application lifecycle

```text
ApplicationStatic.Create(MyApp instance)
  → MyApp.InitializeComponent (LoadApp MyApp.xml)
  → Application.OnInitialized
  → If StartupURI → CreateInstance → Run(window)
  → Cairo message loop
```

**Resources today:** `ObservableDictionary` / nested keys on `Application.Resources`

**Target:** Root `ResourceDictionary` with merged theme + app dictionaries.

---

## 11. ThemesManager

Wraps `ObservableDictionary`; exposes **`ThemeCkanged`** event (typo — fix to **`ThemeChanged`**).

**Active theme swap target:**

```text
Application.Resources.MergedDictionaries(activeIndex) = newThemeDict
  → Style invalidates DynamicResource holders
  → FrameworkElement re-applies Style if needed
```

---

## 12. Utilities split recommendation

| Component | Rationale for split |
|-----------|---------------------|
| **Mail** | CDO dependency; not UI |
| **INIParser** | Config; POS could use KernelLib later |
| **BackgroundWorker** | Threading; optional for UI-only consumers |
| **StringConversion/Processor** | General string ops |

**Build:** Same repo, optional `Demac.VCF.Core.dll` referenced by main DLL — or compile flags.

---

## 13. Performance checklist (P0)

- [ ] Single-item CollectionChanged without List alloc
- [ ] ListCollectionView per-instance init (B1)
- [ ] Move handler (B2)
- [ ] BindingExpression Detach on view dispose
- [ ] ResourceDictionary keyed lookup O(1) — hash map
- [ ] DeferRefresh on bulk ObservableCollection updates

---

## 14. Test matrix

| Test | Component |
|------|-----------|
| Two ListCollectionViews on two collections | ListCollectionView B1 |
| Move item in bound list | Move B2 |
| 1000 Add notifications | ObservableCollection |
| Merged dict lookup | ResourceDictionary |
| Theme swap brush update | ThemesManager + DynamicResource |
| BeginUpdate/EndUpdate single notification | ObservableCollection |

---

*Cross-ref: [VCF_INFRASTRUCTURE.md](./VCF_INFRASTRUCTURE.md) mirrors to VCF repo `doc/` on Phase 0.*
