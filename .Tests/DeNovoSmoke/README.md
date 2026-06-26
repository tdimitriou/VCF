# DeNovo UI harness (DeNovoSmoke)

Minimal VB6 exe to load **DeNovo POS XAML screens** against a pinned **`Demac.VCF.dll`** — without KernelLib, Demac.Data, DB, or hardware.

**Status:** VB6 project checked in — open `DeNovoSmoke.vbp` in IDE (build `Demac.VCF.dll` first). Borderless shell **1024×768** client (`BorderStyle=0`).
**VCF response:** [docs/DENOVO_HARNESS_PROPOSAL_RESPONSE.md](../../docs/DENOVO_HARNESS_PROPOSAL_RESPONSE.md)  
**DeNovo proposal:** denovo monorepo → `docs/migration/VCF_TEAM_HANDOFF_HARNESS_PROPOSAL.md`  
**Milestone 1 (fixtures + stub VMs):** denovo monorepo → `docs/migration/DENOVO_HARNESS_MILESTONE1.md` — **aligned 2026-06-22**

---

## Why this exists

Full `DeNovo.exe` is a deep COM graph (KernelLib → Data → ADODBUtils → UILib → VCF). With Phase 1 binary/typelib churn, failures on Login XAML are indistinguishable from `Data.Dataset` or stale group `.vbg` references.

**Phase 1 acceptance:** Phase0 **31/31** + this harness green for agreed screens with **stub view models**.  
**Phase 2:** re-attach full POS stack; run [POS_INTEGRATION_SMOKE.md](../../docs/POS_INTEGRATION_SMOKE.md) §3.

---

## Ownership split

| Owner | Location | Delivers |
|-------|----------|----------|
| **VCF** | `.Tests/DeNovoSmoke/` | Runner exe, shell, navigation, automated checks |
| **DeNovo** | `denovo/pos-v1/` | Production XAML fixtures + stub VMs (synced into `Resources/XAML/`) |

VCF does **not** duplicate the full POS XAML tree. DeNovo does **not** own the harness runner (keeps Phase0 authoritative on every tag).

---

## References (mandatory)

| DLL | Role |
|-----|------|
| **`Demac.VCF.dll`** | Framework under test |
| **`vbRichClient5.dll`** | Cairo / widget host (mandatory RC dependency) |

**No references:** KernelLib, Demac.Data, UILib, ADODBUtils, POSWidgets (except types registered via minimal local `ObjectConstructor` stubs if required for `res:`).

**Pin for harness work:** **`v2.19.0-wpf-alignment-p7d-denovo-smoke`** (requires **2.18.0+** layout shim)

---

## Planned project layout

```text
DeNovoSmoke.vbp              StdExe, Startup = Sub Main
├── Modules/
│   └── modMain.bas            Cairo message loop, create shell
├── Classes/
│   ├── AppHost.cls            minimal IApplication / theme loader
│   ├── ShellWindow.cls        VCF Window root (pattern from .Tests/Test0)
│   ├── HarnessNavigation.cls  swap active IUserControl (visibility-based)
│   └── Stubs/
│       ├── StubLoginViewModel.cls
│       ├── StubSplashViewModel.cls
│       └── StubMainMenuViewModel.cls
└── Resources/
    └── XAML/                  DeNovo screen subset (copied from migration output)
```

---

## Milestones

### Milestone 1 (target: first tagged harness drop)

| Screen | Assert |
|--------|--------|
| **SplashView.xml** | Parses; layout; theme; no XAML load error |
| **LoginView.xml** | Parses; stub `DataContext`; bindings resolve; focus OK |

**Run:** F5 in VB6 IDE, &lt; 30 s, no database.

**DeNovo delivers:** migrated XAML files, required `res:` fragments (or inlined equivalents), stub VM property/command list — see **`DENOVO_HARNESS_MILESTONE1.md`** in denovo repo.

#### Fixture sync (from denovo `pos-v1/`)

Copy into `Resources/XAML/` (preserve `res:` paths):

| File | Source (denovo) |
|------|-----------------|
| `SplashView.xml` | `UI/Resources/XAML/Screens/Splash/SplashView.xml` |
| `LoginView.xml` | `UI/Resources/XAML/Screens/Login/LoginView.xml` |
| `LoginPad.xml` | `UI/Resources/XAML/Screens/Login/LoginPad.xml` |
| `StatusBar.xml` | `UI/Resources/XAML/Widgets/StatusBar.xml` |
| `MyApp.xml` | `UI/Resources/XAML/MyApp.xml` (styles / theme slice) |
| PNG assets | `Resources/ClockIn.png`, `Reboot.png`, `Close.png` (under `XAML/Resources/` after sync) |

#### Stub VM contracts (milestone 1)

| Context | Members |
|---------|---------|
| **SplashViewModel** | `Progress`, `InfoMessage` (strings) |
| **LoginViewModel** | `UserList` (`ObservableCollection`), `PasswordText`, `LoginCommand`, `ClearInput` |
| **AppCommands** (static) | `KeyClick`, `ClockInOutCommand`, `RestartCommand`, `ShutdownCommand` |
| **LanguageManager** (static) | Keys: `Loading`, `Clear`, `Login`, `Restart`, `Exit` |
| **AppProperties** (static) | `TerminalID`, `CurrentUser`, `CurrentTime` (StatusBar) |

**Harness-only:** assign `MyList.ItemsSource = UserList` after Login parse (production does this in code-behind). Navigation: visibility swap + `RelayoutChildren` + `RebuildNamedItemsList` — no Friend APIs.

### Milestone 2

| Screen | Assert |
|--------|--------|
| **MainMenuView.xml** | Navigation from Login; named-element lookup after `RebuildNamedItemsList` if needed |

### Milestone 3

| Screen | Assert |
|--------|--------|
| **Sales order (layout-only)** | UniformGrid / columns render; stub line collection; no DB commit |

---

## Supported host-app APIs (2.18.x)

Documented for visibility-based navigation — **do not add new public methods from DeNovo**:

| Method | Use |
|--------|-----|
| **`Window.RelayoutChildren`** | After showing/collapsing views |
| **`Window.RebuildNamedItemsList`** | After navigation when named lookup breaks |
| **`Window.ApplyDeferredChildLayout`** | Flush layout deferred during `LockRefresh` |

**Do not call:** `UserControl.ApplyDeferredHostLayout` (Friend). **`RelayoutChildren`** invokes it internally.

**Layout shim (load time):** `ApplyLegacyLayoutProperty` in XAMLReader — not a consumer API. See [POS_RUNTIME_FEEDBACK.md](../../docs/POS_RUNTIME_FEEDBACK.md).

---

## Issue reporting

When filing upstream:

1. **Tag pinned** (e.g. `v2.18.0-wpf-alignment-p7c-layout-shim`)
2. **Screen name** (e.g. `LoginView.xml`)
3. **XAML fragment** (minimal repro)
4. **Stub VM** properties bound in that fragment
5. **Expected vs actual** (layout coords, binding, theme)

---

## Relation to Phase0

| Suite | Scope |
|-------|--------|
| **Phase0** | Framework unit/regression benchmarks (31 tests); runs on every tag |
| **DeNovoSmoke** | Real POS XAML + stub VMs; consumer contract validation |

Both must pass before Phase 7 UI tags that affect DeNovo integration.

---

## Run (VB6 IDE)

1. Build/register **`Demac.VCF.dll`** from repo root (pin **`v2.18.0-wpf-alignment-p7c-layout-shim`** or later).
2. Open **`.Tests/DeNovoSmoke/DeNovoSmoke.vbp`**.
3. Fix reference paths if needed (`vbRichClient5.dll`, `Demac.VCF.dll` — same layout as Phase0).
4. Sync fixtures when denovo XAML changes — [Resources/README.md](Resources/README.md) (`Sync-DeNovoSmokeFixtures.ps1`).
5. **F5** — Splash ~2.5 s, then Login. **Shift+L** skips to Login early.

**Pass (milestone 1):** no XAML load error; Splash layout + bindings; Login list + password pad + focus.

**Out of scope (milestone 1):** multi-resolution / nested Border resize — see [docs/POS_LAYOUT_RESIZE.md](../../docs/POS_LAYOUT_RESIZE.md). Production POS is borderless; XAML design canvas is 1024×768 but runtime supports 800×600, widescreen, etc.
