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

**Pin for harness work:** **`v2.22.0-wpf-alignment-p8b-denovo-m2`** (Phase 8b lazy inherit + DeNovoSmoke milestone 2; requires **2.18.0+** layout shim)

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
| **LoginViewWpf.xml** (7e) | WPF **Grid** layout — inner panel scales on resize; toggle via `modHarnessConfig.USE_WPF_LOGIN_LAYOUT` |

**Run:** F5 in VB6 IDE, &lt; 30 s, no database.

**DeNovo delivers:** migrated XAML files, required `res:` fragments (or inlined equivalents), stub VM property/command list — see **`DENOVO_HARNESS_MILESTONE1.md`** in denovo repo.

#### Fixture sync (from denovo `pos-v1/`)

Copy into `Resources/XAML/` (preserve `res:` paths):

| File | Source (denovo) |
|------|-----------------|
| `SplashView.xml` | `UI/Resources/XAML/Screens/Splash/SplashView.xml` |
| `LoginView.xml` | `UI/Resources/XAML/Screens/Login/LoginView.xml` |
| `LoginViewWpf.xml` | VCF harness only — WPF Grid reference (not synced from denovo) |
| `LoginPad.xml` | `UI/Resources/XAML/Screens/Login/LoginPad.xml` |
| `StatusBar.xml` | `UI/Resources/XAML/Widgets/StatusBar.xml` |
| `MyApp.xml` | `UI/Resources/XAML/MyApp.xml` (styles / theme slice) |
| PNG assets | `Resources/ClockIn.png`, `Reboot.png`, `Close.png` (under `XAML/Resources/` after sync) |

#### Stub VM contracts (milestone 1)

| Context | Members |
|---------|---------|
| **SplashViewModel** | `Progress`, `InfoMessage` (strings) |
| **LoginViewModel** | `UserList` (`ObservableCollection`), `PasswordText`, `LoginCommand`, `ClearInput` |
| **MainMenuViewModel** | `Title`, `Subtitle`, `WelcomeMessage`, `SalesCommand` (m3 stub), `LogoutCommand` → Login |
| **AppCommands** (static) | `KeyClick`, `ClockInOutCommand`, `RestartCommand`, `ShutdownCommand` |
| **LanguageManager** (static) | Keys: `Loading`, `Clear`, `Login`, `Restart`, `Exit` |
| **AppProperties** (static) | `TerminalID`, `CurrentUser`, `CurrentTime` (StatusBar) |

**Harness-only:** assign `MyList.ItemsSource = UserList` after Login parse (production does this in code-behind). Navigation: visibility swap + `RelayoutChildren` + `RebuildNamedItemsList` — no Friend APIs.

### Milestone 2

| Screen / task | Assert |
|---------------|--------|
| **MainMenuView.xml** | Login → MainMenu (Login button / **Shift+M**); Logout → Login; named `MenuTitle` / `WelcomeText` after `RebuildNamedItemsList` |
| **Lazy Login load** (7f) | Login loads on first `ShowLogin` only (`EAGER_LOGIN_LOAD = False`) — production-like startup |
| **Lazy MainMenu load** | MainMenu loads on first `ShowMainMenu` (`EAGER_MAINMENU_LOAD = False`) |
| **P7d-LOAD-*** gates | Immediate: `[P7d-LOAD-SPLASH]`, `[P7d-LOAD-LOGIN]`, `[P7d-LOAD-MAINMENU]`, `[P7d-LOAD-BORDERED]` |

**7f / m2 config** (`modHarnessConfig.bas`):

| Const | Default | Meaning |
|-------|---------|---------|
| `EAGER_LOGIN_LOAD` | `False` | `True` = m1-style preload Login |
| `EAGER_MAINMENU_LOAD` | `False` | `True` = preload MainMenu at startup |
| `ENABLE_LOAD_BENCH` | `True` | Log load ms to Immediate window |

**MainMenu fixture:** harness-owned `Resources/XAML/Screens/MainMenu/MainMenuView.xml` + `StubMainMenuViewModel` (until denovo syncs production MainMenu). Sales button is a stub for milestone 3.

### Keyboard shortcuts (harness)

| Key | Action |
|-----|--------|
| **Shift+L** | Skip splash → Login |
| **Shift+M** | Show MainMenu (lazy-loads XAML; also reached via Login button) |
| **Shift+B** | Open bordered `BorderTestWindow`; check Immediate window for `[BORDER-DIAG]` |
| **Shift+1** | Resize shell to **1024×768** (baseline) |
| **Shift+2** | Resize shell to **800×600** |
| **Shift+3** | Resize shell to **1366×768** (widescreen) |
| **Exit** | `Shutdown` → `RemoveAll` |

Borderless shell cannot be drag-resized (matches production POS). Use **Shift+1/2/3** for Test 5 — presets change **client** size only; host design canvas stays **1024×768** so legacy scale matches drag-resize. Or set `USE_SIZABLE_SHELL_BORDER = True` for manual edge drag (adds title bar temporarily).

See [WINDOW_LIFECYCLE.md](../../docs/WINDOW_LIFECYCLE.md).

### Milestone 3

| Screen | Assert |
|--------|--------|
| **Sales order (layout-only)** | UniformGrid / columns render; stub line collection; no DB commit |

---

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

**Shutdown (smoke):** `AppCommands.ShutdownCommand` → `modHarnessAppManager.Shutdown` — mirrors DeNovo `AppManager.Shutdown` for Cairo widget forms only (`RemoveAll` ends `EnterMessageLoop`). `modApp.ResetSession` runs after `Run` returns (IDE second-F5 globals). Do not use `End` from command handlers.

**Layout shim (load time):** `ApplyLegacyLayoutProperty` in XAMLReader — not a consumer API. See [POS_RUNTIME_FEEDBACK.md](../../docs/POS_RUNTIME_FEEDBACK.md).

**Load performance:** Milestone 1 preloads Splash + Login before show. Phase **7f** adds lazy Login + timed benchmarks — [VCF_PERFORMANCE_BENCHMARKS.md](../../docs/VCF_PERFORMANCE_BENCHMARKS.md) §7f.

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
5. **F5** — Splash ~2.5 s, then Login. **Shift+L** skips to Login early. **Shift+B** opens a bordered test dialog (`BorderTestWindow.xml`, `BorderStyle=2`) to verify window chrome lifecycle — open **View → Immediate Window** (`Ctrl+G`) and look for `[BORDER-DIAG]` (toggle `ENABLE_BORDER_CHROME_DIAG` in `modHarnessConfig.bas`). **Exit** → `modHarnessAppManager.Shutdown` (stop timers + `Cairo.WidgetForms.RemoveAll`, same core as DeNovo `AppManager.Shutdown` without the VB6 `Forms` loop).

**Pass (milestone 1):** no XAML load error; Splash layout + bindings; Login list + password pad + focus; Exit ends `Application.Run` cleanly in IDE and compiled exe.

### Resize validation (7e — `LoginViewWpf.xml`)

**Hybrid layout:** outer shell (branding, `InnerBorder` frame, side buttons, status bar) uses legacy **`Design*`** on the `UserControl` — same positions as production `LoginView.xml`. Only **inside** `InnerBorder` uses a WPF **Grid** (list | password + pad) for resize-friendly content.

With `USE_WPF_LOGIN_LAYOUT = True`, go to Login (**Shift+L**), then press **Shift+1**, **Shift+2**, **Shift+3**. Confirm the **star** column (password + numpad) grows on widescreen; list column stays ~204px (by design). Check Immediate for `[HARNESS-RESIZE]` lines.

With `USE_WPF_LOGIN_LAYOUT = False` (legacy `LoginView.xml`), the same presets should scale **inner** Design* content inside `InnerBorder` via **Option B** (`Border` auto-enables `LegacyScaleLayout` when children use Design*) — see [docs/POS_LAYOUT_RESIZE.md](../../docs/POS_LAYOUT_RESIZE.md).

| Size | Check |
|------|--------|
| 1024×768 | Design baseline |
| 800×600 | Inner panel scales; no clipped numpad |
| Widescreen | List + pad share horizontal space |

**Out of scope (legacy LoginView):** nested Border + `Design*` does not scale — use WPF layout for production migration.
