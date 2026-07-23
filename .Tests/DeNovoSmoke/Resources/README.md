# XAML fixtures

Milestone 1 screens synced from denovo `pos-v1/UI/Resources/XAML/` per [DENOVO_HARNESS_MILESTONE1.md](https://github.com/tdimitriou/denovo/blob/main/docs/migration/DENOVO_HARNESS_MILESTONE1.md) §2.

**Milestone 2:** `Screens/MainMenu/MainMenuView.xml` is **harness-owned** until denovo publishes a MainMenu sync target.  
**Milestone 3:** `Screens/SalesOrder/SalesOrderView.xml` is **harness-owned** (layout-only shell; not a full POS SalesOrder sync).  
**3.14.0:** panel-migrated Login/Sales/StatusBar live under `Migrated/` (no canvas scale); absolute copies remain under `Screens/` / `Widgets/` for comparison.

**VCF 3.0.0+:** fixtures must not use `Design*`. After sync, convert any `DesignLeft`/`DesignTop`/`DesignWidth`/`DesignHeight` to `Margin`/`Width`/`Height` (harness copy already migrated for **3.6.0**).

## Re-sync (preferred)

From denovo `pos-v1`:

```powershell
.\tools\Sync-DeNovoSmokeFixtures.ps1 -DeNovoSmokeRoot 'C:\Users\tdimi\Dev\Projects\Demac\Framework\Demac.VCF\.Tests\DeNovoSmoke'
```

Add `-WhatIf` to preview. The script auto-detects the VCF path when possible; use `-DeNovoSmokeRoot` if your layout differs.

**PNG source:** `pos-v1/UI/Resources/*.png` → `Resources/XAML/Resources/` (matches `ImageKey="Resources\….png"` in Login XAML; harness registers them in `modHarnessImages.bas`).

## Layout

| File | Path under `Resources/XAML/` |
|------|------------------------------|
| SplashView.xml | `Screens/Splash/` |
| LoginView.xml, LoginPad.xml, LoginViewWpf.xml | `Screens/Login/` |
| MainMenuView.xml | `Screens/MainMenu/` (harness m2) |
| SalesOrderView.xml | `Screens/SalesOrder/` (legacy absolute) |
| LoginViewWpf, LoginPad, SalesOrderView, StatusBar | `Migrated/Login/`, `Migrated/SalesOrder/`, `Migrated/Widgets/` (3.14.0) |
| StatusBar.xml | `Widgets/` (legacy absolute) |
| MyApp.xml | root |
| ClockIn.png, Reboot.png, Close.png | `Resources/` |

## Run note

F5 from the VB6 IDE uses `App.Path` = project folder (fixtures load as above). For a compiled `bin\DeNovoSmoke.exe`, copy the whole `Resources\` tree into `bin\Resources\` before running.
