# XAML fixtures (milestone 1)

Synced from denovo `pos-v1/UI/Resources/XAML/` per [DENOVO_HARNESS_MILESTONE1.md](https://github.com/tdimitriou/denovo/blob/main/docs/migration/DENOVO_HARNESS_MILESTONE1.md) §2.

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
| LoginView.xml, LoginPad.xml | `Screens/Login/` |
| StatusBar.xml | `Widgets/` |
| MyApp.xml | root |
| ClockIn.png, Reboot.png, Close.png | `Resources/` |

## Run note

F5 from the VB6 IDE uses `App.Path` = project folder (fixtures load as above). For a compiled `bin\DeNovoSmoke.exe`, copy the whole `Resources\` tree into `bin\Resources\` before running.
