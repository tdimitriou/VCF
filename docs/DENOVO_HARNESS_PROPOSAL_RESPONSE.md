# VCF → DeNovo — harness proposal response

**To:** DeNovo / POS maintainers (`denovo` monorepo)  
**From:** Demac.VCF team  
**Date:** 2026-06-22  
**Re:** Response to integration status & proposed UI harness  
**Status:** **Aligned — proceed with harness-first validation**

**DeNovo proposal (consumer):** denovo monorepo → `docs/migration/VCF_TEAM_HANDOFF_HARNESS_PROPOSAL.md`  
**Harness scaffold:** [`.Tests/DeNovoSmoke/README.md`](../.Tests/DeNovoSmoke/README.md)

---

## 1. Summary

Thank you for pausing integration work and sending a full inventory before more changes. We **agree** with your analysis:

- Full `DeNovo.exe` is the wrong **primary** acceptance vehicle for Phase 1 XAML/VCF work.
- A **minimal harness** (VCF + vbRichClient5 only) is the right next step.
- **Consumers must not extend the VCF typelib** with test hooks or ad-hoc public methods — we share that boundary.

Below: answers to your §6 questions, reconciliation of §2, and VCF-side actions.

---

## 2. Answers to your questions (§6)

### 2.1 Ownership — who builds the harness?

**Split ownership (agreed):**

| Owner | Location | Responsibility |
|-------|----------|----------------|
| **VCF** | `.Tests/DeNovoSmoke/` (next to Phase0) | Harness **runner**: minimal `StdExe`, shell window, navigation stub, automated checks where feasible |
| **DeNovo** | `denovo/pos-v1/` | **XAML corpus** + **stub view models** matching real binding surfaces; sync into harness `Resources/XAML/` |

**Why not DeNovo-only:** Phase0 must stay the authoritative gate on every VCF tag. Putting the runner in VCF keeps regression next to the DLL and avoids typelib skew from KernelLib/Data.

**Why not VCF-only XAML:** Your production tree is the source of truth for `res:` paths, theme keys, and VM binding contracts. We will not duplicate 100+ screens in VCF — we take **fixtures you contribute** (starting with Splash + Login).

**First milestone (both sides):** Splash + Login load with stub `DataContext`; F5 in &lt; 30 s; no DB.

### 2.2 Temporary public APIs — official or replace?

Reviewed against **`v2.18.0-wpf-alignment-p7c-layout-shim`**:

| API | Current status | VCF position |
|-----|----------------|--------------|
| **`Window.RelayoutChildren`** | **Public** on `Window` | Still supported for first show / resize. **Phase 2a.1:** direct-child **Visibility** swaps trigger `OnChildVisibilityChanged` (Relayout + named rebuild) automatically. |
| **`Window.RebuildNamedItemsList`** | **Public** | Still supported; also invoked from **2a.1** on Visibility navigation. |
| **`Window.ApplyDeferredChildLayout`** | **Public** | **Supported** companion to deferred layout during `LockRefresh`. Prefer **`RelayoutChildren`** from app code; treat this as “flush pending layout” after a batched load. |
| **`UserControl.ApplyDeferredHostLayout`** | **`Friend`** | **Not part of the public contract.** Use **`Window.RelayoutChildren`**, which walks children and applies host layout. Do not call from DeNovo. |
| **`MarginFromDesignWhenUnset`** (`modLayoutEngine`) | **Public** module function | **Internal arrange-time helper**, not a consumer API. Do not call from DeNovo. |
| **`ApplyLegacyLayoutProperty`** (XAML load shim) | Invoked from **`XAMLReader.SetProperty`** | **Not a consumer API.** Part of **2.18.0** layout shim (`Margin` on `TextBlock` → `Design*`). |
| **`ObservableCollection.BeginUpdate` / `EndUpdate`** | **Public** (Phase 4b) | **Official** for **batch collection mutations**, not layout. Does **not** replace `RelayoutChildren`. |

**Agreed:** no further **new** public COM members from DeNovo. Open a **VCF issue** with harness repro + proposed signature; we ship in a tagged release with `BREAKING_CHANGES.md` / `MIGRATION.md` entry.

**Follow-up:** document supported host-app methods in `VCF_CLASS_REFERENCE.md` / a dedicated public-surface section (Window layout trio above).

### 2.3 Layout shim pin

**Confirmed pin for all harness work:** **`v2.18.0-wpf-alignment-p7c-layout-shim`**

- Phase0 **31/31** validated (includes **P7c-LAY**).
- No further shim changes planned **before harness milestone 1** unless harness finds a concrete gap.
- Re-run **`Invoke-VcfXamlMigration.ps1`** from this tag if you want legacy types to **keep** `Design*` (script skips layout transform on `TextBlock`, `Image`, `Scene`, `UniformGrid`, `TextBox`, `WindowsFormsHost`). Already-migrated XAML with `Margin` on those types is fine with **2.18.0** shim DLL.

See [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md).

### 2.4 Phase 1 exit criteria

**Agreed:**

| Phase | Acceptance |
|-------|------------|
| **Phase 1 (UI/XAML)** | **Phase0 31/31** + **DeNovoSmoke harness** green for agreed screens (Splash → Login → MainMenu → Sales **layout-only**) with stub VMs |
| **Phase 2a (VCF parallel)** | Continue framework + harness work; every tag: Phase0 (**35/35** incl. B-RESZ/B-NAV/B-BIND-DENSE/P2a-PAD) + DeNovoSmoke (stubs only — **no** KernelLib / Data / DB) |
| **Phase 2b (POS integration)** | When Data + Kernel + UI are ready together: re-attach KernelLib / Data / DB; pin latest VCF tag in `DeNovo.vbp`; full `DeNovo.exe` smoke per [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) §3 |

**Status (2026-07-21):** **Phase 1 complete** — tag **`v2.23.0-wpf-alignment-phase1-complete`**. **Phase 2a in progress** (VCF-local). **Phase 2b deferred** until Data/Kernel (and UI) finish — do **not** rebuild full `DeNovo.exe` for VCF validation.

Phase0 alone is necessary but not sufficient — we need your real XAML + binding contracts in the harness. Full exe startup remains **out of scope** until Phase 2b.

### 2.5 Process going forward

**Agreed:**

1. DeNovo reports issues as **harness project + screen name + XAML fragment + tag pinned**.
2. No consumer-side changes to VCF **public** typelib without VCF review and release note.
3. VCF tags Phase 7 UI work only after **Phase0 + DeNovoSmoke** (once scaffold exists).

---

## 3. Reconciliation of DeNovo disclosure (§2)

### 3.1 Consumer-only changes (§2.1)

| DeNovo change | VCF view |
|---------------|----------|
| Visibility-based view swap (no `RemoveAt`) | **Reasonable.** Document in harness as supported navigation until we ship a dedicated content-host API. |
| Manual `RelayoutChildren` / `RebuildNamedItemsList` after navigation | **Superseded for Visibility swaps (2a.1).** Still valid for first show / resize. |
| Login / Splash / Main menu VM + XAML tweaks | **Expected** during migration. Capture **minimum stub VM shape** in harness docs. |
| MessageBox deferred close / sizing | **Consumer scope** for Phase 1; `@` → `DataTemplate` remains Phase 7c-dialog backlog. |

We do **not** ask you to revert these for Phase 1 unless harness proves a simpler path without manual relayout.

### 3.2 APIs DeNovo relied on (§2.2)

Already in **`v2.18.0`** tree — not DeNovo-only forks. They were **underdocumented**; VCF takes ownership of documenting them (§2.2 table). **`ApplyDeferredHostLayout`** remains Friend.

### 3.3 Test hooks on COM (§2.3)

**Fully aligned.** Test entry points belong in **`.Tests/Phase0`** and **`.Tests/DeNovoSmoke`**, not on production interfaces.

### 3.4 Non-UI blockers (§2.4)

Acknowledged as **orthogonal**. Harness scope excludes KernelLib / `Data.Dataset` / INI.

---

## 4. VCF actions

| # | Action | Owner |
|---|--------|--------|
| 1 | Scaffold **`.Tests/DeNovoSmoke/`** (StdExe, `AppHost`, `ShellWindow`, stub navigation) | VCF |
| 2 | Document **Window layout APIs** + “do not call Friend methods” | VCF |
| 3 | Extend [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) — harness milestone 1 | VCF |
| 4 | Contribute **Splash + Login XAML** + **stub VMs** into harness `Resources/` | DeNovo |
| 5 | Run harness on **`v2.18.0`**, file issues with fragments | DeNovo |

**Target tag after harness milestone 1:** `v2.19.0-wpf-alignment-p7d-denovo-smoke` (docs + harness; DLL changes only if bugs found).

---

## 5. DeNovo milestone 1 (confirmed 2026-06-22)

DeNovo confirmed split ownership, full-exe pause, and no new VCF typelib surface. Fixture paths and stub VM contracts:

**denovo monorepo → `docs/migration/DENOVO_HARNESS_MILESTONE1.md`**

| Deliverable | Status |
|-------------|--------|
| Confirmation + stub contracts + XAML/`res:` paths | **Done** (DeNovo doc) |
| VCF `.Tests/DeNovoSmoke/` VB6 runner | **Done** (open `DeNovoSmoke.vbp`) |
| Sync fixtures into harness `Resources/XAML/` | Pending (when runner path published) |
| Milestone 1 green on `v2.18.0` | Pending (joint) |

**Reconciliation:** KernelLib `modDebug` `Test*` functions are DeNovo-owned, not VCF — no action on our typelib.

---

## 6. References (VCF repo)

| Doc | Purpose |
|-----|---------|
| [MIGRATION.md](./MIGRATION.md) § Phase 7 | Pin matrix including harness slice |
| [POS_RUNTIME_FEEDBACK.md](./POS_RUNTIME_FEEDBACK.md) | Layout shim write-up |
| [POS_INTEGRATION_SMOKE.md](./POS_INTEGRATION_SMOKE.md) | Phase 1 / 2a harness gate + Phase 2b full POS smoke |
| [tools/xaml-migrate/README.md](../tools/xaml-migrate/README.md) | Mechanical migration + legacy-type skip |

---

*Maintained by VCF team. DeNovo proposal lives in the denovo monorepo.*
