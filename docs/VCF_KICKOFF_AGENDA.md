# Demac.VCF — WPF alignment kickoff agenda

**Purpose:** Confirm or defer open items before Phase 1; align VCF + DeNovo on release contract.  
**Duration:** 60–90 minutes  
**Status:** Scheduled — Phase 0 complete (`v2.0.0-wpf-alignment-p0`)  
**Last updated:** 2026-06-20  

---

## Attendees

| Role | Team | Responsibility |
|------|------|----------------|
| VCF lead / architect | Demac VCF | Scope, phasing, API breaks |
| Core platform dev | Demac VCF | DP, binding, XAML |
| POS / DeNovo lead | DeNovo | Migration timing, smoke tests |
| QA / release | Either | `.Tests/Phase0`, POS smoke checklist |

---

## Pre-read (15 min before call)

1. [VCF_TEAM_HANDOFF_GUIDE.md](./VCF_TEAM_HANDOFF_GUIDE.md) — §1–§4, §9  
2. [BREAKING_CHANGES.md](./BREAKING_CHANGES.md) — v2.0.0 Phase 0  
3. [MIGRATION.md](./MIGRATION.md) — DeNovo pin + strict mode  
4. [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md) — if ListView dev joins  

**Phase 0 evidence:** `.Tests/Phase0` — 5/5 pass; baselines in [VCF_PERFORMANCE_BENCHMARKS.md](./VCF_PERFORMANCE_BENCHMARKS.md).

---

## Agenda

| Time | Topic | Goal | Recommendation |
|------|-------|------|----------------|
| 0:00 | **Welcome & Phase 0 recap** | Tag `v2.0.0-wpf-alignment-p0`; DeNovo pins DLL | DeNovo smoke: login → sales → grid → dialog |
| 0:10 | **Release contract** | Semver, `BREAKING_CHANGES.md`, `MIGRATION.md` each breaking drop | Confirm AI-assisted POS migration after Phase 3+ |
| 0:20 | **Grid vs UniformGrid** | Default layout for new screens | **Grid primary** after Phase 2; UniformGrid compat one release |
| 0:30 | **Strict XAML timeline** | When `VCF.StrictXamlLoad` defaults True | Opt-in until Phase 3; POS enables in CI first |
| 0:35 | **Strict `x:Class` failure** | Remove root-tag fallback | **Fail loud** — align with strict reader |
| 0:40 | **Listener-pull vs BindingExpression** | Legacy POS objects without INPC | **Keep opt-in** listener-pull for legacy |
| 0:45 | **Utilities DLL split** | `Mail`, `INIParser`, AsyncKit → VCF.Core | **Defer to Phase 6** (Phase 0 build stayed monolith) |
| 0:50 | **ViewBase / codegen** | Reduce `x:Class` boilerplate | **Optional ViewBase** Phase 1; codegen P2 |
| 0:55 | **`Design*` migration** | One-release XAML alias vs hard break | Decide alias for bulk POS XAML pass |
| 1:00 | **ListView / InvoiceGrid** | Phase 5 scope, Codejock interim | Walk [VCF_LISTVIEW_ARCHITECTURE.md](./VCF_LISTVIEW_ARCHITECTURE.md) §InvoiceGrid |
| 1:10 | **`res:` → ResourceDictionary** | Phase 3 priority; ObjectConstructor shrink | Confirm built-in resolver owner = VCF |
| 1:15 | **Binding P0 (B5, B6)** | DataContext rebind; fail-loud XAML expansion | Phase 4 for B6; strict reader expands Phase 3 |
| 1:20 | **Deferred items** | Z report / Codejock island, ANSI/Greek, StackPanel interim | Ticket + owner; no block on Phase 1 |
| 1:25 | **Phase 1 kickoff** | `DependencyPropertyRegistry`, `FrameworkElement`, Visibility | Assign owners; target next tag |
| 1:30 | **Actions & next meeting** | Update §8/§9 in alignment notes | |

---

## Decisions log (fill in during meeting)

| # | Question | Decision | Owner | Ticket |
|---|----------|----------|-------|--------|
| 1 | Grid vs UniformGrid default | | | |
| 2 | Strict XAML default date | | | |
| 3 | `x:Class` fallback removal | | | |
| 4 | Listener-pull retention | | | |
| 5 | VCF.Core split timing | | | |
| 6 | ViewBase in Phase 1? | | | |
| 7 | `Design*` XAML alias one release? | | | |
| 8 | DeNovo smoke owner + date | | | |

---

## DeNovo actions after tag

1. Pin `Demac.VCF.dll` **v2.0.0-wpf-alignment-p0** in `DeNovo.vbp`.  
2. Run integration smoke (login → sales → grid → dialog).  
3. Report regressions in VCF repo or alignment notes §11 discussion log.  
4. Do **not** enable `VCF.StrictXamlLoad = True` in production until Phase 3 migration doc is ready.

---

## Phase 1 preview (post-kickoff)

- `DependencyPropertyRegistry` — shared metadata per type  
- `FrameworkElement` — composed layout/visibility/DP store  
- Visibility DP (Visible / Hidden / Collapsed)  
- Measure/Arrange pipeline  
- Remove public `Design*` (breaking — update `BREAKING_CHANGES.md`, `MIGRATION.md`)  

**Target tag:** `v2.1.0-wpf-alignment-p1` (or semver per team convention).

---

## Suggested meeting invite (copy/paste)

**Subject:** Demac.VCF WPF alignment — kickoff (Phase 0 done, Phase 1 planning)  

**Body:**  
Phase 0 is shipped and tagged (`v2.0.0-wpf-alignment-p0`). Phase0 tests: 5/5 pass.  

Please read `docs/VCF_TEAM_HANDOFF_GUIDE.md` and `docs/VCF_KICKOFF_AGENDA.md` before the call.  

We will confirm open items (Grid, strict XAML, utilities split, ListView/InvoiceGrid, `res:` migration) and assign Phase 1 work.

---

*After the meeting: move resolved items from [VCF_WPF_ALIGNMENT_NOTES.md §8](./VCF_WPF_ALIGNMENT_NOTES.md) to §9; sync denovo copies if used.*
