# AI session history — distilled project context

**Source:** Cursor agent transcript export under [`../_raw-ai-export/`](../_raw-ai-export/)  
**Primary thread:** `agent-transcripts/88d269db-f15c-4161-b9bd-31c838859802` (Jun–Jul 2026)  
**Distilled:** 2026-08-04 (post system restore)  
**Scope:** Durable decisions and status for Demac.VCF — not chat chatter.

## Workspace map (locked Jul 2026)

| Tree | Role |
|------|------|
| **`v1/`** | Stable / binary-compatible line (June 20 2026 + MeasureRow). **Registered** production DLL for DeNovo. |
| **`v2/`** | WPF-alignment rewrite (Phases 0–6/8 + 2a backlog through ~3.25.x). **Paused** — unfinished teardown vs compiled DLL. Intended COM/project name **`VCF2` / `Demac.VCF2`** (or similar) so it never fights registered `Demac.VCF` — rename may still be pending in the `.vbp`. |

**Git (single repo):** `https://github.com/tdimitriou/VCF.git` — branches `v1` (stable), `master` / WPF tip, `wip/window-teardown`. On this machine after restore, `v1/` is a **broken worktree link** (still points at the old PC path under `…/Demac.VCF/v2/.git/worktrees/v1`); `v2/` holds the real `.git` and remotes. Repair by re-pointing / re-adding the worktree — remote history was not lost.

**Related:** DeNovo Cursor transcripts were lost the same way and recovered separately (not in this export).

## Read order

1. [decisions.md](./decisions.md) — major product/architecture choices and why  
2. [completed-work.md](./completed-work.md) — what shipped on the WPF line  
3. [in-progress.md](./in-progress.md) — open / deferred / POS follow-ons  
4. [conventions.md](./conventions.md) — how we work (compat, gates, docs)  
5. [pitfalls.md](./pitfalls.md) — do-not-regress traps from real debugging  

Canonical framework docs remain under `v1/docs/` and `v2/docs/` (handoff guide, BREAKING, MIGRATION, DP roadmap, etc.). This folder is session memory only.
