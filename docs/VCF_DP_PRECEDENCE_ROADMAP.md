# VCF — DP value precedence roadmap (WPF-aligned)

**Status:** **Deferred** (decision 2026-07-23) — do **not** start until queued after current overlaps / Phase 2a leftovers.  
**Companion:** [VCF_WPF_ALIGNMENT_NOTES.md](./VCF_WPF_ALIGNMENT_NOTES.md) §2.5 · [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md) §5 · [BREAKING_CHANGES.md](./BREAKING_CHANGES.md) **3.2.0**  
**North star:** WPF-experienced developers get familiar effective-value behavior; not a claim of full WPF feature parity on day one.

---

## 1. Current contract (shipping — 3.2.0)

```text
Local (SetValue)  >  Current (SetCurrentValue)  >  Inherit  >  Metadata DefaultValue
```

| Layer | Writers today |
|-------|----------------|
| **Local** | XAML, CLR `Let`, Binding, TemplateBinding |
| **Current** | Style setters, **PropertyTriggers**, template content-align, many `Class_Initialize` seeds — **last write wins** in one slot |
| **Inherit** | Lazy parent walk when both slots unset |
| **Metadata default** | `DependencyPropertyMetadata.DefaultValue` (pilot: `Window.BorderStyle`, `ContentControl.Content`, `TextBox` strings) |

**Trigger lifecycle (interim, aligned with 3.2.0):** condition changes (e.g. Button `IsMouseOver`) call **`ApplyStyle`**, which re-applies style setters then **active** triggers only. That is how **P6b-TRIG** “restores” — not a separate trigger layer. Do **not** add snapshot/fake sub-layers without amending this contract.

**API already in:** public `ClearValue` / `ReadLocalValue`; binding Detach clears local; `P4-PREC` gate.

---

## 2. Decision — full WPF precedence later

| Decision | Detail |
|----------|--------|
| **Do it** | Yes — full(er) WPF-aligned effective-value stack is an accepted **roadmap epic** |
| **Not now** | Keep **3.2.0** two-slot model until this epic is explicitly queued |
| **Why defer** | Big-bang rewrite is not viable; hover/style already work via re-`ApplyStyle`; higher ROI overlaps first (Visibility, RegisterAttached, Min/Max, …) |
| **When to reopen** | Product need (multi-trigger conflicts, template trigger layers) **or** scheduled capacity after fundamentals overlaps |

**Rejected for interim:** thin “snapshot on trigger activate / restore on deactivate” as if it completed #4 — that **simulates** a layer and is **not** fully aligned with the locked two-slot contract.

---

## 3. Target (WPF-ish) layers

High → low (VCF target; drop animation if product never needs it):

1. Coercion (optional late)  
2. Animation / Hold (optional / last)  
3. **Local** value  
4. TemplatedParent / template setters (as needed for lookless)  
5. **Style / template triggers** (activate + **deactivate**)  
6. **Style / template setters**  
7. Theme / default style (if distinct from metadata)  
8. **Inheritance**  
9. **Metadata default**

Binding remains **local** unless a later design explicitly introduces expression-over-local semantics.

---

## 4. Phased implementation (when queued)

| Phase | Scope | Notes |
|-------|--------|------|
| **P0 Spec** | Exact layer list; writer matrix; what VCF will **not** ship (e.g. animation?) | Amend [BREAKING_CHANGES](./BREAKING_CHANGES.md); Phase0 plan |
| **P1 Storage** | Effective-value entries or explicit slots; rewrite `GetValue` / `ClearValue` / `ReadLocalValue` | Highest risk |
| **P2 Retarget writers** | Style / trigger / template align / init → correct layers; finish metadata-default migration | Large audit |
| **P3 Trigger lifecycle** | Activate/deactivate without relying on full style clobber for correctness | Replaces interim ApplyStyle-only story where needed |
| **P4 Binding policy** | Confirm binding-as-local; detach / ClearValue edges | Already partly done |
| **P5 TemplatedParent / template setters** | If lookless requires distinct layer | Ties to ControlTemplate |
| **P6 Coerce (+ animation stub)** | Metadata callbacks; animation only if required | Last |
| **P7 Gates + POS** | Phase0 growth; DeNovoSmoke / layout screens | Every tagged minor |

**Ship as several tagged minors**, not one mega-break. Each step: Phase0 + `BREAKING_CHANGES` + caller rebuild under Project Compatibility.

---

## 5. Effort estimate (order of magnitude)

| Scope | Focused eng days | Calendar (solo + gates) |
|-------|------------------|-------------------------|
| WPF-ish **core** (through style/trigger/template + inherit/default; no animation) | **~25–40** | **~2–3 months** |
| **+ coerce + animation hook** | **~35–55** | **~3–4 months** |

Confidence **±40%**. Big-bang single release: **not recommended**.

---

## 6. Near-term (until epic starts)

- Keep **3.2.0** contract and **P4-PREC**.  
- **Trigger conditions (B3 interim — done 3.2.1–3.2.2):**  
  - DP changes → `NotifyConditionPropertyChanged` → `ReapplyStyleValues` when Style watches that property  
  - Non-DP (e.g. Button `IsMouseOver`) → same Notify API  
  - Full `ApplyStyle` only when Style identity / template must change  
- Incremental: migrate more init `SetCurrentValue` defaults → **metadata DefaultValue**.  
- Fundamentals overlaps (SelectionChanged naming, etc.) are **done or deferred separately** — **not** the full WPF layer stack.
- **Visibility/Hidden** — done **3.3.0**.
- **RegisterAttached (Grid)** — done **3.4.0**.
- **Min/Max layout** — done **3.5.0**.
- **Attached storage → DP bag** — ~~deferred~~ → **3.16.0** (dict shim retained); see handoff / [VCF_PROPERTY_REGISTRY.md](./VCF_PROPERTY_REGISTRY.md) §1.3.

---

## 7. Checklist when starting the epic

- [ ] Confirm layer list + non-goals with VCF owner  
- [ ] Baseline Phase0 **61+** green on current tag  
- [ ] First PR = storage + GetValue only (no writer retarget) **or** writer matrix doc + empty layers — pick one spike  
- [ ] Tag each layer slice (`v3.x.0-wpf-alignment-dp-…`)  
- [ ] Update this file status → **In progress** / **Done**
