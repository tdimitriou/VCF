# In-progress / unfinished items

## Critical open (blocked v2 production use)

| Item | Notes |
|------|-------|
| **Compiled ActiveX teardown hang** | Phase0 / Phase0Clean green in-process; vs **compiled `Demac.VCF.dll`**, unload of **ItemsControl + UniformGrid/Button** graphs freezes IDE/EXE (often after MsgBox / on End). Evidence in `%TEMP%\VCF_Unload.log` (`DrainHold`, `DisarmForRelease`). |
| **Safe `Window.Unload` Release model** | Hold/park/defer experiments partially proved diagnosis; durable fix not landed on a “clean” tip. Last known clean WPF tip cited as **`ad3d0d1`**; WIP on teardown branch / unclean tip. |
| **Mass Tier-1 error-handler experiment** | Applied then largely reverted — do not re-apply broadly; cherry-pick only proven unload fixes. |

## Deferred epics (accepted roadmap)

| Item | Status |
|------|--------|
| **Full DP value precedence** | Deferred; keep 3.2.0 two-slot until queued (`VCF_DP_PRECEDENCE_ROADMAP.md`) |
| **Attached dict shim removal** | Bag primary; shim still present |
| **Drop legacy `Visible` bool** | Compat facade kept |
| **LCV Sort / Filter / Group** | Until a consumer needs it |
| **ItemContainerStyle unify** (IC ↔ ListView template engines) | Deferred |
| **ContentTemplateSelector** | Deferred (ContentTemplate covers current need) |
| **HierarchicalDataTemplate / DataTemplateSelector** | Not built |
| **Storyboard / animation layers** | Out of scope unless product needs |
| **WPF xmlns / clr-namespace** | Keep `VCF.TypeName` CreateObject path |
| **ViewBase / InitializeComponent codegen** | Optional boilerplate reduction |
| **Binary Compatibility freeze** | Finish-line policy, not now |
| **WrapPanel / Viewbox** | Deferred extras |
| **Option C** | Cancelled (listed only for history) |

## “Soon” when v2 resumes (triage Jul 24)

Phased backlog agreed before pause:

1. ~~Relayout APIs internalize~~ — started; escape hatches still public  
2. Binding TODOs cleanup — largely comment/docs  
3. **ComboBox / TabControl / ScrollViewer**  
4. **ContentTemplateSelector** (swapped ahead of Wrap/Viewbox)  
5. WrapPanel / Viewbox  
6. Cleanup + small additions  
7. DP precedence / attached shim / Visible bool (later phases)  
8. Discuss later: xmlns, storyboard, ViewBase  

User paused further control work: POS does not currently require new controls; framework considered “stable enough” for feature pause pending teardown fix.

## POS / DeNovo (consumer — not VCF core)

| Step | State |
|------|-------|
| XAML migration + VCF wiring docs/script | Prepared (~3.26 handoff); production still on **v1** pin |
| Re-wire POS to WPF VCF | Not started against v2 |
| InvoiceGrid → ListView | Blocker for pure `IWindow` root (ActiveX host constraint) |
| RootPanel → top-level IWindow | After InvoiceGrid |
| MainForm → standard classes | Later |

## v1 near-term

- **Bug (DeNovo-reported):** variable row height — **internal height / scrollbar synchronization** out of sync. Fix here when you return; do not treat MeasureRow as fully battle-tested for InvoiceGrid until this is closed.  
- Small non-breaking improvements OK (BackgroundWorker-class lineage, MeasureRow API already on `origin/v1`)  
- Avoid importing v2 layout/DP/binding engine into v1 without a deliberate binary-compat plan  
- Optional: finish **VCF2** project/DLL rename on v2 so registration never clashes with production `Demac.VCF` 

## Resume checklist (when returning to v2)

1. Confirm tip: clean `ad3d0d1`-class tip vs WIP teardown branch  
2. Reproduce Phase0Clean vs **compiled DLL** (both IDE / DLL+IDE / both compiled, second F5)  
3. Fix ItemsControl graph Release without permanent COM leak  
4. Re-run full Phase0 90/90 + DeNovoSmoke  
5. Only then resume Phase 3 controls or DP precedence epic
