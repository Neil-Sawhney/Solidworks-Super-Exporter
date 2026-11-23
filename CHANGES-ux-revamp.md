# SWExporter - UX revamp (ux-revamp branch)

## New mode set

| id | Part | Assembly | Drawing | Description (shown under the mode box + tooltip) |
|---|---|---|---|---|
| 0 | Each body | Each body | | Part: one file per ticked body of the active configuration; only visible bodies listed; Surface bodies adds surfaces. Assembly: one file per ticked visible body; hidden, suppressed, envelope components and hidden bodies skipped. |
| 2 | Each body x configuration | | | One file per visible body, in every ticked configuration. |
| 1 | Whole part (per configuration) | Whole assembly (per configuration) | | One file of the whole model per ticked configuration. |
| 5 | | Each part | | One file per visible part (each part + configuration once); bodies hidden in the assembly left out. |
| 8 | Display states | Display states | | Every visible body of each ticked display state, one file each; one row per state; Configurations = active only or all (Config filter); repeats skipped. |
| 3 | | | Each sheet | One file per ticked sheet; PDF one file or one per sheet. |
| 4 | Batch | Batch | Batch | Every ticked file; parts split per "Parts in batch". |

Box order: Part = Each body, Each body x configuration, Whole part, Display states, Batch.
Assembly = Whole assembly, Each part, Each body, Display states, Batch. Default (first row) per
document type is the same as before. The ? help MsgBox is about 780 characters.

## Old -> new mapping

| Old mode (id) | New |
|---|---|
| Bodies (each visible body) (0) | Each body (0), same code |
| Visible bodies (solid, one file each), part (7) | Each body (0); it already ran the Bodies code path |
| Visible bodies (each visible body), assembly (7) | Each body (0) in an assembly, same code (`ExportAssemblyBodies`, was `ExportVisibleBodies`) |
| Components x bodies (6) | Each body (0) in an assembly |
| Components (each visible part) (5) | Each part (5) |
| Configurations (1) | Whole part / Whole assembly (1) |
| Configurations x bodies (2) | Each body x configuration (2) |
| Drawing sheets (3) | Each sheet (3) |
| Batch (4) | Batch (4); "Parts in batch" still maps to 1 / 0 / 2 |

Ids were not renumbered. Nothing is saved between runs, so the only compatibility point is code:
`NormalizeMode()` (called at the start of `RunExport`) maps the retired ids 6 and 7 to 0.
`MODE_COMPBODIES` / `MODE_VISBODIES` are gone; their numbers live on only as private
`MODE_OLD_*` constants inside `NormalizeMode`.

## Behavior decisions

- Part Each body lists visible bodies only (`GetBodies2(..., True)`), as both old modes did. Hidden
  bodies are usually tool/reference bodies, so leaving them out is the safer choice. Surfaces only
  when "Surface bodies" is ticked.
- Assembly Each body = the old Visible bodies behavior, unchanged (root-component walk, strict filter).
  Difference for former "Components x bodies" users: rows are now per body, not per part, and there
  is no "all bodies" fallback when the assembly's body names don't match the part.
- Config filter is now also used by Display states ("each configuration"). Its box is enabled only in
  Batch and Display states.

## Display states mode: design choices

- **Picking states and configurations:** the item list has two rows per display state:
  `State - active config (<name>)` (ticked by default) and `State - each configuration`.
  States that exist only in other configurations get just the "each configuration" row.
  A plain multi-select ListBox needs no double-click or column tricks, so it is robust at runtime.
  If both rows of a state are ticked, "each configuration" wins and the active config isn't visited twice.
- **Configurations:** "each configuration" = every configuration matching Config filter that has a
  state with that name. This works for linked display states (names differ per configuration) and for
  shared ones (same names everywhere). Missing pairs are skipped and logged.
- **Preview:** the label under the list shows "N of M ticked" and, in this mode, "K state/config
  run(s), J skipped". The exact file list comes from Dry run.
- **Run:** visits are grouped by configuration, active one first. For each visit: `ShowConfiguration2`,
  `Configuration.ApplyDisplayState`, `GraphicsRedraw2`, then verify that `GetDisplayStates()(0)`
  (the active state) is the requested one. If not, log FAILED and skip the visit. Bodies come from
  `IBody2` visibility after the state is applied.
- **Duplicates:** key = configuration + body name (assembly: part path + referenced config + body).
  A repeat is logged "SKIPPED: duplicate - already exported from display state X" and counted as skipped.
- **Export path:** parts reuse `ExportBodies` (body selection = the not-yet-exported names). Assemblies
  reuse `VisibleBodyMap` + `ExportAsmPart` with a strict body filter. The engine itself is unchanged.
  eDrawings / 3D PDF (WHOLE formats) are written once per visit and show that display state.
- **Naming:** new token `{ds}` (alias `{displaystate}`). Automatic name `{file}_{ds}_{config}_{body}`.
  `_{config}` is dropped when the state is exported in one configuration only, or the model has one
  configuration. In assemblies `{config}` stays the part's configuration.
- **Restore:** in the Cleanup path (normal end, error and Cancel), each configuration's own active state
  is put back when the run leaves it. At the end, the original configuration, then its original display
  state, then the selection are restored.
- **Assemblies:** supported, using the same plan. The visible-bodies walk is evaluated under each assembly
  display state. Bodies hidden inside the part file's own display state are still left out, as in Each body.
- **Batch:** not supported (it needs per-file display-state picks).
- **API assumption:** `Configuration.GetDisplayStates` returns the active state first. Display state
  calls are early-bound on `SldWorks.Configuration` (`ApplyDisplayState`, `GetDisplayStates`).

## Form layout

All controls are still created in code. Section headers are bold labels with a 1 pt line (no Frames).
Left column: **1 What to export** (mode + ?, description, item list, All/None, count) and
**2 Scope / filters** (Filter, Tick/Untick matches, Match case, Surface bodies, Config filter, and,
in Batch only, batch folder, Refresh, Add files, Include subfolders, Parts in batch).
Right column: **3 Formats**, **4 Output & naming**, **5 Options**. Status bottom-left, Export / Close
bottom-right. Creation order = tab order. The form sizes itself from InsideWidth/InsideHeight, so
nothing is clipped. PDF one-file is enabled only for Each sheet / Batch, and the STL options are
disabled for Each sheet.

## Code cleanup

- Module: MODES banner with commented ids, a display-states helper block, the `ExportComponents`
  COMPBODIES branch removed, and `ExportVisibleBodies` renamed to `ExportAssemblyBodies`.
- Form: grouped declarations, small layout helpers (`Lbl/Btn/Chk/Txt/Combo/CheckList/SectionHeader`),
  one Build* function per section, per-mode `ModeLabel`/`ModeHelp`/`UpdateModeControls`, one Fill*
  per list type, and `btnExport_Click` split into `CollectItems/CollectFormats/CollectOptions`.
- Kept: STEP AP214-only default and log off (form, INI, `LoadDefaultFormats`), drawing defaults,
  `main` as the last and only parameterless public Sub, `LoadFormats(Optional ...)` hidden from the macro list.

## Install

1. Tools > Macro > Edit `SWExporter.swp`.
2. Module: select all code and replace it with `src/SWExporterMain.bas`.
3. UserForm `frmSWExporter`: open its code window, delete everything, paste `src/frmSWExporter.txt`
   (the designer form stays empty).
4. Copy `SWExporter.ini` next to the `.swp` (unchanged; no new keys). Save and run.


## Revision 2 (Nov 23, after Neil's first run)

- **Display states list: one row per state.** The "active config" / "each configuration" double rows are gone.
  A new **Configurations** drop-down in section 2 (enabled in Display states only) picks
  "Active configuration only" (default; lists the active configuration's states, all ticked) or
  "All configurations (Config filter)" (lists the union of state names in the matching configurations;
  the active configuration's states start ticked; each ticked state runs in every matching configuration that has it).
  Config filter is enabled in Display states only with "All configurations" (and in Batch, as before).
  Changing the scope, Config filter or Match case rebuilds the list and keeps ticks by state name.
  `DisplayStatePlan(m, names, allCfgs, cfgFilter, caseSens, missing)`; `DS_ACTIVE` / `DS_EACH` removed; new `gDsAllCfgs`.
- **Multi-keyword filters.** `MatchesFilter` (Filter box, Config filter in Batch and Display states):
  `,` or `;` = OR terms, spaces inside a term = AND words, substring match, Match case respected, blank = everything.
  Tooltips and ? help explain it.
- **Display state apply fix.** Neil's run logged "could not apply display state" for every state. The old code
  skipped ApplyDisplayState when GetDisplayStates()(0) already equalled the name, and failed whenever the
  call returned False or index 0 didn't change. Now the configuration is activated first, ApplyDisplayState is always
  called on a freshly fetched active configuration (`GetActiveConfiguration`), its True is trusted, a False is retried
  once after `ForceRebuild3 False`, and only then is the index-0 check used as a fallback ("already active").
  A remaining failure logs the return value, any COM error and which state GetDisplayStates lists first.
  Restoring a state at the end always calls ApplyDisplayState too.
