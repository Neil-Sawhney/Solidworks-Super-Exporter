# SWExporter UX revamp - design

## 1. Mode set (ids kept numeric; nothing is persisted between runs)

| id | constant          | Part label                         | Assembly label                        | Drawing      |
|----|-------------------|------------------------------------|---------------------------------------|--------------|
| 0  | MODE_BODIES       | Each body                          | Each body                             |              |
| 2  | MODE_CFGBODIES    | Each body x configuration          | -                                     |              |
| 1  | MODE_CONFIGS      | Whole part (per configuration)     | Whole assembly (per configuration)    |              |
| 5  | MODE_COMPONENTS   | -                                  | Each part                             |              |
| 8  | MODE_DISPSTATES   | Display states                     | Display states                        |              |
| 3  | MODE_SHEETS       |                                    |                                       | Each sheet   |
| 4  | MODE_BATCH        | Batch (folder or files)            | same                                  | same         |

Order in the box: Part = Each body, Each body x configuration, Whole part, Display states, Batch.
Assembly = Whole assembly, Each part, Each body, Display states, Batch (defaults unchanged:
first row = what was first before).

Retired ids: 6 (Components x bodies) and 7 (Visible bodies). `NormalizeMode()` maps both to
MODE_BODIES, so a caller passing an old id still works.

- Part "Each body" = old Bodies = old Visible bodies (they ran the same code). The list shows
  the visible bodies, all ticked. Hidden bodies are NOT listed (safer: hidden bodies are
  usually tool/reference bodies). Surface bodies only when "Surface bodies" is ticked.
- Assembly "Each body" = old Visible bodies, unchanged: walk from the root component, skip
  hidden/suppressed/envelope components (and children) and hidden bodies, one file per body.
  Old "Components x bodies" folds into it (it was the same output, listed per part).
- Batch "Parts in batch" keeps MODE_CONFIGS / MODE_BODIES / MODE_CFGBODIES.

## 2. One-line description (label under the mode box + tooltip)

- Each body (part): One file per ticked body of the active configuration. Only visible bodies are listed; tick Surface bodies to add surfaces.
- Each body x configuration: One file per visible body in every ticked configuration.
- Whole part: One file of the whole part per ticked configuration.
- Whole assembly: One file of the whole assembly per ticked configuration.
- Each part: One file per visible part (each part + configuration once). Bodies hidden in the assembly are left out.
- Each body (assembly): One file per ticked visible body. Hidden, suppressed and envelope components and hidden bodies are skipped.
- Display states: Every visible body of each ticked display state as its own file. One row per state; Configurations (section 2) = active only or all (Config filter). Repeats are skipped.
- Each sheet: One file per ticked sheet (PDF: one file or one per sheet).
- Batch: Every file in the list; parts split per "Parts in batch".

## 3. Form layout (all built in code; section headers = bold label + 1pt line)

```
+-- 1 What to export ---------------+-- 3 Formats -----------------------+
| Mode [combo..............] [?]    | [format list               ]       |
| description (2-3 lines)           | [ ] Drawing PDF: one file          |
| Items label                       | Units [mm] Quality [Fine]          |
| [item list                ]       | [x] Binary STL                     |
| [All] [None]  3 of 8 ticked       +-- 4 Output & naming ---------------+
+-- 2 Scope / filters --------------+ Output [Next to file    ]          |
| Filter [....] [ ] Match case      | Folder [..............] [...]      |
| [Tick matches] [Untick matches]   | Name   [..............]            |
| [ ] Surface bodies                | token hint                         |
| Config filter [.........]         | If exists [Auto-number]            |
| Batch folder [.......] [...]  (*) +-- 5 Options -----------------------+
| [ ]Subfolders [Refresh][Add...](*)| [ ] Dry run  [ ] Write log CSV     |
| Parts in batch [........]     (*) |                                    |
+-----------------------------------+------------------------------------+
| status ...........................              [Export] [Close]      |
```
(*) visible only in Batch. Config filter enabled in Batch and Display states.
Surface bodies enabled for body modes, Display states and Batch. PDF one-file enabled for
Each sheet / Batch; STL units/quality/binary disabled for Each sheet.
Creation order = tab order (left column top-down, then right column, then buttons).
Frames not used: plain labels are less error-prone (no frame.Controls, no offsets).

## 4. Display states mode

Item list (one or two rows per display state name, union over all configurations):
```
Display state rows:
  [x] Bodies-all      -  active config (Default)
  [ ] Bodies-all      -  each configuration
  [x] Variant-body    -  active config (Default)
  [ ] Variant-body    -  each configuration
  [ ] Other-state     -  each configuration (not in active config)
```
Key = "A" or "E" & Tab & state name. All "active config" rows start ticked.
Each configuration means every configuration matching Config filter (blank = all).
Neil's example: tick "Bodies-all - active config" and "Variant-body - each configuration".

Count label: "4 of 10 ticked / 5 runs, 1 skipped" (runs = configuration x state visits).
Dry run gives the exact file list.

Run (part):
1. Plan = list of (config, state, oneCfg). Active-config visits first, then the other configs
   in model order. If both rows of a state are ticked, "each" wins (no double visit).
   Pairs where that config has no state of that name are skipped and logged.
   (Linked display states: names differ per config; shared states: same names everywhere.)
2. For each visit: ShowConfiguration2 (if needed), Configuration.ApplyDisplayState, redraw, then
   check that GetDisplayStates()(0) (the active state) is the one we asked for; otherwise
   log FAILED and skip.
3. Visible bodies (IBody2.Visible via GetBodies2(..., True)), surfaces per option.
   Dedupe key config + body name; repeats are logged "SKIPPED: duplicate ...".
4. Remaining bodies go through the existing ExportBodies (body selection = those names).
5. Cleanup (also on error/cancel): restore original configuration, then original display
   state, then selection.

Assembly: the same plan; each visit applies the assembly display state and then runs the
existing visible-bodies walk (VisibleBodyMap) and ExportAsmPart with a strict body filter.
Dedupe key = part path + referenced config + body. {config} stays the part's configuration.

Naming: new token {ds} (alias {displaystate}). Automatic name = {file}_{ds}_{config}_{body};
_{config} is dropped when that state is exported in one configuration only (or the model has
one configuration). CleanName sanitises as before.

Batch: Display states not offered in Batch (needs per-file picks).


## Revision 2 (Nov 23)
Display states list = one row per state name; scope comes from the "Configurations" drop-down in section 2
(Active configuration only | All configurations (Config filter)). Filters are multi-keyword: commas/semicolons = any,
spaces = all. ApplyDisplayState is always called on a fresh active configuration and its return value is trusted.
See CHANGES-ux-revamp.md, "Revision 2".
