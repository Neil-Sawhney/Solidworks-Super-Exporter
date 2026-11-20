# Solidworks Super Exporter

SolidWorks macro for batch exporting parts, assembly bodies, configurations and display states to STL, STEP, PDF and more, from one dialog.

## Install

1. Put `SWExporter.swp` and `SWExporter.ini` in the same folder.
2. In SolidWorks: Tools > Macro > Run and pick `SWExporter.swp`.

Updating the code by hand (Tools > Macro > Edit):
- Module: select all of its code, replace it with `src/SWExporterMain.bas`.
- UserForm `frmSWExporter`: open its code window, replace everything with `src/frmSWExporter.txt`
  (the form stays empty in the designer; every control is created in code).
- Copy `SWExporter.ini` next to the `.swp`. Save the macro.

## Modes

| Document | Mode | What you get |
|---|---|---|
| Part | Each body | One file per ticked visible body (active configuration). Hidden bodies are not listed; surfaces only with "Surface bodies". |
| Part | Each body x configuration | One file per visible body in every ticked configuration. |
| Part | Whole part (per configuration) | The whole part, once per ticked configuration. |
| Assembly | Whole assembly (per configuration) | The whole assembly, once per ticked configuration. |
| Assembly | Each part | Each visible part once (per part + configuration), without bodies hidden in the assembly. |
| Assembly | Each body | One file per ticked visible body. Skips hidden, suppressed and envelope components (and everything under them) and hidden bodies. |
| Part / Assembly | Display states | Every visible body of each ticked display state, one file each, in the active configuration or in each configuration. Repeats are skipped. |
| Drawing | Each sheet | One file per ticked sheet. PDF as one file or one per sheet. |
| Any | Batch (folder or files) | Every listed file. Parts split per "Parts in batch". Assemblies export each part inside. |

### Display states

The list shows each display state twice: `State - active config (Default)` and `State - each configuration`.
Tick the one you want for each state. "Each configuration" means every configuration matching **Config filter**
(blank = all) that has a display state with that name. The count under the list shows how many
configuration x state visits that makes. Use **Dry run** to see every file name first.

Example: one multibody part whose body "Bracket" has several configurations. Make a display state `Main`
(everything except Bracket) and `Bracket` (Bracket only). Then tick `Main - active config` and
`Bracket - each configuration`. You get one file per body in the active configuration, plus one Bracket file per configuration.

Files are named `{file}_{ds}_{config}_{body}`. `_{config}` is dropped when a state is exported in one configuration only.
A body already exported in the same configuration (through another state) is skipped and logged.
The original configuration and display state are restored afterwards. Not available in Batch.

## Form layout

1. **What to export**: mode, one-line description, item list, All / None, ticked count
2. **Scope / filters**: name filter + Match case, Tick / Untick matches, Surface bodies, Config filter, batch folder / files / subfolders / parts in batch (Batch only)
3. **Formats**: format list, drawing PDF one-file, STL/OBJ units, quality, binary
4. **Output & naming**: output location, folder, name template, if exists
5. **Options**: dry run, write log CSV, then **Export** / **Close** bottom-right

Name tokens: `{file} {config} {body} {ds} {sheet} {component} {date} {rev}` (`{displaystate}` = `{ds}`).

Defaults: only STEP AP214 is ticked and the CSV log is off (change the ON/OFF flags in `SWExporter.ini`).
Formats: STL, STEP, IGES, x_t, SAT, 3MF, GLB, OBJ, PDF, eDrawings, DWG, DXF, JPG, PNG.

Inspired by Lenny Kikstra's DocExport. MIT license.
