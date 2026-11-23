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
| Part / Assembly | Display states | Every visible body of each ticked display state, one file each, in the active configuration or in all (filtered) configurations. Repeats are skipped. |
| Drawing | Each sheet | One file per ticked sheet. PDF as one file or one per sheet. |
| Any | Batch (folder or files) | Every listed file. Parts split per "Parts in batch". Assemblies export each part inside. |

### Display states

The list shows one row per display state. **Configurations** (section 2) decides where each ticked state is exported:

- **Active configuration only** (default): the list shows the active configuration's display states, all ticked.
- **All configurations (Config filter)**: the list shows every display state name found in the configurations
  matching **Config filter** (blank = all). Each ticked state is exported in every matching configuration that has
  a state with that name; configurations without it are skipped and logged. States of the active configuration start ticked.

Changing Configurations, Config filter or Match case rebuilds the list and keeps your ticks by name. The count under
the list shows how many configuration x state runs that makes. Use **Dry run** to see every file name first.

Example: a multibody part whose body "Bracket" has several configurations. With display states linked to configurations,
make `Main` (everything except Bracket) in the default configuration and `Bracket` (Bracket only) in every configuration.
Pick **All configurations** and tick both: you get one file per Main body, plus one Bracket file per configuration.

Files are named `{file}_{ds}_{config}_{body}`. `_{config}` is dropped when a state is exported in one configuration only.
A body already exported in the same configuration (through another state) is skipped and logged.
The original configuration and display state are restored afterwards. Not available in Batch.

### Filters

**Filter** (Tick / Untick matches) and **Config filter** take several keywords: commas or semicolons separate
terms and a name matches if **any** term matches; spaces inside a term mean **all** words must appear.
`bracket, plate` matches either; `top cover` matches `Top_Cover_Plate` and `cover_top`. Match case applies to both.

## Form layout

1. **What to export**: mode, one-line description, item list, All / None, ticked count
2. **Scope / filters**: name filter + Match case, Tick / Untick matches, Surface bodies, Configurations (Display states), Config filter, batch folder / files / subfolders / parts in batch (Batch only)
3. **Formats**: format list, drawing PDF one-file, STL/OBJ units, quality, binary
4. **Output & naming**: output location, folder, name template, if exists
5. **Options**: dry run, write log CSV, then **Export** / **Close** bottom-right

Name tokens: `{file} {config} {body} {ds} {sheet} {component} {date} {rev}` (`{displaystate}` = `{ds}`).

Defaults: only STEP AP214 is ticked and the CSV log is off (change the ON/OFF flags in `SWExporter.ini`).
Formats: STL, STEP, IGES, x_t, SAT, 3MF, GLB, OBJ, PDF, eDrawings, DWG, DXF, JPG, PNG.

Inspired by Lenny Kikstra's DocExport. MIT license.
