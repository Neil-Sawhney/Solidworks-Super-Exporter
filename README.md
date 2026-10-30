# Solidworks Super Exporter

SolidWorks macro for batch exporting parts, assembly bodies and configurations to STL, STEP, PDF and more.

## Supported formats

| Models (parts and assemblies) | Drawings |
|---|---|
| STL | PDF |
| STEP | eDrawings |
| IGES | DWG |
| Parasolid (.x_t) | DXF |
| ACIS SAT | JPG |
| 3MF | PNG |
| GLB | |
| OBJ | |
| 3D PDF | |
| eDrawings | |
| JPG / PNG | |

## What it does

You pick a mode, tick the items and formats you want, and click Export. It handles:

- each visible body of a part as its own file
- each configuration as its own file
- every body of every ticked configuration
- each visible part of an assembly, once per part and configuration
- every body of each of those parts
- each drawing sheet as its own file, or the ticked sheets combined into one PDF
- a folder of files, or files picked with Add files, optionally including subfolders

For the run, it sets its own STL units, quality and binary/ASCII setting (OBJ uses the same units), STEP AP214, and DXF/DWG "active sheet only". When the run ends, it puts back your previous settings, active configuration, active sheet, body visibility and selection.

## Modes

The modes you see depend on what is open:

| Open document | Modes |
|---|---|
| Part | Bodies, Configurations, Configurations x bodies, Batch |
| Assembly | Configurations, Components, Components x bodies, Batch |
| Drawing | Drawing sheets, Batch |
| Nothing | Batch |

**Bodies** exports each visible solid body of the active configuration. Tick *Surface bodies* to include visible surface bodies too. Hidden bodies are never exported.
- Mesh and image formats (STL, 3MF, GLB, OBJ, JPG, PNG) are saved with every other body hidden.
- STEP, IGES, Parasolid and SAT are saved from a temporary part that holds only that body. The temporary part is closed without saving.
- eDrawings and 3D PDF are written once for the whole model.

**Configurations** exports the whole model once per ticked configuration. The active configuration is ticked by default.

**Configurations x bodies** activates each ticked configuration and exports every visible body in it.

**Components** exports each visible part of the assembly once per part and configuration. Suppressed parts, envelopes, hidden parts, and parts inside a hidden subassembly are left out. The part is exported in the configuration the assembly references.

**Components x bodies** does the same, then exports every visible body of each of those parts. Tick *Surface bodies* to include visible surface bodies.

**Drawing sheets** exports the ticked sheets.
- DWG, DXF, JPG and PNG give one file per sheet.
- PDF gives one file with all ticked sheets when *Drawing PDF: one file with all sheets* is ticked, and one file per sheet when it isn't.
- eDrawings gives one file per drawing.

**Batch** lists every `.sldprt`, `.sldasm` and `.slddrw` in a folder (tick *Subfolders* to include subfolders), all ticked. **Add files...** adds parts, assemblies or drawings from a multi-select dialog. A file already in the list is not added again; it is ticked.
- **Parts in batch** sets how each part is split:
  - all configurations
  - bodies of the active configuration
  - configurations x bodies
- Assemblies are expanded to every part they contain, at every level, and those parts are exported with the same split. Each part is exported once per run. With *Bodies (active config)* that is once per configuration the assembly uses; with the other two settings it is once for the part, however many assemblies contain it.
- Drawings export all their sheets.
- **Config filter** limits configuration export to configurations whose names contain the text. Leave it blank for all of them.
- Files that weren't already open are opened silently and closed without saving afterwards. Files you had open stay open.

The **Filter** box under the item list ticks or unticks matching items (*Select matches* / *Clear matches*), with an optional *Match case* (also used by **Config filter**). It works in every mode.

Every control has a hover tip. The **?** button opens a short summary of the modes, batch options, name tokens and dry run.

## File names

Type a template in **Name**, using these tokens:

| Token | Value |
|---|---|
| `{file}` | Source file name without extension |
| `{config}` | Configuration name |
| `{body}` | Body name |
| `{sheet}` | Sheet name |
| `{component}` | Instance name of the component, when a part is exported from an assembly |
| `{date}` | Today's date, `yyyy-mm-dd` |
| `{rev}` | Custom property `Revision` or `Rev`, configuration-specific first, then file-level |

If you leave Name blank, the template is picked for you: `{file}_{body}`, `{file}_{config}`, `{file}_{config}_{body}` or `{file}_{sheet}`. The configuration is left out for single-configuration parts and in Bodies mode. Characters Windows doesn't allow in file names become `_`, and separators left by empty tokens are removed.

When a part is exported from an assembly, `{file}` in the file name is the part. `{file}` in a subfolder name is the assembly.

**Output** sets where files go:
- *Next to file*
- *Subfolder* (default `{file}_export`, which accepts `{file}` and `{date}`)
- *Custom folder*

For **Components**, **Components x bodies**, and parts reached through an assembly in a batch, *Next to file* and *Subfolder* follow the assembly.

## When a file already exists

**If exists** sets what happens:
- **Auto-number** (default) writes `name_2`, `name_3` and so on.
- **Skip** leaves the existing file alone.
- **Overwrite** replaces it.

Two outputs from the same run never overwrite each other.

## Dry run and log

**Dry run** writes no exports, but still works out and logs every path that would be written. Use it to check names before a big batch.

**Write log CSV** saves `SWExporter_log_<timestamp>.csv` in the output folder, or in the batch folder for batch runs. Columns: time, source, config, body, sheet, format, output, status, errors, warnings.

Progress shows in the SolidWorks status bar and on the form. During a run, Close becomes Cancel, which stops after the current file.

## Installation

1. In SolidWorks, go to **Tools > Macro > New** and save the macro as `SWExporter.swp`. Put `SWExporter.ini` in the same folder.
2. In the VBA editor, choose **Insert > Module** and paste in the contents of `src/SWExporterMain.bas`.
3. Choose **Insert > UserForm** and set its `(Name)` to `frmSWExporter`. Open its code window and paste in the contents of `src/frmSWExporter.txt`. Leave the form empty in the designer, because the controls are created in code.
4. Under **Tools > References**, make sure these are ticked:
   - SldWorks Type Library
   - SOLIDWORKS Constant type library
   - Microsoft Forms 2.0 Object Library
5. Choose **Debug > Compile VBAProject**, then save.
6. Run it from **Tools > Macro > Run**, choosing `SWExporter.swp` and then `main`. You can also put it on a toolbar button.

## Files

| File | Contents |
|---|---|
| `src/SWExporterMain.bas` | Entry point (`main`) and the export code |
| `src/frmSWExporter.txt` | UserForm code |
| `SWExporter.ini` | Format list |

## INI format

```
[MODEL]
"Label", ".ext", FLAG, ON|OFF
[DRAWING]
"Label", ".ext", FLAG, ON|OFF
```

- `MODEL` flags:
  - `MESH` hides the other bodies before saving.
  - `ISOLATE` saves from a temporary part holding one body.
  - `WHOLE` saves the whole model once.
- `DRAWING` flags:
  - `SINGLE` gives one file per drawing.
  - `MULTI` gives one file per sheet.
- `ON` means the format is ticked by default.
- `.e*` becomes `.eprt`, `.easm` or `.edrw` depending on the document type.
- Lines starting with `;` are comments.

## Credit

Inspired by Lenny Kikstra's DocExport macro.
