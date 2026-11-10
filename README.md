# Solidworks Super Exporter

SolidWorks macro for batch exporting parts, assembly bodies and configurations to STL, STEP, PDF and more, from one dialog.

**Use:** Download `SWExporter.swp` and `SWExporter.ini` into the same folder, then in SolidWorks go to Tools > Macro > Run and pick `SWExporter.swp`.

- Export each body of a part, each configuration, or every body of every configuration
- Export each component of an assembly, once per part and configuration (bodies hidden in the assembly are left out)
- Visible bodies mode: export all and only the visible bodies of a part or assembly, one file per body (skips hidden, suppressed and envelope components, and hidden bodies)
- Export drawing sheets to PDF (one file or one per sheet), DWG and DXF
- Batch a folder (optionally with subfolders) or picked files, including every part inside assemblies
- Formats: STL, STEP, IGES, x_t, SAT, 3MF, GLB, OBJ, PDF, eDrawings, DWG, DXF, JPG, PNG
- Filename tokens, dry run, optional CSV log, hover tips and a ? help button
- Defaults: only STEP is ticked and the CSV log is off (change the ON/OFF flags in `SWExporter.ini`)

Inspired by Lenny Kikstra's DocExport. MIT license.
