# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## What this repo is

The repo root **is** a pyRevit extension. Git clones it as `ESAextensions.extension` into
`%APPDATA%\pyRevit\Extensions\`, and pyRevit discovers it from there. There is no build
step, no package manager, no test suite: the scripts are IronPython source executed
in-process by Revit.

## Running / testing changes

There is no CLI to run. The only way to exercise a change is:

1. Open Revit with a model (most tools need an active document).
2. **pyRevit tab → Reload** to pick up edited scripts, new bundles, or changed `bundle.yaml`.
3. Click the button in the `pyESA` tab.

Syntax can be checked out-of-band without Revit (matches how the codebase is already
validated — see `.claude/settings.local.json`):

```powershell
py -c "import io,sys; src=io.open(sys.argv[1],encoding='utf-8').read(); compile(src, sys.argv[1], 'exec'); print('SYNTAX OK')" <path-to-script.py>
```

This only catches parse errors. Anything touching the Revit API must be tested in Revit.
The most recent verified environment is **Revit 2026.4 / pyRevit 5.3.1**.

Autocomplete for the Revit API requires a local `pyrightconfig.json` pointing at the
RevitAPI stubs (gitignored — each developer creates their own; see README.md).

## Git workflow (mandatory, from README.md)

- Never commit or push directly to `main`.
- Branch naming: `newfeature/<desc>`, `upgrade/<desc>`, `fix/<desc>`, `docs/<desc>`.
- Commit messages: short plain English, no prefixes (`Fix export crash on RVT25`).
- All changes merge to `main` via admin-approved PR. `main` is what end users pull
  through **pyRevit → Update**, so a broken `main` is a broken toolbar for everybody.

## Bundle structure — how the UI is assembled

pyRevit builds the ribbon from directory suffixes, not from any manifest:

```
pyESA.tab/                       tab
  <Name>.panel/                  panel
    <Name>N.stack/               stack (up to 3 buttons in a vertical column)
      <Name>.pushbutton/         button
    <Name>.pulldown/             dropdown containing pushbuttons
    <Name>.urlbutton/            button that opens a hyperlink
```

Inside a `.pushbutton` folder:

- `script.py` **or** `<Anything>_script.py` — the command body. Both naming styles are in
  use across the repo.
- `bundle.yaml` — title / tooltip / author. **The filename must be exactly `bundle.yaml`.**
  Several buttons carry a prefixed variant (`JoinUtils_bundle.yaml`,
  `PaintRemove_bundle.yaml`, `CategoriesVisibility_bundle.yaml`,
  `ReplaceTitleBlocks_bundle.yaml`, `PointCloudAnalysis_bundle.yaml`,
  `RenameFilters&Views_bundle.yaml`); pyRevit does not read those, so those buttons fall
  back to the folder name for their title. If you touch one of those tools, renaming the
  file to `bundle.yaml` is the fix.
- `icon.png` + `icon.dark.png` — light/dark theme icons (96×96). **Mandatory: every button
  ships both files** (see "Icons" below).
- Optional `README.md` per tool, and data files (`.csv`, `.xlsx`, `.txt`) read via
  `script.get_bundle_file("name.csv")`.

`bundle.yaml` on a `.panel`, `.stack`, or `.pulldown` carries a `layout:` list that fixes
the display order of its children. Adding a new button means adding its name to the parent
`bundle.yaml` layout, otherwise ordering is arbitrary.

Folders named `_old` / `_Old` and files suffixed `_BK01`, `_v1`, `_script BK02`,
`- Copia` are dead backups kept in-tree. pyRevit ignores `_old` (no recognized suffix) but
**does** load stray `*_script.py` siblings in a live button folder, so don't leave a
second `_script.py` variant next to an active one.

## Icons

**Every button must ship two icons, one per Revit theme.** A button with only `icon.png`
turns into a dark smudge on the dark ribbon, so the pair is not optional:

| File | Theme | Artwork |
| --- | --- | --- |
| `icon.png` | light | black / dark-grey strokes on transparent background |
| `icon.dark.png` | dark | white / light-grey strokes on transparent background |

Rules:

- 96×96 PNG, 32-bit with alpha, **transparent** background (never a white plate).
- Monochrome, not coloured: one hue is off-brand here and colour never reads on both
  themes. Keep a grey ramp so fills stay distinguishable from outlines — roughly grey
  20…140 for `icon.png` and the reversed ramp, 248…138, for `icon.dark.png` (the darkest
  element of the light icon becomes the lightest element of the dark one).
- The two files are the same drawing, only re-inked. Don't redraw the shape per theme.
- Bold, simple shapes with strokes around 3.5 px in the 96 px grid: the ribbon also renders
  the icon at 32 px and 16 px, so check legibility at those sizes before committing.
- Converting an existing coloured icon: map luminance to the grey ramp above and keep the
  alpha channel, rather than desaturating (a flat desaturation collapses light fills and
  dark outlines into the same mid-grey).
- `icon.small.png` is **not** a name pyRevit reads — it is inert. Don't add new ones.

## Script conventions

Every script is standalone — there is **no shared library module**. Helpers are duplicated
per-script by design; when you fix a helper, check whether the same helper exists elsewhere
(`get_element_id_value` alone appears in ~6 files with three slightly different bodies).

Standard header:

```python
# -*- coding: utf-8 -*-
__title__   = "Button\nLabel"        # \n splits the label across two lines
__doc__     = """Version = 5.1
Date    = 27.08.2026
...
Author(s): Name
"""
__author__  = "..."
__context__ = "zero-doc"             # only for tools that run without an open model
```

pyRevit injects `__shiftclick__` (bool) at runtime — the widespread pattern is a single
command with two modes, documented in the `bundle.yaml` tooltip as `CLICK:` / `SHIFT + CLICK:`.
Reference it as `__shiftclick__` directly (add `# noqa: F821` if the linter complains).

Runtime is **IronPython 2.7**: no f-strings anywhere in the repo, `.format()` throughout.
Non-ASCII characters in comments are frequently avoided (`e'` instead of `è`) in the newer
scripts; keep that where you find it.

Entry points, in order of preference:

```python
from pyrevit import revit, script, DB, forms
doc    = revit.doc          # some scripts use __revit__.ActiveUIDocument.Document
uidoc  = revit.uidoc
output = script.get_output()
```

Transactions: `with revit.Transaction("name"):` is the dominant form (~80 sites); raw
`Transaction(doc, ...)` with explicit `Start()`/`Commit()` appears mostly in the
`XamlReader`-based scripts that don't import `pyrevit.revit`.

Cancellation is `script.exit()`. It raises `SystemExit`, which derives from
`BaseException` and is therefore **not** caught by a trailing `except Exception` — several
scripts rely on that to keep user cancellation out of the error report. Don't wrap it in a
bare `except:`.

## UI: two competing WPF approaches

Both are current; pick whichever the file already uses.

1. **`forms.WPFWindow` subclass** — pyRevit wraps the XAML, names become attributes.
   Used by `ReLevelMEP`, `TagLinkedRooms_new`, `AddParamsToSchedules`, `CreateSchedule`.
   ```python
   class MyWindow(forms.WPFWindow):
       def __init__(self, xaml_path): forms.WPFWindow.__init__(self, xaml_path)
   ```
2. **Raw `XamlReader.Load`** — plain .NET, no pyRevit dependency; requires
   `clr.AddReference` for `PresentationFramework` / `PresentationCore` / `WindowsBase`
   and manual `FindName()` for every control. Used by `HiddenFinder`, `ModelReport1`,
   `DWGManage`, `ClassificationTool`, `PointCloudAnalysis`.

Resolve the XAML path with `script.get_bundle_file(XAML_FILE_NAME)`, falling back to
`op.join(op.dirname(__file__), XAML_FILE_NAME)`.

For simple prompts prefer the pyRevit built-ins already used everywhere:
`forms.alert`, `forms.SelectFromList`, `forms.ProgressBar`, `forms.WarningBar`,
`forms.CommandSwitchWindow`, `forms.pick_file` / `save_file`.

Note: `forms.alert(..., options=[...])` is not available in every pyRevit version; the
codebase isolates such calls with a `forms.CommandSwitchWindow.show` fallback
(see `ask_link_strategy()` in `TagLinkedRooms_script.py`).

## Reporting

User-facing results go to the pyRevit output panel, not to `print`:
`output.print_md()`, `output.print_table()`, `output.linkify(element_id)` for clickable
element links, `output.close_others()`. Reports are the primary debugging surface for
these tools — when a tool finds nothing, still emit the report explaining why, rather than
an alert pointing at an empty panel.

## Revit version compatibility

The extension targets **Revit 2022 through 2026**. Older scripts guard with
`if int(doc.Application.VersionNumber) < 2022: script.exit()`.

`ElementId.IntegerValue` was removed in 2026 (`.Value`, an `Int64`, replaces it). Use the
existing helper shape:

```python
def get_element_id_value(eid):
    if hasattr(eid, "Value"):
        return eid.Value          # Revit 2026+
    return eid.IntegerValue       # Revit <= 2025
```

`ModelReport_script.py` still uses `.IntegerValue` directly in many places and is not
2026-safe.

Coordinate gotcha, learned the hard way (documented in
`TagLinkedRooms_NOTE-SVILUPPO.md`): `Level.Elevation` can be reported against the Survey
Point, while all geometry (bounding boxes, `LocationPoint`, link `Transform`) is always in
internal coordinates. Mixing the two silently excludes everything. Use
`Level.ProjectElevation` with `Level.Elevation` as fallback when comparing level heights
against geometry.

## Language

Button titles and tooltips are English, sometimes with `it_it:` localization keys in
`bundle.yaml`. Code comments and per-tool `README.md` files are mostly Italian. Match the
file you are editing.

## Development notes worth reading

`pyESA.tab/Utilities.panel/Utilities4.stack/Tag.pulldown/TagLinkedRooms_new.pushbutton/TagLinkedRooms_NOTE-SVILUPPO.md`
is a detailed record of a full redesign: view-range resolution, crop-shape testing,
link transforms, the dry-run-before-transaction pattern, and the open/unverified points.
It is the best reference in the repo for how visibility-dependent tools should be
structured.
