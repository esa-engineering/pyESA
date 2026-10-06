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
The one exception to "one file per tool" is MEPQTO, which splits a large tool into helper
modules *inside its own bundle folder* (see "Multi-module tools" below). That is still not
a shared library: nothing outside the bundle imports those modules.

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
Non-ASCII characters are avoided inside the `.py` (see "Language" below).

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
   `DWGManage`, `ClassificationTool`, `PointCloudAnalysis`, `MEPQTO`. MEPQTO's variant
   subclasses `System.Windows.Window`, loads the XAML root and copies `Content`, `Title`,
   size and `ResizeMode` onto `self` (`_load_xaml()` in `mepqto_ui.py`), so handlers are
   plain methods on the class.

Resolve the XAML path with `script.get_bundle_file(XAML_FILE_NAME)`, falling back to
`op.join(op.dirname(__file__), XAML_FILE_NAME)`.

For simple prompts prefer the pyRevit built-ins already used everywhere:
`forms.alert`, `forms.SelectFromList`, `forms.ProgressBar`, `forms.WarningBar`,
`forms.CommandSwitchWindow`, `forms.pick_file` / `save_file`.

Note: `forms.alert(..., options=[...])` is not available in every pyRevit version; the
codebase isolates such calls with a `forms.CommandSwitchWindow.show` fallback
(see `ask_link_strategy()` in `TagLinkedRooms_script.py`).

### House style for new XAML windows

New dialogs use approach 2 (raw `XamlReader.Load` from a `<Tool>_ui.py` module beside the
script) and follow the look of `AutoComponents.pushbutton/legend_form.xaml` and
`ElementsInRoom.pushbutton/ElementsInRoom_form.xaml` — copy from either rather than
inventing a new one:

- Default WPF chrome. **No** `Background`, `FontFamily`, `SizeToContent` or ControlTemplates
  on the `Window`; `WindowStartupLocation="CenterScreen"`, `ResizeMode="CanResizeWithGrip"`.
- `Window.Resources` carries the same five named styles — `SectionHeader` (bold 12pt,
  `Foreground="#2D5A8A"`), `FieldLabel`, `InputField` (`Width="70"`, left-aligned),
  `ComboStyle`, `UnitLabel` (gray 11pt) — plus `HintText` for gray 11pt wrapping notes.
- Root `Grid Margin="15"` with rows `Auto / * / Auto`: centred emoji + title TextBlock,
  a `ScrollViewer` holding the sections, then the button row.
- One section per `Border BorderBrush="#CCCCCC" BorderThickness="1" CornerRadius="5"
  Padding="10" Margin="0,0,0,10"`, opened by a `SectionHeader` TextBlock reading
  `emoji + ALL CAPS TITLE`. Informational panels use `BorderBrush="#E8E8E8"
  Background="#F8F8F8"` and gray text.
- Label/field rows are a 2-column `Grid` (fixed-width label column, `*` control column,
  `Margin="0,5"`); sub-options indent with `Margin="20,5,0,0"`.
- Bottom row right-aligned: Cancel (plain) then OK (`Width="100" Height="30"
  IsDefault="True" Background="#2D5A8A" Foreground="White"`).
- Multi-select lists are a plain `ListBox` filled in code with `CheckBox` objects carrying
  the model item in `.Tag` (no `DataTemplate`, no binding), preceded by a `🔍` search row
  and followed by a `Thumb` resize grip and Select All / Deselect All buttons.
- Numeric input is validated all-at-once in `OnOK` via a `(bool, message)` helper, showing
  `MessageBox.Show(...)` and leaving the window open — never with input masks.
- Persist the last-used settings with `script.get_config('ESA_<ToolName>')` /
  `script.save_config()`, wrapped in silent `try/except`.

## Reporting

User-facing results go to the pyRevit output panel, not to `print`:
`output.print_md()`, `output.print_table()`, `output.linkify(element_id)` for clickable
element links, `output.close_others()`. Reports are the primary debugging surface for
these tools — when a tool finds nothing, still emit the report explaining why, rather than
an alert pointing at an empty panel.

Exception: tools whose whole workflow lives in a long-lived window (MEPQTO) show results,
totals and anomalies inside the window (an Issues tab) and in their Excel export, and
deliberately emit **no** pyRevit report on close.

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

**Everything the user sees is in English — no exceptions.** Button titles and tooltips,
`bundle.yaml` (with `it_it:` localization keys where a translation is wanted), `__title__`
and `__doc__`, every XAML label / ToolTip / button caption, every `forms.alert` and
`MessageBox` text, `forms.ProgressBar` titles, and the whole `output.print_md` /
`print_table` report, column headers and status labels included. Constants whose *value* is
a displayed label get English names too (`TO_WRITE = "TO WRITE"`, not
`DA_SCRIVERE = "DA SCRIVERE"`).

Code comments, docstrings and per-tool `README.md` files are mostly Italian — keep writing
those in Italian, matching the file you are editing. The split is: **English out, Italian
in.** Older tools still carry Italian UI strings; translate them when you touch them.

## Multi-module tools: MEPQTO

`pyESA.tab/MEP.panel/MEPQTO.pushbutton` (MEP quantity takeoff driven by Type Mark and price
codes, ~6,800 lines) is the only tool split across several modules. Use it as the template
when a tool outgrows one file. Its `README.md` (Italian) is the functional spec: categories,
measurement formulas, file formats, the full Issues list, merge and concurrency rules.
Read it before changing behaviour, and keep it in sync with the code.

| Module | Role |
| --- | --- |
| `MEPQTO_script.py` | entry point only: document checks, then `show_takeoff_window(doc)` |
| `mepqto_model.py` | the **only** module that reads Revit. `list_links()` lists the link instances; `collect_records()` reads the open model plus the chosen links, only the chosen categories, and reduces elements to plain `InstanceRecord`s (geometry already in mm / m); aggregation, bill, Type Mark summary and issues are pure Python on those records. `CATEGORY_RULES` lists the categories and their measure kind |
| `mepqto_rules.py` | measurement formulas, allowances, duct sheet kg/mq bands, pipe densities. No Revit imports |
| `mepqto_store.py` | price list readers (JSON, .xlsx, .csv), unit aliases, project file, field-by-field merge (`merge_item`) |
| `mepqto_xlsx.py` | minimal .xlsx writer and the export sheets |
| `mepqto_ui.py` + `MEPQTO_form.xaml` | main window (`TakeoffForm`, `Session`) |
| `mepqto_<x>_ui.py` + `MEPQTO_<x>.xaml` | one pair per sub-dialog: `scope` (models, worksets and categories to read, shown before every read; its `_SetPicker` handles both kinds of named sets), `wbs`, `params`, `allowance`, `pricelist` (price list editor) |

Rules that come with the pattern:

- Helper modules live in the bundle folder and are imported by bare name
  (`import mepqto_model as qm`; aliases `qm`, `qr`, `qs`, `qx`). Prefix every module with
  the tool name so it cannot shadow another bundle's module, and never end a helper's name
  with `_script.py` (pyRevit would load it as a command).
- **Read the model once, compute in Python.** `_collect_model()` runs only when phase, phase
  status, parameter map, WBS levels, project file or the scope (links
  and categories, chosen in `mepqto_scope_ui`) change. Category ticks in the main window, units, prices, rules and allowance overrides only call `_refresh()` on the cached
  records. Keep new options on the right side of that line.
- **The model is never modified**: no transactions anywhere in the tool. Everything the user
  types goes to files, not to Revit parameters.
- Display labels that end up in UI and Excel (`ISSUE_*`, `CATEGORY_RULES` labels,
  `WBS_NOT_SET`, `PHASE_STATUS_*`) are English constants in `mepqto_model.py`.
- Its own `get_element_id_value` returns `-1` for `None`: a fourth variant of the helper.
- **Linked models.** Each record carries `source` (index into `CollectResult.docs` /
  `.sources`; 0 = open model) and `ref`, which issues use instead of the bare ElementId:
  an `ElementId` in the open model, a `LinkedId(label, id)` in a link, printed by
  `id_text()`. Element ids of different documents overlap, so per-type and per-host caches
  are per source, and any `GetElement` on a record must use `docs[record.source]`. A link is
  read in the phase with the same name as the host phase, or skipped with an Issue.
  Worksets are excluded **by name** in every source (`CollectOptions.excluded_worksets`,
  counted in `CollectResult.skipped_worksets`): workset ids differ between documents, and
  names keep saved sets valid across projects. In the dialog a tick means *read*; the
  dialog turns the unticked names in the list into the exclusions. To keep a workset
  added later from being silently dropped, it remembers the ticked **and** the already
  seen names (`last_included_worksets`, `last_known_worksets`): unseen names come up
  ticked and flagged. A named set is applied exactly.
- Design options: only the main model and primary options are read
  (`PRIMARY_OPTIONS_ONLY = True` in `mepqto_ui.py`; the UI checkbox was removed). The
  `CollectOptions.primary_only` logic is kept on purpose; the tool README ("Opzioni di
  progetto") explains how to restore it and why "all options at once" double counts.

### Data outside the model

| Data | Where | Written by |
| --- | --- | --- |
| Shared price list | `.json` (`"format": "ESA_MEPQTO_PriceList"`) on a network path, or a read-only `.xlsx` / `.xlsm` / `.csv` | price list editor (`mepqto_pricelist_ui.py`) |
| Project file | `<Model>_MEPQTO.json` next to the **central** model (`default_project_file()`); cloud / unsaved models pick a path, remembered in config | Save button |
| Per-user settings | `script.get_config('ESA_MEPQTO')`: last phase, categories, category set, price list; `project_files` as `"<doc key>::<path>"` and `scope_links` as `"<doc key>::<uid>|<uid>"` strings, both capped at 50; `category_sets` as `"<name>::<key>,<key>"`, `workset_include_sets` as `"<name>::<ws>|<ws>"` (worksets to read; Revit forbids `|` and `:` in names; the older `workset_sets` held exclusions and is ignored); last included / known worksets and workset set | window close; scope and sets as soon as they are chosen / saved |

- Excel / CSV price lists are read **by position**, not by header (`EPU_COLUMNS`: A code,
  B short description, C description, D unit, E unit price, F..I optional). The export's
  EPU sheet uses the same order so it can be read back. `MergedItem.in_price_list` drives
  the red rows of the EPU tab and the "Code not in the price list" issue (only when a
  price list is loaded). Unit aliases (`n`, `nr`, `n°` -> `cad`) live in `UNIT_ALIASES`.
- Window values = empty item, then non-empty price list fields, then non-empty project
  fields. An emptied field falls back to the price list; project values never flow back
  into the shared list.
- **Concurrent saves** (both files sit on shared paths): the store records which
  `(code, field)` pairs and overrides were touched (`_dirty_*`). On save, if the file
  mtime changed since load, it re-reads the disk copy and re-applies only the local edits
  (`ProjectStore._merge_with_disk()`). Rules, WBS levels and the parameter map are saved as
  a block, so last writer wins. The price list editor does the same per code.
- `write_json_file()` writes via a `.tmp` file and escapes non-ASCII by hand (see the
  IronPython `ensure_ascii` trap); reads accept UTF-8 with BOM. A project file that is not
  valid JSON is never overwritten: the user is asked for another one.
- **Adding a field to a price item** (as `short_description` was) touches: `FIELDS`,
  `MergedItem.__slots__` / `__init__` and `EPU_COLUMNS` (positional Excel/CSV layout) in
  `mepqto_store.py`; `COLUMNS`, `FIELD_OF`, `_add_row()` and the import loop in
  `mepqto_pricelist_ui.py` plus its XAML column; `prices_table`, `PRICE_FIELDS`,
  `_write_price_values()` and the search in `mepqto_ui.py` plus the EPU tab column;
  `_price_list_sheet()` in `mepqto_xlsx.py`. Merge, project file and concurrent save
  pick it up from `FIELDS` with no further change.
- Older files must keep loading without conversion: new keys are optional, and the
  `parameters` keys keep their historical names (`piece_codes`, `linear_codes`,
  `linear_include`) even where they no longer describe the content.

### Excel without Excel

Both directions go through `System.IO.Compression` (zip) and XML, with no COM and no
Excel installed: `mepqto_store._read_xlsx_rows()` reads the first sheet (shared strings
included), `mepqto_xlsx.write_workbook()` writes inline strings, fixed styles, formulas
with cached values and `fullCalcOnLoad`. A locked target raises `FileLockedError`. Reuse
these two modules when another tool needs .xlsx I/O instead of introducing COM interop.

### WPF patterns specific to MEPQTO

- Grids are bound to `System.Data.DataTable` (`grid.ItemsSource = table.DefaultView`): .NET
  handles two-way binding, so no IronPython `INotifyPropertyChanged` objects. Search uses
  `DefaultView.RowFilter` with `LIKE`, escaped by `escape_like()`.
- Excel-style column filters come from `mepqto_grid_filter.GridFilters`: it replaces each
  bound column's `Header` with title + funnel button, builds the value popup in code and
  only *returns* a RowFilter fragment (`expression()`); the window ANDs it with its search
  and owns the final `RowFilter`. Set a filtered column's title with `set_title()`, never
  `column.Header =` (that would drop the funnel), and look up columns by header text
  (as `col_unit` does) **before** creating `GridFilters`. Reusable on any DataTable grid.
- Editable numeric cells are **string** columns converted by hand in the table's
  `ColumnChanging` handler, because WPF binding uses the en-US culture and would read
  `12,5` as `125`. `qs.parse_decimal()` accepts both comma and dot.
- A `MessageBox` raised from inside a cell commit is deferred with
  `Dispatcher.BeginInvoke(DispatcherPriority.Background, Action(...))` (`_defer()`).
  Call `_commit_edits()` before Save, Export and close, or the cell being edited is lost.
- `MEPQTO_form.xaml` follows the house style and adds three styles for data grids:
  `GridStyle`, `NumberCell`, `WrapCell`. Copy those for any DataGrid-heavy window.
- Main window: `OnWindowClosing` asks to save if `ProjectStore.is_dirty`, and cancels the
  close on Cancel or on a failed save.

## Development notes worth reading

`pyESA.tab/Utilities.panel/Utilities4.stack/Tag.pulldown/TagLinkedRooms_new.pushbutton/TagLinkedRooms_NOTE-SVILUPPO.md`
is a detailed record of a full redesign: view-range resolution, crop-shape testing,
link transforms, the dry-run-before-transaction pattern, and the open/unverified points.
It is the best reference in the repo for how visibility-dependent tools should be
structured.
