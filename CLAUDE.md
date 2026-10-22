# CLAUDE.md

Version: 1.0.0 | Last revision: 22.10.2026

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Editing this file

Edit `CLAUDE.md` **only when the user explicitly asks for it**. Never update it on your own
initiative, and never as a follow-up to work on scripts, bundles, XAML, icons or docs, even
when that work makes something here incomplete or out of date. In that case, say so in your
reply and propose the change; apply it only after the user asks.

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

`startup.py` in the repo root runs at every Revit start and every pyRevit Reload, for every
user of the extension. Today it only re-hides the ribbon tabs saved by
`pyESA.tab/Settings.panel/TabVisibility` (config section `ESA_TabVisibility`, key
`hidden_tabs`). Keep it silent and cheap, and wrap everything in `try`: an error there shows
up at Revit start.

Autocomplete for the Revit API requires a local `pyrightconfig.json` pointing at the
RevitAPI stubs (gitignored — each developer creates their own; see README.md).

## Git workflow (mandatory, from README.md)

- Never commit or push directly to `main`.
- Branch naming: `newfeature/<desc>`, `upgrade/<desc>`, `fix/<desc>`, `docs/<desc>`.
- Commit messages: short plain English, no prefixes (`Fix export crash on RVT25`).
- All changes merge to `main` via admin-approved PR. `main` is what end users pull
  through **pyRevit → Update**, so a broken `main` is a broken toolbar for everybody.

## Bundle structure: how the UI is assembled

pyRevit builds the ribbon from directory suffixes, not from any manifest:

```
pyESA.tab/                       tab
  <Name>.panel/                  panel
    <Name>N.stack/               stack (up to 3 buttons in a vertical column)
      <Name>.pushbutton/         button
      lib/                       code shared by the buttons of the stack (optional)
    <Name>.pulldown/             dropdown containing pushbuttons
      lib/                       code shared by the buttons of the pulldown (optional)
    <Name>.urlbutton/            button that opens a hyperlink
```

### Files of a button folder

Every folder holding a script that a ribbon button runs ships the files below. `<Tool>` is the
folder name without the `.pushbutton` suffix.

| File | Content |
| --- | --- |
| `<Tool>_script.py` | the command: header, then calls into the helper modules. Keep it short |
| `<Tool>_Readme.md` | how the tool works: logic, inputs, outputs, limits. Italian |
| `<Tool>_DeveloperNote.md` | log of the attempts made on the tool (see "Developer notes"). Italian |
| `icon.png` | icon for the light theme |
| `icon.dark.png` | icon for the dark theme (see "Icons") |
| `bundle.yaml` | `title`, `tooltip`, `author`, `help_url` (see below). **The filename must be exactly `bundle.yaml`** |
| `<Tool>_<Window>.xaml` | the interface of one window; one file per window |
| `<Tool>_<Window>_ui.py` | behaviour of a window that has logic. Only when needed: a window without logic has none, and one `<Tool>_ui.py` serving several windows is fine when that is enough |
| `<Tool>_<Function>.py` | any further module holding one feature of the tool outside the script |

Data files (`.csv`, `.xlsx`, `.txt`, `.json`) are allowed and read via
`script.get_bundle_file("name.csv")`.

`help_url` stays empty or absent until the wiki page of the tool exists. Whoever edits the
script fills it in when the page is available. Never invent a URL, and never change a
`help_url` that is already there.

`bundle.yaml` on a `.panel`, `.stack`, or `.pulldown` carries a `layout:` list that fixes
the display order of its children. Adding a new button means adding its name to the parent
`bundle.yaml` layout, otherwise ordering is arbitrary.

### Naming

- **`<Tool>`**: PascalCase, no spaces, hyphens or symbols (`RiserUp`, not `Rise&Turn` or
  `Set Pipe Insulation - DPR412-93`). The same name appears in the folder, in the files
  (`<Tool>_script.py`, `<Tool>_Form.xaml`), in the modules (`<Tool>_Form_ui.py` or `<Tool>_ui.py`) and in the
  config section (`ESA_<Tool>`).
- **`<Window>`**: PascalCase name of the window; the main window is `Form`, the others take
  the name of their role (`MEPQTO_Scope.xaml`).
- **`<Function>`**: PascalCase name of the feature. It is never `Script`: pyRevit loads every
  file ending in `_script.py` as a command, so no helper may end that way. The suffix
  `_script.py` is always lowercase.
- **Unique module names.** A module imported by bare name stays in memory between commands
  (Rocket mode), so two modules with the same name and different content would run the wrong
  code. Every module name is unique in the whole extension (the `<Tool>_` prefix guarantees it).

### Shared code: `lib`

When several buttons of the same pulldown or stack share windows, XAML or helpers, those files
live once in a `lib` folder inside that `.pulldown` or `.stack`, not copied into each button.
pyRevit puts the `lib` of any container on the search path of every button below it.

- Name shared files after the group (`Branch_Form.xaml`, `Branch_Form_ui.py`,
  `Riser_Core.py`), with the same rules and the same uniqueness as above.
- A shared XAML is resolved from `os.path.dirname(__file__)` first, because
  `script.get_bundle_file` only knows the bundle of the running command.
- Each button keeps its own script, Readme, DeveloperNote, icons and `bundle.yaml`, and only
  the window, UI and function files that are specific to it.
- Documentation of the shared code lives in `lib/<Group>_Readme.md` and
  `lib/<Group>_DeveloperNote.md`; the button files describe only what is specific to the button.
- `lib` is scoped to its container. Nothing outside it imports from it, and there is no `lib`
  at panel or tab level. `ClashFinder.pulldown/lib` is the model.

### Developer notes (`<Tool>_DeveloperNote.md`)

The file exists so that nobody (human or AI) repeats a solution that was already tried and
did not work.

- Read it **before** changing the tool.
- After **every** change made with AI, add an entry, fixes included. Record the attempts that
  failed or were discarded as carefully as the one that worked.
- Entry (Italian, newest first): date, script version, what was asked, what was tried, result
  (works / fails / partial) and why, final decision.
- It is not the Readme and not the revision history. The Readme says how the tool works now;
  the revision history in the script says what changed in each version; the DeveloperNote says
  how we got there and what does not work. Do not copy one into the other.
- `TagLinkedRooms_NOTE-SVILUPPO.md` is an older note of this kind; keep it as the reference
  for its tool.

### Backups and variants

No `_old` / `_Old` folders and no `_BK01`, `_v1`, `_v2`, `- Copia` files in the repository.
Git history is the archive: tag the version before a major change. Archives outside Git live
on the server, one folder per tool, named `<Tool>_<version>_<date>`. pyRevit **does** load
stray `*_script.py` siblings in a live button folder, so a second `_script.py` variant next to
an active one is a bug. Experimental tools live on a branch, not in a `_hidden` folder on `main`.

### Existing tools: nothing is applied retroactively

The rules in this file apply to **new tools**, and to existing tools only when the user
explicitly asks for it. Do not rename, move, add or delete files of an existing tool on your
own initiative, and do not bring a tool to this layout as a side effect of another change.
Existing tools stay as they are: `script.py` instead of `<Tool>_script.py`, `README.md`
instead of `<Tool>_Readme.md`, no DeveloperNote, folders with spaces or symbols, lowercase
`_form.xaml`, windows copied in several buttons, backups still in the tree, buttons with a
prefixed `bundle.yaml` (`JoinUtils_bundle.yaml`, `PaintRemove_bundle.yaml`,
`CategoriesVisibility_bundle.yaml`, `ReplaceTitleBlocks_bundle.yaml`,
`PointCloudAnalysis_bundle.yaml`, `RenameFilters&Views_bundle.yaml`) which pyRevit does not
read, so those buttons fall back to the folder name for their title.

When you edit an existing tool for another reason, follow the conventions of that tool's own
files. When the user asks to update one to the current layout, the request says how far to
go; renaming a button folder changes the identity of the command in pyRevit, so say so in the
PR description.

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
  themes. Strokes use grey roughly 20…70 for `icon.png` and the reversed ramp, 248…200,
  for `icon.dark.png` (the darkest element of the light icon becomes the lightest element
  of the dark one).
- **Outline style, no fill.** Closed shapes are drawn as outlines with a transparent
  interior, as in the `pyESA` tab (e.g. `MEPConnect`). Never flood a shape with a mid-grey
  (≈ 90…200, the ~139 / ~137 grey of the old `pyESA-MEP` icons): on the ribbon it reads as
  a flat grey block. Only deliberately solid marks (a black tile, a dot, an arrowhead, the
  roll of insulation) are allowed to be filled, in the stroke colour itself.
- The two files are the same drawing, only re-inked. Don't redraw the shape per theme.
- Bold, simple shapes with strokes around 3.5 px in the 96 px grid: the ribbon also renders
  the icon at 32 px and 16 px, so check legibility at those sizes before committing.
- Converting an existing coloured icon: map luminance to the grey ramp above and keep the
  alpha channel, rather than desaturating (a flat desaturation collapses light fills and
  dark outlines into the same mid-grey). Then drop the fill: make the mid-grey interior
  transparent (alpha 0) and keep the outline, blending only the anti-aliased edge pixels so
  no grey halo is left. Do it per theme, since the fill grey differs (e.g. 139 light /
  137 dark, or 199 / 104 on a few icons), and check that thin mid-grey strokes or lines
  (not fills) survive.
- `icon.small.png` is **not** a name pyRevit reads — it is inert. Don't add new ones.

## Tooltips and text inside windows

### Button tooltip (`bundle.yaml`)

```yaml
title: Turn
tooltip: |-
  Adds a new run at the end of a pipe, duct or tray.
  CLICK: opens the settings window, then runs.
  SHIFT + CLICK: runs at once with the last settings.
  Requires: a selection of one run.
author: Name Surname
help_url: <page of the tool on the company wiki>
```

- Line 1: what the tool does, one sentence, about 100 characters, present tense.
- The `CLICK:` and `SHIFT + CLICK:` lines exist only if the tool has both modes (see
  "SHIFT + CLICK"). A tool without SHIFT + CLICK has no such line.
- The `Requires:` line exists only if a selection or a specific view is needed.
- No history, no technical notes, no long warnings: they belong in the guide
  (`<Tool>_Readme.md` and the wiki page behind `help_url`). At most 4 lines.
- English. `it_it:` keys are optional.

### Field tooltips (XAML `ToolTip`)

A field tooltip is written only when it adds what the label does not say: a unit or a limit
("0-90, degrees from the horizontal"), a special value ("0 = merge only when the overpasses
touch"), a side effect ("existing connections are removed"). One sentence, at most 100 to
120 characters. Never repeat the label, never put a tooltip on OK, Cancel, search or clear
icons, and write it once for twin fields (Set A / Set B) or write only what tells them apart.
Use the same terms in every window (run, riser, branch, tee, tap, clash).

### Notes inside windows

A window is for filling in fields, not for explaining the method. Only what is indispensable
stays on screen:

- A `HintText` is allowed only to prevent an error, announce an irreversible effect or clarify
  a unit or sign. One sentence, at most 2 lines at the standard width (about 100 characters).
- At most 2 or 3 hints per window, each right under the field it refers to.
- A live status line (`Now: last junction, same size as the run`) is fine: it informs, it does
  not explain.
- Acceptance test: the window shows everything without a scroll bar (see "No vertical
  scrolling"). A hint that forces a scroll gets shortened or moved to the guide.
- Method, edge cases, criteria and formulas go to `<Tool>_Readme.md` and the wiki.
- `MessageBox` / `forms.alert`: a short title, one or two sentences, and what to do next.

## SHIFT + CLICK

One meaning for every tool:

| Action | Behaviour |
| --- | --- |
| CLICK | opens the settings window and, after OK, runs |
| SHIFT + CLICK | runs at once with the last settings, no window |

- The last settings come from `script.get_config('ESA_<Tool>')`. With no saved settings
  (first use), SHIFT + CLICK behaves like CLICK.
- SHIFT + CLICK never does more, or something different, than CLICK: it does not open another
  tool and does not skip the confirmation of a destructive operation.
- A tool without a window does not read `__shiftclick__` and does not mention it.
- Ctrl + CLICK and Alt + CLICK belong to pyRevit (debug mode, configuration): do not redefine
  them.
- A tool that really needs a different second mode documents it as an exception in its
  tooltip and Readme. Existing tools that behave differently are left as they are until the user
  asks to align them.

## Versions and revision history

Every script has a version `MAJOR.MINOR.PATCH` (a two number version such as `5.1` stays
valid until the first fix).

| Increase | When |
| --- | --- |
| MAJOR | behaviour changes in an incompatible way, or the settings / file formats change |
| MINOR | a new function, field, option or a visible change of the interface |
| PATCH | a fix with no change in expected behaviour |

- Each PR that changes a script updates `Version` and `Date` in `__doc__` and adds a row to the
  revision history block (types: `FEAT`, `UI`, `FIX`, `DOC`, `REFACTOR`), with the PR number.
- The version belongs to the button. Modules in a `lib` have no version of their own: a change
  to a shared module raises the version of every button that uses it.
- Buttons of a same stack or pulldown that share a window keep the same version.
- `Testato su Revit:` lists the versions the tool was actually tried on (see "Revit version
  compatibility"). At least the oldest and the newest available for a new tool or a functional
  change.

When you change a tool, in the same PR:

1. Read its `<Tool>_DeveloperNote.md` first.
2. Raise the version and add the history row.
3. If the change is `FEAT`, `UI` or MAJOR, update `<Tool>_Readme.md` and tell the owner that
   the wiki page behind `help_url` needs the same update. A `FIX` does not touch the guide.
4. Append the entry to `<Tool>_DeveloperNote.md`, failed attempts included (see "Developer
   notes").

## Script conventions

There is **no extension-wide shared library**. Helpers are duplicated per script by design;
when you fix a helper, check whether the same helper exists elsewhere
(`get_element_id_value` alone appears in ~6 files with three slightly different bodies).
Code shared by the buttons of one pulldown or stack goes in that container's `lib` folder,
never in a folder shared by the whole extension (see "Shared code").

The command script stays short. Window logic goes in `<Tool>_<Window>_ui.py` (or in one `<Tool>_ui.py` for several windows), each further feature in
`<Tool>_<Function>.py`. MEPQTO is the template for a large tool (see "Multi-module tools:
MEPQTO"); it keeps no code outside its own bundle folder.

Openings (`pyESA-MEP.tab/MEP-Utilities.panel/Openings.pulldown`) shares one code base between
its buttons in `Openings.pulldown/openings_modules/`, which each button's `script.py` adds to
`sys.path`. That folder is the old form of a `lib`; it stays as it is until the user asks to convert it. Its modules keep no document, output or settings at module
level; they live in `openings_core.ctx`, filled by `start(MODE)` at the start of every command.
Read the tool `README.md` before changing it.

Standard header (what the user sees stays in English, the history block is Italian):

```python
# -*- coding: utf-8 -*-
__title__   = "Button\nLabel"        # \n splits the label across two lines
__doc__     = """Version = 2.3.1
Date    = 21.10.2026
One sentence saying what the tool does.
Author(s): Name
"""
__author__  = "..."
__context__ = "zero-doc"             # only for tools that run without an open model

# ======================= STORICO REVISIONI =======================
# Versione | Data       | Tipo | Descrizione                         | PR
# 2.3.1    | 21.10.2026 | FIX  | Corretto calcolo con tubi inclinati | #47
# 2.3.0    | 02.10.2026 | FEAT | Aggiunta opzione Keep slope         | #44
# Testato su Revit: 2024, 2026
# =================================================================
```

pyRevit injects `__shiftclick__` (bool) at runtime. Its meaning is fixed for every tool, see
"SHIFT + CLICK". Reference it as `__shiftclick__` directly (add `# noqa: F821` if the linter
complains).

Runtime is **IronPython 2.7**: no f-strings anywhere in the repo, `.format()` throughout.
Accented letters only in comments, never in string literals (see "Language").

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

Scripts must stay compatible with **Rocket mode** (pyRevit Settings → Core): it reuses the
IronPython engine between commands and avoids a freeze of a second or two at the end of every
command, so users are expected to turn it on. The command script
is executed fresh each time, but **helper modules imported by bare name stay in `sys.modules`**
across commands. Hence:

- A helper module's name must be unique in the whole extension. Name every helper after its
  tool, or after its group when it lives in a `lib` folder (`<Tool>_<Function>.py`,
  `<Group>_<Function>.py`), so two tools can never ship a different `Turn_ui.py`. Existing
  byte-identical copies (`Branch_ui`, `MatchElevation_ui`, `Riser_ui`, `Select_ui`) are
  tolerated until they are moved into a `lib`; while they exist, edit all copies together.
  When you rename a module, keep the old local name with
  `import ClashOverpass_ui as Overpass_ui` if the script refers to it everywhere.
- No document, UI document or output window at module level in a helper
  (`OUTPUT = script.get_output()`, `doc = revit.doc`): it would stay bound to the first run.
  Fetch them inside the function, or pass them in. Command scripts may keep doing it.
- Module-level caches in helpers would survive across documents: keep caches local or on an
  instance.
- After editing a helper module, pyRevit **Reload** is needed even for a quick test.

Cancellation is `script.exit()`. It raises `SystemExit`, which derives from
`BaseException` and is therefore **not** caught by a trailing `except Exception` — several
scripts rely on that to keep user cancellation out of the error report. Don't wrap it in a
bare `except:`.

### File order and names of new scripts

Order of a script file: 1. header (`coding`, `__title__`, `__doc__`, `__author__`,
`__context__`); 2. revision history block; 3. imports (standard, .NET with `clr`, Revit,
pyRevit); 4. constants and XAML file names; 5. helper functions; 6. classes (windows); 7. main
function; 8. the call to the main function.

| Element | Style | Example |
| --- | --- | --- |
| Functions and variables | `lowercase_with_underscores` | `get_element_id_value`, `run_length` |
| Constants | `UPPERCASE_WITH_UNDERSCORES` | `XAML_FILE_NAME`, `TO_WRITE` |
| Classes | `PascalCase` | `TakeoffForm` |
| XAML controls (`x:Name`) | `PascalCase` ending with the control type | `LengthBox`, `OkButton` |
| Config section | `ESA_<Tool>` | `ESA_MEPQTO` |

Errors are caught with `except Exception`, never with a bare `except:`; transaction names are
English (they show up in Revit's Undo menu).

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
   size and `ResizeMode` onto `self` (`_load_xaml()` in `MEPQTO_ui.py`), so handlers are
   plain methods on the class.

Resolve the XAML path with `script.get_bundle_file(XAML_FILE_NAME)`, falling back to
`op.join(op.dirname(__file__), XAML_FILE_NAME)`.

**Modeless windows** (the user keeps working in Revit while the window is open) exist only in
`ClashFinder.pulldown` (`pyESA-MEP.tab/MEP-Utilities.panel`: Clash Setup, Quick Clash and
Last Report). Its buttons share their modules and XAML through `ClashFinder.pulldown/lib`, which pyRevit
puts on the search path of every button in the pulldown (any container's `lib/` is): the model for every `lib` (see "Shared code"), still
private to that tool. From a modeless window's WPF events there is no
valid API context, so every model read, selection change or transaction goes through
`clashfinder_bridge.RevitBridge`: a queue of callables run by one `ExternalEvent`, created in
the script (API context) and disposed in the window's `Closed`. The script sets
`__persistentengine__ = True` so the engine outlives the command, and helpers resolve their XAML
from `os.path.dirname(__file__)` first, because `script.get_bundle_file` needs the running
command. Copy that bridge (renamed after the tool) for any new modeless window.

For simple prompts prefer the pyRevit built-ins already used everywhere:
`forms.alert`, `forms.SelectFromList`, `forms.ProgressBar`, `forms.WarningBar`,
`forms.CommandSwitchWindow`, `forms.pick_file` / `save_file`.

Note: `forms.alert(..., options=[...])` is not available in every pyRevit version; the
codebase isolates such calls with a `forms.CommandSwitchWindow.show` fallback
(see `ask_link_strategy()` in `TagLinkedRooms_script.py`).

### House style for new XAML windows

New dialogs use approach 2 (raw `XamlReader.Load` from a `<Tool>_<Window>_ui.py` module beside the
script) and follow the look of `AutoComponents.pushbutton/legend_form.xaml` and
`ElementsInRoom.pushbutton/ElementsInRoom_form.xaml` — copy from either rather than
inventing a new one:

- Default WPF chrome. **No** `Background`, `FontFamily`, `SizeToContent` or ControlTemplates
  on the `Window`; `WindowStartupLocation="CenterScreen"`, `ResizeMode="CanResizeWithGrip"`.
- `Window.Resources` carries the same five named styles — `SectionHeader` (bold 12pt,
  `Foreground="#2D5A8A"`), `FieldLabel`, `InputField` (`Width="70"`, left-aligned),
  `ComboStyle`, `UnitLabel` (gray 11pt) — plus `HintText` for gray 11pt wrapping notes.
- Root `Grid Margin="15"` with rows `Auto / * / Auto`: centred emoji + title TextBlock,
  the sections laid out as described in "No vertical scrolling", then the button row.
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
- The window title is the button name in English. `Cancel` has `IsCancel="True"` so Esc closes
  the window, and the tab order follows the visual order, top to bottom.
- One emoji per section header, chosen by concept and the same in every tool (for example
  `📏` sizes and lengths, `🔍` search, `⚙` options).
- File names: `<Tool>_<Window>.xaml` with its `<Tool>_<Window>_ui.py` when the window has logic (see "Naming"). The main
  window is called `Form`.
- Text inside the window follows "Tooltips and text inside windows": no long notes, at most
  2 or 3 hints, and the window never needs a vertical scroll bar (see "No vertical scrolling").

### No vertical scrolling

Every window shows all the functions it offers at once. The user never has to scroll to find
an option, a field or a button.

- No vertical `ScrollViewer` on the window or on its root grid, and no vertical scroll bar on
  the window itself. If the content does not fit, **make the window wider and arrange the
  sections differently**; never turn it into a long single column.
- Put the sections side by side: two or three columns (a root `Grid` with `*` columns, or a
  `WrapPanel` of section `Border`s), with related sections stacked in the same column. The
  reading order is left to right, then top to bottom. Label and field rows can also use two
  columns inside a section.
- Size by hand, since `SizeToContent` is not allowed: set `Width` and `Height` so that all the
  content fits at 100% display scaling and still fits at 125% and 150%. Indicative limits, in
  device independent units: width up to about 1100, height up to about 700, so that the window
  also fits a 1920 x 1080 screen at 125%. Keep `MinWidth` and `MinHeight`, and the resize grip.
- `TabControl` and expanders hide functions, so they are not the answer to a lack of space
  unless the tool owner agrees.
- The only exception is a list or a table whose length depends on the model (categories,
  elements, report rows). It keeps its own internal scroll bar and a fixed visible height. Fields,
  options, hints and buttons are never inside a scrolling area.
- If the tool does not fit even in three columns of about 1100 units, do not fall back to a
  scroll bar: ask the tool owner whether to split it into two buttons or two steps.

## Reporting

**By default a tool produces no final report.** It does its job and ends. Do not print
results, totals or a "nothing found" message to the pyRevit output panel
(`output.print_md()`, `output.print_table()`) for the end user.

A report exists only when the user asks for it before running:

- The tool's window carries a checkbox `Show report at the end`, **unticked at every run**
  (not persisted between runs).
- A tool without a window has no report. If it needs one, give it a window first.
- The report is an **XAML window**, never the output panel, written to be as readable as
  possible for the end user:
  - title with tool name, date and time, model name;
  - a summary line on top: processed, skipped, warnings, errors, each with a symbol as well
    as a colour (colour alone is not enough);
  - a table with one row per element (Id, category, name or type, outcome, reason), sortable,
    with an outcome filter and a text search, long cells wrapped;
  - double click on a row selects the element in Revit. That needs a modeless window and the
    `RevitBridge` / `ExternalEvent` pattern of ClashFinder;
  - `Copy`, `Export` (CSV or Excel) and `Close` buttons;
  - same look as the other windows (see "House style for new XAML windows").
  Models to copy: `ClashFinder_report.xaml`, `OpeningsCheck_report.xaml`. For Excel export
  without Excel installed, copy the writer of `mepqto_xlsx.py` (see "Excel without Excel").
- If something goes wrong (an error, or elements skipped) a short alert is shown even when no
  report was requested: how many elements, the main reason, and a `Show details` button that
  opens the report. When everything worked, nothing is shown.
- Diagnostics for the developer go through `script.get_logger()`, which pyRevit shows in debug
  mode only. Do not leave `print` calls behind.

Existing tools that still write to the output panel stay as they are until the user asks to
convert them. Exception:
tools whose whole workflow lives in a long-lived window (MEPQTO) show results, totals and
anomalies inside the window (an Issues tab) and in their Excel export, and emit **no** report
on close.

## Revit version compatibility

The extension targets **Revit 2022 through 2027, all included**. Every script must run on
every one of them without edits. A tool that truly cannot (a missing API) says so in its
tooltip and its Readme, and exits cleanly instead of failing half way. Old scripts guard with
`if int(doc.Application.VersionNumber) < 2022: script.exit()`; do not add checks that block
versions *newer* than the last one tested, so a tool can still run on a release nobody has
tried yet.

How to write version-proof code:

- Every API difference between versions lives in one small helper function (like
  `get_element_id_value` below), not scattered through the code.
- Test the feature, not the version, when possible: `hasattr(eid, "Value")` is more robust
  than comparing version numbers.
- Never import at module level a name that is missing in some versions: wrap it in
  `try/except` with a fallback.
- Prefer the APIs that already exist in 2022 over the ones being removed: `ForgeTypeId`
  based units and specs (`UnitTypeId`, `SpecTypeId`), `GroupTypeId` for parameter groups
  instead of `BuiltInParameterGroup`, `ElementId` values as `Int64`.
- Revit 2025 and later run on .NET 8: avoid references to libraries tied to .NET Framework.

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

The list of API differences above is indicative: check it against the Revit API release notes
when touching a tool. Compatibility cannot be proven without running the tool, so the revision
history of every script carries a `Testato su Revit:` line with the versions actually tried
(see "Versions and revision history"). The last verified environment is Revit 2026.4 / pyRevit
5.3.1; Revit 2027 has not been verified yet.

Coordinate gotcha, learned the hard way (documented in
`TagLinkedRooms_NOTE-SVILUPPO.md`): `Level.Elevation` can be reported against the Survey
Point, while all geometry (bounding boxes, `LocationPoint`, link `Transform`) is always in
internal coordinates. Mixing the two silently excludes everything. Use
`Level.ProjectElevation` with `Level.Elevation` as fallback when comparing level heights
against geometry.

## Language

**Everything the user sees is in English, no exceptions.** Button titles and tooltips,
`bundle.yaml` (with `it_it:` localization keys where a translation is wanted), `__title__`
and `__doc__`, every XAML label / ToolTip / hint / button caption, every `forms.alert` and
`MessageBox` text, `forms.ProgressBar` titles, the names of Revit transactions (they show up
in the Undo menu) and the whole report window, column headers and status labels included.
Constants whose *value* is a displayed label get English names too (`TO_WRITE = "TO WRITE"`,
not `DA_SCRIVERE = "DA SCRIVERE"`).

Italian is for people who read the source: code comments, docstrings, the revision history
block, `<Tool>_Readme.md` and `<Tool>_DeveloperNote.md`. Keep writing those in Italian,
matching the file you are editing. The split is: **English out, Italian in.** Older tools
still carry Italian UI strings; leave them as they are unless the user asks to translate them.

Other languages:

- Commit messages and branch names: English (see "Git workflow").
- Wiki guides (the page `help_url` points to): Italian.
- Accented letters are allowed in comments only (the file is UTF-8 and declares it). Never put
  them in a string literal: user-facing text is English, and for anything else use an escape
  (`u"è"`), because IronPython 2.7 is fragile with non-ASCII literals.

## Multi-module tools: MEPQTO

`pyESA-MEP.tab/MEP-Utilities.panel/MEPQTO.pushbutton` (MEP quantity takeoff driven by Type Mark and price
codes, ~6,800 lines) is the only tool split across several modules. Use it as the template
when a tool outgrows one file. Its `README.md` (Italian; it keeps its current name) is the functional spec: categories,
measurement formulas, file formats, the full Issues list, merge and concurrency rules.
Read it before changing behaviour, and keep it in sync with the code.

| Module | Role |
| --- | --- |
| `MEPQTO_script.py` | entry point only: document checks, then `show_takeoff_window(doc)` |
| `MEPQTO_model.py` | the **only** module that reads Revit. `list_links()` lists the link instances; `collect_records()` reads the open model plus the chosen links, only the chosen categories, and reduces elements to plain `InstanceRecord`s (geometry already in mm / m); aggregation, bill, Type Mark summary and issues are pure Python on those records. `CATEGORY_RULES` lists the categories and their measure kind |
| `MEPQTO_rules.py` | measurement formulas, allowances, duct sheet kg/mq bands, pipe densities. No Revit imports |
| `MEPQTO_store.py` | price list readers (JSON, .xlsx, .csv, PriMus .xpwe), unit aliases, project file, field-by-field merge (`merge_item`) |
| `MEPQTO_xlsx.py` | minimal .xlsx writer and the export sheets |
| `MEPQTO_ui.py` + `MEPQTO_form.xaml` | main window (`TakeoffForm`, `Session`) |
| `MEPQTO_<x>_ui.py` + `MEPQTO_<x>.xaml` | one pair per sub-dialog: `scope` (models, worksets and categories to read, shown before every read; its `_SetPicker` handles both kinds of named sets), `wbs`, `params`, `allowance`, `pricelist` (price list editor) |

Rules that come with the pattern:

- Helper modules live in the bundle folder and are imported by bare name
  (`import mepqto_model as qm`; aliases `qm`, `qr`, `qs`, `qx`). Prefix every module with
  the tool name so it cannot shadow another bundle's module (MEPQTO keeps its lowercase
  `mepqto_` prefix, grandfathered; new tools use `<Tool>_<Function>.py`), and never end a helper's name
  with `_script.py` (pyRevit would load it as a command).
- **Read the model once, compute in Python.** `_collect_model()` runs only when phase, phase
  status, parameter map, WBS levels, project file or the scope (links
  and categories, chosen in `MEPQTO_scope_ui`) change. Category ticks in the main window, units, prices, rules and allowance overrides only call `_refresh()` on the cached
  records. Keep new options on the right side of that line.
- **The model is never modified**: no transactions anywhere in the tool. Everything the user
  types goes to files, not to Revit parameters.
- Display labels that end up in UI and Excel (`ISSUE_*`, `CATEGORY_RULES` labels,
  `WBS_NOT_SET`, `PHASE_STATUS_*`) are English constants in `MEPQTO_model.py`.
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
  (`PRIMARY_OPTIONS_ONLY = True` in `MEPQTO_ui.py`; the UI checkbox was removed). The
  `CollectOptions.primary_only` logic is kept on purpose; the tool README ("Opzioni di
  progetto") explains how to restore it and why "all options at once" double counts.

### Data outside the model

| Data | Where | Written by |
| --- | --- | --- |
| Shared price list | `.json` (`"format": "ESA_MEPQTO_PriceList"`) on a network path, or a read-only `.xlsx` / `.xlsm` / `.csv` / PriMus `.xpwe` | price list editor (`MEPQTO_pricelist_ui.py`) |
| Project file | `<Model>_MEPQTO.json` next to the **central** model (`default_project_file()`); cloud / unsaved models pick a path, remembered in config | Save button |
| Per-user settings | `script.get_config('ESA_MEPQTO')`: last phase, categories, category set, price list; `project_files` as `"<doc key>::<path>"` and `scope_links` as `"<doc key>::<uid>|<uid>"` strings, both capped at 50; `category_sets` as `"<name>::<key>,<key>"`, `workset_include_sets` as `"<name>::<ws>|<ws>"` (worksets to read; Revit forbids `|` and `:` in names; the older `workset_sets` held exclusions and is ignored); last included / known worksets and workset set | window close; scope and sets as soon as they are chosen / saved |

- PriMus `.xpwe` files (XML, root `PweDocumento`) are a third table format:
  `_read_xpwe_rows()` turns `PweElencoPrezzi/EPItem` into rows with the `XPWE_HEADERS`
  columns (A..I in `EPU_COLUMNS` order, so the default layout fits; J..M `Prezzo2..5`),
  resolving chapters and price book by ID from `PweDGCapitoliCategorie`. Everything else
  (Columns..., layout in the project file, `_rows_to_items`) is the Excel path.
  `PweVociComputo` (the PriMus bill) is not read.
- The **WBS structure** (`qs.load_wbs_structure()`, project file keys
  `wbs_structure_path` / `wbs_structure_layout`) is one sheet with a code / description
  table per WBS level, side by side; the same code can mean different things in two
  levels, so lookups are per level. `qs.WbsLayout` holds sheet, common `first_row` and
  `levels {parameter name: (code col, description col)}`, keyed by **name** so reordering
  the WBS levels keeps the mapping; levels not active now are kept on save. It is edited
  in the same columns window: `ColumnsForm` takes a *kind* (`PriceListKind` /
  `WbsKind(level_names)`: fields, defaults, `guess`, `make`, `check`, `header_row`);
  `guess_wbs_layout()` finds the Codice / Descrittivo header pairs and matches each to
  the level named above it. `qs.WbsStructure.names` is `{name: {code: description}}`;
  the window turns it into `session.wbs_names` `{level index: {code: description}}` for
  the active levels. It only relabels: `pivot_bill(..., wbs_names)` turns WBS group
  labels into "code - description" via `qm.wbs_label()`, while paths, sort keys and item
  cells stay on the code; pivot column headers do the same. Values missing from their
  level's table raise `ISSUE_WBS_NOT_IN_STRUCTURE` (`wbs_structure_issues()`, mapped
  levels only). CME and the export stay on codes.
- Excel / CSV price lists are read through a `PriceListLayout` (sheet + `{field: column
  index}`), chosen in the Price list columns window (`MEPQTO_columns_ui`) and stored in the
  project file as `price_list_layout` (column letters), with the path. No layout = first
  sheet, positional `EPU_COLUMNS` (A code, B short description, C description, D unit,
  E unit price, F..I optional), which is also the order of the export's EPU sheet.
  `guess_layout()` proposes a mapping from known header names. `MergedItem.in_price_list` drives
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
  `MergedItem.__slots__` / `__init__`, `EPU_COLUMNS`, `LAYOUT_FIELDS` and `HEADER_GUESSES` in
  `MEPQTO_store.py`; `COLUMNS`, `FIELD_OF`, `_add_row()` and the import loop in
  `MEPQTO_pricelist_ui.py` plus its XAML column; `prices_table`, `PRICE_FIELDS`,
  `_write_price_values()` and the search in `MEPQTO_ui.py` plus the EPU tab column;
  `_price_list_sheet()` in `MEPQTO_xlsx.py`. Merge, project file and concurrent save
  pick it up from `FIELDS` with no further change.
- Older files must keep loading without conversion: new keys are optional, and the
  `parameters` keys keep their historical names (`piece_codes`, `linear_codes`,
  `linear_include`) even where they no longer describe the content.

### Excel without Excel

Both directions go through `System.IO.Compression` (zip) and XML, with no COM and no
Excel installed: `MEPQTO_store._read_xlsx_rows()` reads the first sheet (shared strings
included), `MEPQTO_xlsx.write_workbook()` writes inline strings, fixed styles, formulas
with cached values and `fullCalcOnLoad`. A locked target raises `FileLockedError`. Reuse
these two modules when another tool needs .xlsx I/O instead of introducing COM interop.

### WPF patterns specific to MEPQTO

- Grids are bound to `System.Data.DataTable` (`grid.ItemsSource = table.DefaultView`): .NET
  handles two-way binding, so no IronPython `INotifyPropertyChanged` objects. Search uses
  `DefaultView.RowFilter` with `LIKE`, escaped by `escape_like()`.
- The Bill of quantities tab is an **outline**, not WPF grouping: `qm.pivot_bill()` (pure
  Python) returns group and item entries for the row levels chosen in the Layout dialog
  (`MEPQTO_layout_ui`; `group_options()`: chapter, subchapter, type mark, unit, `wbs:N`)
  plus the column keys of the WBS levels put on the columns. Pivot columns are rebuilt on
  every fill (`_rebuild_pivot_columns()`, table columns `P1..Pn`, removed and re-added
  with the table empty): items carry quantities, group rows amounts. The window writes
  them as rows of one DataTable with hidden `Kind`, `Gid`, `Anc` ("|1|4|"), `Level`,
  `Shade`, `Toggle` and `Hidden` columns. Collapsing sets `Hidden` on the rows under a
  group (state keyed by the group path, so it survives a recompute); filters run on the
  items and keep the groups listed in their `Anc`. Group totals sit in the Amount column,
  aligned with the items. The Excel export still uses `bill_outline()` (WBS combinations),
  and so does the **CME** tab (`cme_table`, `_fill_cme_table()`), the pre-grouping view kept
  as a flat mirror of the export. Both grids share the allowance-override handlers:
  `_selected_bill_lines()` reads the grid of the selected tab.
- **Manual items** (`qm.ManualItem`, project file key `manual_items`) are rows typed in
  their own tab. They carry only WBS values (by parameter name), Type Mark, code and
  quantity; description, unit and price are the EPU fields of the code (edits go through
  `ProjectStore.set_item_field`). `compute_bill()` keeps them in `bill.manual_quantities`,
  separate from `wbs_quantities`, so `pivot_bill()` / `bill_outline()` emit them as their
  own rows with `manual=True` (hidden `Manual` column -> lavender italics; skipped by the
  allowance override). Their codes join `session.price_codes` (EPU `Source` column).
- Excel-style column filters come from `MEPQTO_grid_filter.GridFilters`: it replaces each
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
It is also the model of a `<Tool>_DeveloperNote.md`. It is the best reference in the repo for how visibility-dependent tools should be
structured.
