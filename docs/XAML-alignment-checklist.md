# Allineamento dialoghi allo stile XAML

Obiettivo: ogni finestra di input segue lo stile della casa di `CLAUDE.md`.
Si spunta solo dopo la prova in Revit. Aggiornato al 14.10.2026.

## Come si converte

1. Copiare `MEP-Tools.panel/Stub.pushbutton/Stub_form.xaml` e `Stub_ui.py` come base.
2. Nuovi file `<Tool>_form.xaml` + `<Tool>_ui.py` (nome unico nell'estensione).
3. Controlli nel `_validate()` della finestra, non dopo con `forms.alert`.
4. Ultime scelte con `script.get_config('ESA_<Tool>')`.
5. Testi in inglese, report nel pannello di output.
6. Reload e prova: Cancel, OK vuoto, OK valido, riapertura.

Esempi gia' fatti: i quattro pulsanti di `MEP-Utilities / Insulation`.

## pyESA-MEP.tab

- [x] Tutti i pulsanti visibili provati e funzionanti (14.10.2026), compresi i tre
      Insulation convertiti da FlexForm / WinForms.
- [x] Set Pipe Insulation - DPR412-93: convertito e di nuovo nel `layout`, da provare.
- [x] InsulationTypeMarkToHost: icone aggiunte (copia di quelle del pulldown Insulation).

Facoltativi: ReLevelMEP e MEPAlignToWall da `forms.WPFWindow` a `XamlReader`.

## pyESA.tab

Legenda: **FF** = rpw FlexForm, **rpw** = TaskDialog / TextInput / SelectFromList di rpw,
**SFL** = catena di `forms.SelectFromList`.

### 1. Sostituzioni rapide (rpw -> `forms`)

- [ ] DwgToRevit, JoinUtils, PaintRemove, ExportSchedules: `TaskDialog` -> `forms.alert`
- [ ] SectionBoxRevitLinks, QuantityTakeoff: `TextInput` -> `forms.ask_for_string`
- [ ] DatumExtents: `SelectFromList` / `CommandLink` -> `forms.SelectFromList` / `CommandSwitchWindow`

### 2. FlexForm piccoli -> XAML

- [ ] Coordination / TemplateCleanup (FF, solo checkbox)
- [ ] Utilities / SuperToShared (FF)
- [ ] Utilities / SectionLines (FF)
- [ ] Views-Sheets / CategoriesVisibility (FF + SFL)

### 3. FlexForm con piu' form -> una finestra XAML

- [ ] Utilities / GroupByParameters (2 FF)
- [ ] Utilities / TransferParameter (2 FF)
- [ ] Utilities / ElementsAtLevel (FF, aggiunge la validazione che oggi manca)
- [ ] Utilities / TagElementsInMultipleViews (FF + 6 SFL)
- [ ] Coordination / ModelCleanup (FF)
- [ ] Import-Export / IfcExport (FF)
- [ ] Views-Sheets / RenameFilters&Views (FF)

### 4. Casi grandi

- [ ] Views-Sheets / ChangeTitleblock: finestra WPF scritta in Python -> XAML
- [ ] Coordination / ModelPurge: 4 SFL + CommandSwitchWindow -> una finestra
- [ ] Views-Sheets / FilterColorsManager: 11 SFL + 3 ask_for_string -> una finestra
- [ ] Coordination / ModelReport1: `System.Windows.Forms` file dialog -> `forms.pick_file` / `pick_folder`

### Facoltativi

- WPFWindow -> `XamlReader`: DeleteDuplicates, TagLinkedRooms_new, AddParamsToSchedules, CreateSchedule.
- Da valutare dopo FilterColorsManager: FilterApplyToActiveView, FilterCreate, FiltersCopy.

### Nulla da fare

ModelReport, CountInstances, IsolateWarnings, FilteredSelection, TagInViews,
SumUpDimensions, WorksetViews, BestPractices, e le XAML esistenti (DWGManage,
TabVisibility, ClassificationTool, ElementsInRoom, ModelGroupToElements, HiddenFinder,
PointCloudAnalysis, FamilyNaming, WorksetCreate, AutoComponents, QuickSection).

## Pulizia (quando capita)

- [ ] Backup morti da spostare in `_old`: `SetInsulation_script_v1/_v2.py`,
      `Model Setup_script_*.py` (ModelPurge), `..._withLeaderLenghtWIP.py` (TagElementsInMultipleViews).
      pyRevit non li carica.
- [ ] Layout di Utilities: voce `Utilities7` senza cartella.
- [ ] Insulation: `get_existing_insulation()` lenta su modelli grandi (scorre tutti gli isolamenti per elemento).
