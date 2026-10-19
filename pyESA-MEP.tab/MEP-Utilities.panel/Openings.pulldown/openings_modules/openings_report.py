# -*- coding: utf-8 -*-
"""Openings: New Report e Last Report. File dell'ultimo report, righe e dati del report, sessione di controllo,
azioni della finestra del report (non modale, tramite ExternalEvent), comandi check_main e last_main."""
from openings_check import *  # noqa: F401,F403


def ask_check_scope(ui):
    """Finestra dell'ambito del New Report: tutto il modello, selezione MEP (fori associati agli elementi),
    selezione ARC (fori su muri, solai e strutture). Usa la selezione corrente se c'e', altrimenti si seleziona
    dopo. Ritorna il CheckScope."""
    sel_meps, sel_openings, sel_hosts = selection_for_check()
    kind = ui.ScopeWindow(bundle_path(SCOPE_XAML), last_scope_kind(ui.SCOPE_ALL), len(sel_meps) + len(sel_openings),
                          len(sel_hosts)).show()
    if kind is None:
        script.exit()
    remember_scope_kind(kind)
    openings = existing_openings()
    if kind == ui.SCOPE_MEP:
        if not sel_meps and not sel_openings:
            sel_meps = pick_meps()
        chosen = linked_openings(openings, sel_meps, sel_openings)
        description = "MEP selection ({} MEP elements, {} openings)".format(len(sel_meps), len(sel_openings))
        return CheckScope(kind, description, set(code_of(o) for o in chosen), sel_meps)
    if kind == ui.SCOPE_ARC:
        hosts = sel_hosts or pick_hosts()
        chosen = openings_on_hosts(openings, hosts)
        description = "ARC selection ({} walls, floors and structural elements)".format(len(hosts))
        return CheckScope(kind, description, set(code_of(o) for o in chosen))
    if kind == ui.SCOPE_AREA:
        area = pick_area()
        chosen = [o for o in openings if opening_in_area(area, o)]
        return CheckScope(kind, "Area: {}".format(area_text(area)), set(code_of(o) for o in chosen),
                          meps_in_area(area), area=area)
    return CheckScope(kind, "whole model")


# ----------------------------------------------------------------------------------------------------------------
# Ultimo check (Last Report): stati dal modello, dettaglio per elemento dall'ultimo check fatto su questo computer
# ----------------------------------------------------------------------------------------------------------------
LAST_CHECK_FILE = "ESA_Openings_LastCheck"
ELEMENT_UNCHECKED = "not checked yet"


def last_check_path():
    try:
        return script.get_document_data_file(LAST_CHECK_FILE, "json")
    except Exception:
        return None


def read_last_file():
    """Contenuto del file dell'ultimo report; vuoto se non c'e' o non si legge."""
    path = last_check_path()
    if not path or not op.isfile(path):
        return {}
    try:
        with open(path) as handle:
            data = json.loads(handle.read().lstrip(unichr(0xFEFF)))  # noqa: F821 (IronPython 2: BOM)
        return data if isinstance(data, dict) else {}
    except Exception:
        return {}


def load_last_check():
    """{codice del foro: dettaglio} dell'ultimo controllo di ogni foro."""
    return read_last_file().get("openings", {})


def load_last_report():
    """Ultimo New Report (data, ambito, fori mancanti), o None se non ce n'e'."""
    report = read_last_file().get("report")
    return report if isinstance(report, dict) and report.get("scope") else None


def save_last_check(rows, removed, scope=None, missing_rows=None):
    """Aggiorna il file dell'ultimo report: i fori controllati si sovrascrivono, gli altri restano, quelli
    cancellati si tolgono; con scope, anche l'ultimo report (ambito e fori mancanti) per Last Report. Errori
    silenziosi: il file serve solo a Last Report."""
    path = last_check_path()
    if not path:
        return
    data = read_last_file()
    openings = data.get("openings", {})
    for code in removed:
        openings.pop(code, None)
    stamp = datetime.now().strftime("%d.%m.%Y %H:%M")
    for row in rows:
        openings[row["code"]] = {"status": row["status"], "no_element": row["no_element"], "notes": row["notes"],
                                 "move": row["can_move"], "origin": row["origin"], "date": stamp,
                                 "elements": [[e["id"], e["state"]] for e in row["elements"]]}
    data.update({"format": "ESA_Openings_LastCheck", "openings": openings})
    if scope is not None:
        data["report"] = {
            "date": stamp,
            "scope": {"kind": scope.kind, "description": scope.description,
                      "ids": sorted(scope.ids) if scope.ids is not None else None,
                      "meps": [m.text for m in scope.meps],
                      "released": scope.released, "area": scope.area},
            "missing": [dict((key, row[key]) for key in MISSING_SAVED) for row in (missing_rows or [])]}
    try:
        temp = path + ".tmp"
        with open(temp, "w") as handle:
            handle.write(json.dumps(data))
        if op.isfile(path):
            os.remove(path)
        os.rename(temp, path)
    except Exception:
        pass


# ----------------------------------------------------------------------------------------------------------------
# Righe del report (comuni a New Report e Last Report)
# ----------------------------------------------------------------------------------------------------------------
def check_row(r, origins):
    instance = r["instance"]
    return {"code": r["code"], "instance": instance, "status": WRITTEN[r["state"]], "no_element": r["state"] == OBSOLETE,
            "family": instance.Symbol.Family.Name, "size": size_text(r["old"]),
            "notes": list(r["notes"]) + ([recreate_plan(r)] if r["state"] not in (OK, NEW) else []),
            "elements": r["elements"],
            "can_recreate": bool(r["items"]) and not r["unloaded"] and r["state"] not in (OK, NEW),
            "can_move": alignment_vector(r) is not None, "origin": origins.get(r["code"], "")}


def stored_row(instance, saved):
    """Riga del Last Report: stato e Id dal modello, dettaglio dall'ultimo check su questo computer (se c'e')."""
    code = code_of(instance)
    status = get_text(instance, ctx.status_param) or VALID
    info = saved.get(code)
    states = dict((e[0], e[1]) for e in info.get("elements", [])) if info else {}
    elements, found = [], []
    for text in parse_refs(get_text(instance, ctx.id_param)):
        mep = mep_by_ref(text)
        if mep is None:
            state = ELEMENT_LINK_NOT_LOADED if link_not_loaded(text) else ELEMENT_DELETED
            elements.append({"id": text, "element": None, "name": "Element", "state": state})
        else:
            found.append(mep)
            elements.append({"id": text, "element": mep, "name": mep_name(mep),
                             "state": states.get(text, ELEMENT_UNCHECKED)})
    kind, shape = opening_kind(instance, found)
    if info:
        header = "Last check on this computer: {}.".format(info.get("date", "?"))
        if info.get("status") != status:
            header += " The status in the model ({}) differs from that check ({}): checked again by someone " \
                      "else, or changed since.".format(status, info.get("status"))
        notes = [header] + list(info.get("notes", []))
    else:
        notes = ["Not checked on this computer yet: run New Report for the detail of the elements."]
    return {"code": code, "instance": instance, "status": status,
            "no_element": bool(info.get("no_element")) if info else not found,
            "family": instance.Symbol.Family.Name, "size": size_text(current_data(instance, kind, shape)),
            "notes": notes, "elements": elements, "can_recreate": status != VALID and bool(found),
            "can_move": bool(info.get("move")) if info else False, "origin": info.get("origin", "") if info else ""}


def row_label(row):
    return "Opening {} | {} | {}{}".format(row["code"], row["family"], row["size"],
                                          " | " + row["origin"] if row["origin"] else "")


def status_totals(report):
    """Righe del riepilogo; None = linea di separazione."""
    rows = report.rows

    def count(status):
        return len([r for r in rows if r["status"] == status])

    totals = [("Scope", report.scope.description), ("Openings", len(rows)), None]
    if report.missing_rows is not None:  # None: il report non cerca fori nuovi (li cerca Place&Check)
        totals.append((MISSING, "{} ({} crossings without an opening)".format(
            len(report.missing_rows), sum(m["crossings"] for m in report.missing_rows))))
    totals += [(NOT_ASSOCIATED, "{} ({} with no element left)".format(count(NOT_ASSOCIATED),
                                                                      len([r for r in rows if r["no_element"]]))),
               (TO_REFRESH, count(TO_REFRESH)), (VALID, count(VALID))]
    if report.added:
        totals.append(("Elements added to openings", "{} (crossing an opening, not associated before)".format(
            report.added)))
    return totals


def owners_of(report):
    """{Id elemento (testo): [(codice, stato dell'elemento, Id del foro)]} dei fori del report che lo elencano."""
    owners = {}
    for row in report.rows:
        for e in row["elements"]:
            owners.setdefault(e["id"], []).append((row["code"], e["state"], id_value(row["instance"].Id)))
    return owners


def missing_status(mep, owners, released):
    """Stato di un elemento in una riga Missing openings: se e' ancora elencato da fori presenti a modello, il suo stato
    in quei fori (es. moved out) e i loro Id; se i fori sono stati cancellati, i loro Id; altrimenti no opening."""
    text = mep.text
    links = owners.get(text, [])
    if links:
        states = []
        for _, state, _ in links:
            if state not in states:
                states.append(state)
        return "{} (opening {})".format(" / ".join(states), ", ".join(code for code, _, _ in links))
    if released.get(text):
        return "{} (removed from opening {})".format(ELEMENT_NO_OPENING, ", ".join(released[text]))
    return ELEMENT_NO_OPENING


def report_groups(ui, report):
    groups = []
    owners = owners_of(report)
    if report.missing_rows:
        groups.append((MISSING, [missing_entry(ui, row, owners, report.scope.released)
                                 for row in report.missing_rows]))
    known = (NOT_ASSOCIATED, TO_REFRESH, VALID)
    for status in known:
        groups.append((status, [row_entry(ui, row, owners) for row in report.rows if row["status"] == status]))
    other = [row_entry(ui, row, owners) for row in report.rows if row["status"] not in known]
    if other:
        groups.append(("Other status", other))
    return groups


def row_entry(ui, row, owners):
    """owners: owners_of(report). Selezionare un elemento seleziona anche i fori a cui e' associato. Gli elementi
    dei link non si selezionano per Id: per loro si selezionano solo i fori."""
    select = [id_value(row["instance"].Id)]
    elements = []
    for e in row["elements"]:
        ids = e["element"].select_ids() if e["element"] is not None else []
        select.extend(ids)
        elements.append(("{} {} - {}".format(e["name"], e["id"], e["state"]),
                         ids + [i for _, _, i in owners.get(e["id"], []) if e["element"] is not None]))
    label = row_label(row)
    detail = "{}\nStatus: {}\n\n{}".format(label, row["status"], "\n".join(row["notes"]) or "Elements in place.")
    return ui.OpeningEntry(row["code"], label, detail, select, elements, can_recreate=row["can_recreate"],
                           can_move=row["can_move"], can_delete=row["status"] == NOT_ASSOCIATED)


def link(element, fallback):
    """Link cliccabile all'elemento; solo testo se nel frattempo e' stato cancellato (la finestra del report e'
    non modale: il modello puo' cambiare)."""
    try:
        if element is not None and element.IsValidObject:
            return ctx.output.linkify(element.Id)
    except Exception:
        pass
    return fallback


def report_data(report):
    """Dati del Text Report e dell'export Excel (solo testo: la finestra e l'Excel non leggono Revit). Sezioni per
    stato: Not Associated, To Be Refreshed, Valid (compatti). Ogni voce: codice, titolo, descrizione, note, righe (elemento, categoria, stato dell'elemento)."""
    sections = []
    owners = owners_of(report)
    if report.missing_rows is not None:
        sections.append({"title": MISSING, "empty": "None: every crossing of the elements in this report has an "
                                                    "opening.",
                         "items": [{"code": "-", "title": row["label"],
                                    "description": "{} | {} | crossing {}".format(row["kind_text"], row["size"],
                                                                                  row["host_text"]),
                                    "notes": ["No opening at this crossing: Create in the Interactive Report."],
                                    "rows": [(m.text, mep_name(m),
                                              missing_status(m, owners, report.scope.released))
                                             for m in row["meps"]]}
                                   for row in report.missing_rows]})
    for status in (NOT_ASSOCIATED, TO_REFRESH, VALID):
        items = []
        for row in report.rows:
            if row["status"] != status:
                continue
            items.append({"code": row["code"], "title": row_label(row), "compact": status == VALID,
                          "description": "{} | {}{}".format(row["family"], row["size"],
                                                            " | " + row["origin"] if row["origin"] else ""),
                          "notes": list(row["notes"]),
                          "rows": [(e["id"], e["name"], e["state"]) for e in row["elements"]]})
        sections.append({"title": status, "items": items})
    settings = "Settings: MEC/ELE {}, To Be Refreshed {}, minimum angle {:g} deg, layers joined across {:g} mm.".format(
        "separate" if ctx.settings["separate"] else "together", "on" if ctx.settings["misaligned"] else "off",
        ctx.settings["min_angle_deg"], ctx.settings["layer_gap_mm"])
    return {"title": report.title, "model": ctx.doc.Title, "date": getattr(report, "date", "") or
            datetime.now().strftime("%d.%m.%Y %H:%M"), "summary": report.totals(), "warning": report.warning(),
            "details": [settings] + report.details(), "log": list(report.log), "sections": sections}


def saved_missing_rows(saved):
    """Righe Missing openings del Last Report, dal file: elementi MEP ancora nel modello, niente geometria (Create
    passa prima da un New Report)."""
    rows = []
    for entry in saved or []:
        row = dict((key, entry.get(key)) for key in MISSING_SAVED)
        row["meps"] = [m for m in (mep_by_ref(str(i)) for i in row["mep_ids"] or []) if m is not None]
        row["host"] = None
        rows.append(row)
    return rows


def missing_entry(ui, row, owners, released):
    """owners: owners_of(report); released: scope.released. Gli elementi ancora elencati da fori del modello mostrano
    lo stato in quei fori e i loro Id (selezionati con l'elemento)."""
    select = [i for m in row["meps"] for i in m.select_ids()]  # elementi dei link: non selezionabili per Id
    if row["host"] is not None:
        select.append(id_value(row["host"].Id))  # muro/solaio del modello host (quelli dei link non si selezionano)
    elements = [("{} {} - {}".format(mep_name(m), m.text, missing_status(m, owners, released)),
                 m.select_ids() + [i for _, _, i in owners.get(m.text, [])]) for m in row["meps"]]
    detail = "{}\n\nNo opening at this crossing. Create makes it as Place would (status Valid).".format(row["label"])
    return ui.OpeningEntry(row["key"], row["label"], detail, select, elements, can_recreate=True)


# ----------------------------------------------------------------------------------------------------------------
# New Report e Last Report
# ----------------------------------------------------------------------------------------------------------------
class CheckSession(object):
    """New Report di un ambito: controlla i fori, scrive stati e Id (Id vecchi tolti), salva il dettaglio per il
    Last Report; ricrea e sposta i fori dalla finestra del report."""
    title = "New Report"

    def __init__(self, scope):
        self.scope = scope
        self.symbols, self.missing_families = family_symbols()
        self.missing = set()
        self.log = []  # azioni fatte dalla finestra del report, per il report testuale
        self.origins = {}  # codice di un foro ricreato -> da dove viene (recreated / split / joined from ...)
        self.run([])

    def run(self, removed):
        reset_models()  # link ricaricati mentre la finestra del report era aperta
        # prima gli attraversamenti degli elementi dell'ambito: elementi dentro un foro senza esservi associati
        # (entrano nel suo check, e il foro nell'ambito) e attraversamenti senza foro (Missing openings)
        items, self.scan_stats = scan_crossings(self.scope)
        unlisted = unlisted_crossings(items)
        self.scope.add(unlisted.keys())
        scanned = {}  # riferimento -> attraversamenti: profondita' dei fori (strati aggiunti o tolti)
        for item in items:
            scanned.setdefault(item["mep"].text, []).append(item)
        self.results = sorted([check_opening(o, list(unlisted.get(code_of(o), {}).values()), scanned)
                               for o in self.scope.openings()], key=lambda r: int(r["code"]))
        discipline_checks(self.results)
        self.missing_rows = None  # i fori mancanti li cerca Place&Check: qui solo lo stato dei fori presenti
        self.added = sum(len(r["added"]) for r in self.results)
        self.stale = sum(len(r["stale"]) for r in self.results)  # Id di elementi cancellati, tolti dai fori
        # si scrivono solo i fori (stato, riferimenti, categorie): sui MEP niente
        transaction = DB.Transaction(ctx.doc, "Openings: check")
        transaction.Start()
        try:
            for r in self.results:
                write_opening(r["instance"], r["found_meps"], r["state"], self.missing, r["unloaded"])
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        self.rows = [check_row(r, self.origins) for r in self.results]
        self.date = datetime.now().strftime("%d.%m.%Y %H:%M")
        save_last_check(self.rows, removed, self.scope, self.missing_rows)

    def totals(self):
        return status_totals(self)

    def details(self):
        """Righe in piu' del report testuale."""
        stats = self.scan_stats
        return models_lines() + ["Statuses and element Ids written on the openings, in one transaction (one Undo). "
                                 "MEP elements are not written.",
                "Only the openings already in the model are checked; new openings: Place&Check.",
                "Crossings of the MEP elements in this report: {} inside an opening ({} elements not associated "
                "before, now added), {} running inside a wall or floor (no opening).".format(
                    stats["existing"], self.added, stats["inside"])] + \
               (["Crossings outside the area (not checked): {}.".format(stats["outside_area"])]
                if stats.get("outside_area") else []) + \
               ["- {}".format(p) for p in stats["problems"]] + [
                "Ids of deleted elements removed from the openings: {}.".format(self.stale)]

    def warning(self):
        """Avviso del riepilogo, solo se qualcosa non si e' potuto scrivere."""
        if not self.missing:
            return ""
        return "Not written, parameter missing: {}. Bind these parameters to those categories in the model, " \
               "then run New Report again.".format(missing_text(self.missing))

    def recreate_preview(self, keys):
        lines = []
        for row in self.missing_rows or []:
            if row["key"] in keys:
                lines.append("Create a new opening on {} element(s) crossing {} ({}, {}).".format(
                    len(row["meps"]), row["host_text"], row["kind_text"], row["size"]))
        for r in self.results:
            if r["code"] in keys:
                lines.append("{}: {}".format(r["code"], recreate_plan(r)))
        return "\n".join(lines)

    def recreate(self, keys):
        """Crea i fori mancanti spuntati e ricrea i fori spuntati, in una transazione; poi ricontrolla e scrive."""
        if self.missing_families:
            return "Creating openings needs the families chosen in Settings: {}.".format(
                ", ".join(self.missing_families))
        # elementi usciti dai fori da ricreare (anche dai fori di un gruppo da unire non spuntati): perdono
        # l'associazione ma restano nell'ambito, cosi' la scansione li riporta in Missing openings
        for r in self.results:
            group = [m["code"] for m in r.get("merge_group") or []]
            if r["items"] and (r["code"] in keys or set(group) & set(keys)):
                for mep in outside_meps(r):
                    self.scope.release(mep, [r["code"]])
        transaction = DB.Transaction(ctx.doc, "Openings: create / recreate")
        transaction.Start()
        try:
            created, removed, skipped, events, origins = recreate_openings(self.results, set(keys), self.symbols,
                                                                           self.missing)
            for row in self.missing_rows or []:
                if row["key"] not in keys:
                    continue
                new = code_of(create_opening(row["data"], row["frame"], row["kind"], self.symbols, row["meps"],
                                             NEW, self.missing))
                created.append(new)
                origins[new] = "new"
                events.append("New opening {} on {} ({})".format(
                    new, ", ".join(str(i) for i in row["mep_ids"]), row["host_text"]))
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        for code in removed:
            self.origins.pop(code, None)
        self.origins.update(origins)
        self.scope.replace(removed, created)
        self.run(removed)  # ricontrollo e scrittura degli stati
        lines = list(events) or ["Nothing recreated"]
        if skipped:
            lines.append("Skipped (no element crossing, or link not loaded): {}.".format(", ".join(skipped)))
        message = " ".join(line if line.endswith(".") else line + "." for line in lines)
        self.log.append(message)
        return message

    def move(self, keys):
        moves = [(r, alignment_vector(r)) for r in self.results if r["code"] in keys]
        moves = [(r, v) for r, v in moves if v is not None]
        if not moves:
            return "None of the ticked openings has elements outside it."
        transaction = DB.Transaction(ctx.doc, "Openings: move onto their elements")
        transaction.Start()
        try:
            for r, vector in moves:
                DB.ElementTransformUtils.MoveElement(ctx.doc, r["instance"].Id, vector)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        self.run([])
        message = "Moved onto their elements: {} ({}).".format(
            len(moves), ", ".join("{} by {:.0f} mm".format(r["code"], v.GetLength() * MM_PER_FT) for r, v in moves))
        self.log.append(message)
        return message


class LastReport(object):
    """Last Report: nessun controllo. L'ultimo New Report (stesso ambito e fori mancanti, dal file), con stati e Id
    letti dal modello (scritti dall'ultimo New Report, di chiunque) e il dettaglio per elemento dell'ultimo
    controllo fatto su questo computer."""
    title = "Last Report"

    def __init__(self, saved):
        scope = saved["scope"]
        ids = scope.get("ids")
        meps = [m for m in (mep_by_ref(str(i)) for i in scope.get("meps") or []) if m is not None]
        self.scope = CheckScope(scope.get("kind", "all"), scope.get("description", ""),
                                set(ids) if ids is not None else None, meps, scope.get("released"),
                                scope.get("area"))
        self.date = saved.get("date", "")
        self.log = []
        self.missing_rows = None  # file di versioni precedenti: i fori mancanti li cerca Place&Check
        self.added = 0
        details = load_last_check()
        self.rows = sorted([stored_row(o, details) for o in self.scope.openings()], key=lambda r: int(r["code"]))

    def totals(self):
        totals = status_totals(self)
        totals.insert(1, ("Report of", self.date))
        return totals

    def details(self):
        return ["No check was run: the New Report of {} again, with the statuses as written in the model now. "
                "Create, Recreate, Move and Delete run a New Report on the same scope first.".format(self.date)]

    def warning(self):
        return ""


class ReportRunner(IExternalEventHandler):
    """Esegue nel contesto API di Revit le richieste della finestra del report, che e' non modale: dai suoi pulsanti
    non si puo' chiamare l'API direttamente (selezione, transazioni)."""

    def __init__(self):
        self.jobs = []

    def Execute(self, uiapp):
        jobs, self.jobs = self.jobs, []
        for job in jobs:
            job(uiapp)

    def GetName(self):
        return "ESA Openings report"


class ReportController(object):
    """Azioni della finestra del report (OpeningsCheck_ui.ReportWindow), chiamate con call(nome, argomenti, fatto)
    e messe in coda su un ExternalEvent. Il modello puo' cambiare mentre la finestra e' aperta: Create / Recreate,
    Move e Delete rifanno prima il New Report dell'ambito (da un Last Report e' il primo controllo vero)."""

    def __init__(self, report, ui):
        self.report, self.ui = report, ui
        self.runner = ReportRunner()
        self.event = ExternalEvent.Create(self.runner)  # va creato nel contesto del comando
        # stato del comando che ha aperto la finestra: un altro comando di Openings, lanciato mentre la finestra e'
        # aperta, cambia ctx (Rocket mode, stessi moduli); ogni azione lo rimette com'era
        self.state = ctx.snapshot()

    def call(self, name, args, done):
        """Accoda l'azione name(*args); a fine esecuzione done(risultato, errore). Solo nel modello del report."""
        def job(uiapp):
            ctx.restore(self.state)
            active = uiapp.ActiveUIDocument
            if active is None or not active.Document.Equals(ctx.doc):
                done(None, "switch back to the model '{}' to use this window.".format(ctx.doc.Title))
                return
            try:
                result = getattr(self, name)(*args)
            except Exception as ex:
                done(None, str(ex))
                return
            finally:
                self.state = ctx.snapshot()  # link e riferimenti riletti dall'azione (reset_models)
            done(result, None)

        self.runner.jobs.append(job)
        self.event.Raise()

    # -- azioni (girano nel contesto API) -----------------------------------------------------------------------
    def check_again(self):
        """New Report dell'ambito: un Last Report diventa un check vero, un check si ripete."""
        if isinstance(self.report, CheckSession):
            self.report.run([])
        else:
            self.report = CheckSession(self.report.scope)
            self.report.log.append("New Report run on the scope from the Interactive Report.")

    def groups(self):
        return report_groups(self.ui, self.report)

    def refresh(self):
        self.check_again()
        return "New Report done on the scope: {} openings.".format(len(self.report.rows)), self.groups()

    def select(self, ids, zoom):
        element_ids = List[DB.ElementId]([DB.ElementId(i) for i in ids if ctx.doc.GetElement(DB.ElementId(i))])
        ctx.uidoc.Selection.SetElementIds(element_ids)
        if zoom and element_ids.Count:
            ctx.uidoc.ShowElements(element_ids)
        ctx.uidoc.RefreshActiveView()
        missing = len(ids) - element_ids.Count
        return "Selected {} elements{}.".format(element_ids.Count, " ({} no longer in the model)".format(missing)
                                                 if missing else "")

    def prepare_recreate(self, keys):
        self.check_again()
        return self.report.recreate_preview(keys), self.groups()

    def recreate(self, keys):
        return self.report.recreate(keys), self.groups()

    def move(self, keys):
        self.check_again()
        return self.report.move(keys), self.groups()

    def delete(self, keys):
        """Cancella i fori spuntati (la finestra offre solo i Not Associated), poi New Report dell'ambito: gli Id dei
        fori cancellati si tolgono dagli elementi MEP (CheckScope.removed)."""
        managed = dict((code_of(o), o) for o in existing_openings())
        codes = [k for k in keys if k in managed]
        if not codes:
            return "None of the ticked openings is still in the model.", self.groups()
        for code in codes:  # gli elementi dei fori cancellati restano nell'ambito e nel report
            for text in parse_refs(get_text(managed[code], ctx.id_param)):
                mep = mep_by_ref(text)
                if mep is not None:
                    self.report.scope.release(mep, [code])
        transaction = DB.Transaction(ctx.doc, "Openings: delete")
        transaction.Start()
        try:
            for code in codes:
                ctx.doc.Delete(managed[code].Id)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        self.report.scope.replace(codes, [])
        if isinstance(self.report, CheckSession):
            self.report.run(codes)
        else:
            self.report = CheckSession(self.report.scope)
        message = "Deleted {} openings ({}).".format(len(codes), ", ".join(codes))
        self.report.log.append(message)
        return message, self.groups()

    def report_data(self):
        return report_data(self.report)


def show_report(ui, report):
    """Riepilogo con le due schede, poi Text Report o Interactive Report (non modali: il comando finisce, la finestra
    resta aperta; l'Interactive Report lavora tramite ExternalEvent, vedi __persistentengine__)."""
    choice = ui.SummaryWindow(bundle_path(SUMMARY_XAML), report.title, report.totals(), report.warning()).show()
    if choice == ui.SHOW_TEXT:
        ui.TextReportWindow(bundle_path(TEXT_XAML), report_data(report), revit_owner()).show()
    elif choice == ui.SHOW_WINDOW:
        controller = ReportController(report, ui)
        ui.ReportWindow(bundle_path(REPORT_XAML), controller.groups(), controller, revit_owner(),
                        bundle_path(TEXT_XAML)).show()


def check_main(scope=None):
    """New Report: ambito (finestra, o dato da Place), controllo e scrittura, poi il report."""
    import OpeningsCheck_ui as ui  # finestre dei report, in openings_modules
    if scope is None:
        scope = ask_check_scope(ui)
    show_report(ui, CheckSession(scope))


def last_main():
    """Last Report: l'ultimo New Report del modello, senza controllo e senza chiedere l'ambito."""
    import OpeningsCheck_ui as ui
    saved = load_last_report()
    if saved is None:
        ui.MessageBox.Show("No report yet for this model on this computer: run New Report first.", ui.TITLE)
        script.exit()
    show_report(ui, LastReport(saved))
