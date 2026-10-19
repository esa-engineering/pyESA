# -*- coding: utf-8 -*-
"""Openings: New Report e Last Report. File dell'ultimo report, righe e dati del report, sessione di controllo,
azioni della finestra del report (non modale, tramite ExternalEvent), comandi check_main e last_main."""
from openings_exchange import *  # noqa: F401,F403


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


def opening_select(row):
    """Cosa selezionare per il foro della riga: il suo Id nel modello aperto, o per un report importato (nessuna
    istanza qui) "opening:<UniqueId>", risolto nel link Openings da ImportController.select."""
    if row["instance"] is None:
        return "opening:" + row.get("unique_id", "")
    return id_value(row["instance"].Id)


def owners_of(report):
    """{Id elemento (testo): [(codice, stato dell'elemento, foro da selezionare)]} dei fori del report che lo
    elencano."""
    owners = {}
    for row in report.rows:
        for e in row["elements"]:
            owners.setdefault(e["id"], []).append((row["code"], e["state"], opening_select(row)))
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
    dei link non si selezionano per Id: per loro si selezionano solo i fori. Report importato (instance None): niente
    spunte ne' azioni, solo Show / Select."""
    select = [opening_select(row)]
    elements = []
    for e in row["elements"]:
        ids = e["element"].select_ids() if e["element"] is not None else []
        select.extend(ids)
        elements.append(("{} {} - {}".format(e["name"], e["id"], e["state"]),
                         ids + [i for _, _, i in owners.get(e["id"], []) if e["element"] is not None]))
    label = row_label(row)
    detail = "{}\nStatus: {}\n\n{}".format(label, row["status"], "\n".join(row["notes"]) or "Elements in place.")
    local = row["instance"] is not None
    return ui.OpeningEntry(row["code"], label, detail, select, elements, can_recreate=row["can_recreate"],
                           can_move=row["can_move"], can_delete=local and row["status"] == NOT_ASSOCIATED,
                           checkable=local)


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
    stato: Not Associated, To Be Refreshed, Valid (compatti). Ogni voce: codice, titolo, descrizione, note, righe
    (elemento, categoria, stato dell'elemento)."""
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
    Last Report; ricrea, sposta e cancella i fori dalla finestra del report.
    run() controlla tutto l'ambito (New Report, Refresh); run_partial() solo i fori su cui agisce un'azione
    dell'Interactive Report, con il loro gruppo di unione e i vicini: gli altri restano come all'ultimo controllo.
    last: LastReport da cui nasce la sessione alla prima azione, senza controllo; i suoi fori restano righe del Last
    Report (stored) finche' un'azione o Refresh non li ricontrolla."""
    title = "New Report"

    def __init__(self, scope, last=None):
        self.scope = scope
        self.symbols, self.missing_families = family_symbols()
        self.missing = set()
        self.log = []  # azioni fatte dalla finestra del report, per il report testuale
        self.origins = {}  # codice di un foro ricreato -> da dove viene (recreated / split / joined from ...)
        self.checked = {}  # codice -> risultato di check_opening, prima delle unioni (discipline_checks)
        self.stored = {}  # codice -> riga del Last Report, foro non ancora ricontrollato in questa sessione
        self.results = []  # risultati di self.checked dopo le unioni, ordinati per codice
        self.scan_stats = None  # statistiche della scansione dell'ultimo controllo completo
        self.missing_rows = None  # i fori mancanti li cerca Place&Check: qui solo lo stato dei fori presenti
        self.since = ""
        if last is None:
            self.run([])
            return
        self.stored = dict((row["code"], row) for row in last.rows)
        self.origins = dict((row["code"], row["origin"]) for row in last.rows if row["origin"])
        self.since = last.date
        self.rows = sorted(self.stored.values(), key=lambda row: int(row["code"]))
        self.added = self.stale = 0
        self.timing = "no check yet (Last Report of {})".format(last.date)
        self.date = last.date

    def run(self, removed):
        """Controllo completo dell'ambito."""
        reset_models()  # link ricaricati mentre la finestra del report era aperta
        started = time.time()
        # prima gli attraversamenti degli elementi dell'ambito: elementi dentro un foro senza esservi associati
        # (entrano nel suo check, e il foro nell'ambito)
        items, self.scan_stats = scan_crossings(self.scope)
        unlisted = unlisted_crossings(items)
        self.scope.add(unlisted.keys())
        scanned = scanned_by_ref(items)
        self.checked = dict((code_of(o), check_opening(o, list(unlisted.get(code_of(o), {}).values()), scanned))
                            for o in self.scope.openings())
        self.stored = {}
        self.finish(set(removed), None, started)

    def run_partial(self, codes, removed=()):
        """Ricontrollo dei soli fori codes, con i fori del loro gruppo di unione, quelli vicini (le unioni li
        riguardano) e quelli dell'ambito mai visti in questa sessione (es. tornati con un Undo). La scansione si
        limita a loro (scan_openings). Gli altri fori restano come all'ultimo controllo: Refresh rifa' tutto."""
        reset_models()
        started = time.time()
        removed = set(removed)
        for code in removed:
            self.checked.pop(code, None)
            self.stored.pop(code, None)
        in_scope = self.scope.openings()
        targets = self.related(codes, in_scope)
        items, _ = scan_openings(self.scope, targets, in_scope)
        unlisted = unlisted_crossings(items)
        scanned = scanned_by_ref(items)
        for opening in targets:
            code = code_of(opening)
            self.checked[code] = check_opening(opening, list(unlisted.get(code, {}).values()), scanned)
            self.stored.pop(code, None)
        self.finish(removed, set(code_of(o) for o in targets), started)

    def related(self, codes, in_scope):
        """Fori dell'ambito da ricontrollare per un'azione su codes (vedi run_partial)."""
        wanted = set(codes)
        for r in self.results:
            if r["code"] in wanted and r.get("merge_group"):
                wanted.update(m["code"] for m in r["merge_group"])
        # vicini: le unioni guardano le sezioni entro unione + 2 franchi, e il box di un foro contiene le sue
        reach = (ctx.settings["merge_mm"] + 2.0 * ctx.settings["clearance_mm"] + MOVE_TOL_MM) / MM_PER_FT
        boxes = [o.get_BoundingBox(None) for o in in_scope if code_of(o) in wanted]
        boxes = [b for b in boxes if b is not None]
        targets = []
        for opening in in_scope:
            code = code_of(opening)
            if code in wanted or (code not in self.checked and code not in self.stored):
                targets.append(opening)
                continue
            box = opening.get_BoundingBox(None)
            if box is not None and any(boxes_touch(box, b, reach) for b in boxes):
                targets.append(opening)
        return targets

    def finish(self, removed, targets, started):
        """Unioni e divisioni su tutti i fori controllati, scrittura, righe del report, file del Last Report.
        targets: codici controllati ora (None: tutti). Degli altri si riscrive solo lo stato, e solo se le unioni lo
        cambiano: i loro riferimenti vengono da un controllo precedente."""
        for table in (self.checked, self.stored):
            for code in list(table):
                if code in removed or not table[code]["instance"].IsValidObject:
                    del table[code]  # cancellato, anche a mano o con un Undo mentre la finestra era aperta
        self.results = []
        for r in self.checked.values():
            copy = dict(r)  # discipline_checks cambia stato, gruppi e note: il risultato del check resta com'e'
            copy["notes"] = list(r["notes"])
            self.results.append(copy)
        self.results.sort(key=lambda r: int(r["code"]))
        discipline_checks(self.results)
        self.added = sum(len(r["added"]) for r in self.results)
        self.stale = sum(len(r["stale"]) for r in self.results)  # Id di elementi cancellati, tolti dai fori
        # si scrivono solo i fori (stato, riferimenti, categorie): sui MEP niente
        written = set()
        transaction = DB.Transaction(ctx.doc, "Openings: check")
        transaction.Start()
        try:
            for r in self.results:
                if targets is None or r["code"] in targets:
                    write_opening(r["instance"], r["found_meps"], r["state"], self.missing, r["unloaded"])
                    written.add(r["code"])
                elif get_text(r["instance"], ctx.status_param) != WRITTEN[r["state"]]:
                    set_text(r["instance"], ctx.status_param, WRITTEN[r["state"]], self.missing)
                    written.add(r["code"])
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        rows = [check_row(r, self.origins) for r in self.results]
        self.rows = sorted(rows + list(self.stored.values()), key=lambda row: int(row["code"]))
        elapsed = time.time() - started
        self.timing = "{:.1f} s, {}".format(elapsed, "whole scope ({} openings)".format(len(self.results))
                                            if targets is None else "{} openings (the ones acted on and their "
                                                                    "neighbours)".format(len(targets)))
        self.date = datetime.now().strftime("%d.%m.%Y %H:%M")
        save_last_check([row for row in rows if row["code"] in written], removed, self.scope, self.missing_rows)

    def totals(self):
        return status_totals(self)

    def details(self):
        """Righe in piu' del report testuale."""
        stats = self.scan_stats
        lines = models_lines() + ["Statuses and element Ids written on the openings, in one transaction (one Undo). "
                                  "MEP elements are not written.",
                                  "Only the openings already in the model are checked; new openings: Place&Check."]
        if stats is not None:
            lines.append("Crossings of the MEP elements in this report (last full check): {} inside an opening ({} "
                         "elements not associated before, now added), {} running inside a wall or floor (no "
                         "opening).".format(stats["existing"], self.added, stats["inside"]))
            if stats.get("outside_area"):
                lines.append("Crossings outside the area (not checked): {}.".format(stats["outside_area"]))
            lines.extend("- {}".format(p) for p in stats["problems"])
        if self.stored:
            lines.append("{} openings not checked again in this window: as in the Last Report of {}. Refresh checks "
                         "the whole scope.".format(len(self.stored), self.since))
        lines += ["Ids of deleted elements removed from the openings: {}.".format(self.stale),
                  "Time: last check {}.".format(self.timing)]
        return lines

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
        self.run_partial(created, removed)  # ricontrollo dei fori nuovi e dei vicini, scrittura degli stati
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
        self.run_partial([r["code"] for r, _ in moves])
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


# ----------------------------------------------------------------------------------------------------------------
# Export (modello Openings) e report importato (modelli MEP / ARC)
# ----------------------------------------------------------------------------------------------------------------
ELEMENT_NOT_HERE = "not in this model any more"


def build_export():
    """Dati del report esportato: **tutti** i fori del modello, non solo lo scope dell'ultimo report (i modelli MEP /
    ARC vedono il quadro intero). Stato dal parametro del foro (scritto dall'ultimo New Report, di chiunque),
    dettaglio degli elementi dall'ultimo controllo su questo computer. Modello di ogni elemento = chiave del link
    (model_name: nome del file centrale, uguale nel modello MEP / ARC), o il titolo di questo modello per i suoi elementi.
    Ritorna (dati, elementi mai controllati su questo computer)."""
    me = model_name(ctx.doc)
    models = {}
    for source in link_sources()[0]:
        models[source.key or me] = {"path": source.path, "transform": transform_list(source.transform),
                                    "phase_found": source.searchable}
    details = load_last_check()
    openings, unchecked = [], 0
    for instance in existing_openings():
        row = stored_row(instance, details)
        box = instance.get_BoundingBox(None)
        elements = []
        for e in row["elements"]:
            model, number = split_ref(e["id"])
            mep = e["element"]
            unchecked += e["state"] == ELEMENT_UNCHECKED
            elements.append({"model": model or me, "id": int(number) if number.isdigit() else number,
                             "ref": e["id"], "unique_id": mep.element.UniqueId if mep is not None else "",
                             "category": e["name"], "state": e["state"]})
        openings.append({"code": row["code"], "unique_id": instance.UniqueId, "status": row["status"],
                         "no_element": row["no_element"], "family": row["family"], "size": row["size"],
                         "origin": row["origin"], "notes": row["notes"][1:],  # senza "Last check on this computer"
                         "box": {"min": xyz_list(box.Min), "max": xyz_list(box.Max)} if box is not None else None,
                         "elements": elements})
    data = {"format": REPORT_FORMAT, "version": REPORT_VERSION, "openings_model": me,
            "openings_path": doc_path(ctx.doc), "date": datetime.now().strftime("%d.%m.%Y %H:%M"),
            "user": ctx.doc.Application.Username, "phase": phase_name(), "units": "ft",
            "status_param": ctx.status_param, "models": models, "openings": openings}
    return data, unchecked


def export_report(folder):
    """Scrive <modello>_Openings.json (sovrascritto: quello che leggono i modelli MEP / ARC) e una copia datata
    <modello>_Openings_AAAAMMGG-HHMM.json come storico. Ritorna il messaggio per la finestra."""
    data, unchecked = build_export()
    name = "{}_Openings".format(model_name(ctx.doc))
    main = op.join(folder, name + ".json")
    write_json_file(main, data)
    write_json_file(op.join(folder, "{}_{}.json".format(name, datetime.now().strftime("%Y%m%d-%H%M"))), data)
    save_model_value(EXPORT_FOLDER_OPTION, folder)
    message = "Exported {} openings to {} (and a dated copy).".format(len(data["openings"]), main)
    if unchecked:
        message += " {} elements were never checked on this computer: run New Report on the whole model for " \
                   "their detail.".format(unchecked)
    return message


def imported_element(e, me):
    """Riga di un elemento di un report importato. Gli elementi di questo modello si cercano per Id e si confermano
    con lo UniqueId (un Id riusato dopo una cancellazione non e' lo stesso elemento)."""
    model = e.get("model") or ""
    number = e.get("id")
    if model != me:
        return {"id": e.get("ref") or ref_text(model, number), "element": None, "name": e.get("category", "Element"),
                "state": e.get("state", "")}
    element = None
    try:
        element = ctx.doc.GetElement(DB.ElementId(int(number)))
    except Exception:
        pass
    if not isinstance(element, MEP_CLASSES) or (e.get("unique_id") and element.UniqueId != e.get("unique_id")):
        return {"id": str(number), "element": None, "name": e.get("category", "Element"),
                "state": "{}; {}".format(e.get("state", ""), ELEMENT_NOT_HERE)}
    return {"id": str(number), "element": mep_ref(element, host_source()), "name": e.get("category", "Element"),
            "state": e.get("state", "")}


def imported_box(entry, inverse):
    """Box del foro nelle coordinate di questo modello (min, max), o None. inverse: inverso del transform con cui
    questo modello e' linkato nel modello Openings."""
    box = entry.get("box")
    if not box or inverse is None:
        return None
    low, high = box["min"], box["max"]
    corners = [inverse.OfPoint(DB.XYZ(x, y, z)) for x in (low[0], high[0]) for y in (low[1], high[1])
               for z in (low[2], high[2])]
    return (DB.XYZ(min(c.X for c in corners), min(c.Y for c in corners), min(c.Z for c in corners)),
            DB.XYZ(max(c.X for c in corners), max(c.Y for c in corners), max(c.Z for c in corners)))


class ImportedReport(object):
    """Report importato in un modello MEP / ARC (Last Report): i fori del file che citano questo modello, con stato e
    note dell'export. Niente controllo e niente azioni sui fori (stanno nel modello Openings): solo Show / Select
    (ImportController). Questo modello si riconosce nel file per nome (model_name: nome del file centrale, uguale
    alla chiave del link nel modello Openings anche se qui e' aperta una copia locale) o, se il nome non c'e', per
    percorso del file centrale."""
    title = "Imported Report"

    def __init__(self, data, path):
        models = data.get("models") or {}
        me = model_name(ctx.doc)
        if me not in models:
            here = (doc_path(ctx.doc) or "").lower()
            me = next((key for key, info in models.items() if here and (info.get("path") or "").lower() == here), me)
        self.data, self.path, self.me = data, path, me
        self.date = data.get("date", "")
        self.known = me in models
        transform = transform_of(models[me].get("transform")) if self.known else None
        inverse = transform.Inverse if transform is not None else None
        self.scope = CheckScope("imported", "openings citing this model ({})".format(me))
        self.log, self.added, self.missing_rows = [], 0, None
        rows = []
        for entry in data.get("openings") or []:
            elements = entry.get("elements") or []
            if not any(e.get("model") == me for e in elements):
                continue
            rows.append({"code": str(entry.get("code", "")), "instance": None,
                         "unique_id": entry.get("unique_id", ""), "status": entry.get("status", ""),
                         "no_element": bool(entry.get("no_element")), "family": entry.get("family", ""),
                         "size": entry.get("size", ""), "origin": entry.get("origin", ""),
                         "notes": list(entry.get("notes") or []),
                         "elements": [imported_element(e, me) for e in elements],
                         "can_recreate": False, "can_move": False, "box": imported_box(entry, inverse)})
        self.rows = sorted(rows, key=lambda r: int(r["code"]) if r["code"].isdigit() else 0)

    def totals(self):
        totals = status_totals(self)
        totals[1:1] = [("Report of", self.date),
                       ("From", "{} ({})".format(self.data.get("openings_model", "?"),
                                                 self.data.get("openings_path") or "path unknown")),
                       ("Exported by", self.data.get("user", "?")), ("File", self.path)]
        return totals

    def details(self):
        return ["Report imported from {}, exported on {} by {}; phase {}. Only the openings citing this model ({}) "
                "are listed. Nothing is checked or written here: openings are created, recreated and checked in "
                "the Openings model.".format(self.data.get("openings_model", "?"), self.date,
                                             self.data.get("user", "?"), self.data.get("phase", "?"), self.me),
                "File: {}. Last Report reopens it from disk every time, so a new Export to the same folder is "
                "read at once.".format(self.path)]

    def warning(self):
        if self.known:
            return ""
        return "This model ({}) is not among the models of the report ({}). The Openings model must link this " \
               "file with the same name.".format(self.me, ", ".join(sorted(self.data.get("models") or {})) or "none")


class NoImportedReport(object):
    """Modello MEP / ARC senza un report importato leggibile: solo il messaggio e il pulsante Import."""
    title = "Imported Report"

    def __init__(self, message):
        self.message = message
        self.rows, self.log = [], []

    def totals(self):
        return [("Report", self.message)]

    def warning(self):
        return ""


def load_imported():
    """Ultimo report importato per questo modello, riletto dal disco (un Export nuovo nella stessa cartella si vede
    subito); NoImportedReport se non c'e' o non si legge."""
    path = model_value(REPORT_FILE_OPTION)
    if not path:
        return NoImportedReport("No report imported for this model yet: use Import report below, on the file that "
                                "the Openings model exports.")
    if not op.isfile(path):
        return NoImportedReport("The last imported report is not there any more: {}. Import it again.".format(path))
    try:
        return ImportedReport(read_report_file(path), path)
    except Exception as ex:
        return NoImportedReport("The last imported report cannot be read ({}): {}".format(path, ex))


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
    e messe in coda su un ExternalEvent. Il modello puo' cambiare mentre la finestra e' aperta: Recreate e Move
    ricontrollano prima i fori spuntati (con gruppo di unione e vicini), e dopo l'azione solo i fori toccati
    (CheckSession.run_partial). Refresh rifa' il New Report di tutto l'ambito."""

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

    def session(self):
        """Sessione di controllo per le azioni. Da un Last Report si crea senza controllare niente: ogni azione
        ricontrolla solo i fori su cui agisce."""
        if not isinstance(self.report, CheckSession):
            self.report = CheckSession(self.report.scope, last=self.report)
        return self.report

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
        self.session().run_partial(keys)
        return self.report.recreate_preview(keys), self.groups()

    def recreate(self, keys):
        return self.session().recreate(keys), self.groups()

    def move(self, keys):
        self.session().run_partial(keys)
        return self.report.move(keys), self.groups()

    def delete(self, keys):
        """Cancella i fori spuntati (la finestra offre solo i Not Associated). Niente ricontrollo: un foro Not
        Associated non entra nelle unioni, quindi gli altri fori non cambiano; si tolgono le sue righe."""
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
        self.session().run_partial([], codes)
        message = "Deleted {} openings ({}).".format(len(codes), ", ".join(codes))
        self.report.log.append(message)
        return message, self.groups()

    def report_data(self):
        return report_data(self.report)


NOT_IN_THIS_MODEL = "not available here: this report was imported from the Openings model, where openings are " \
                    "checked, recreated, moved and deleted."


class ImportController(ReportController):
    """Azioni della finestra di un report importato (modello MEP / ARC): solo Show / Select. Gli elementi di questo
    modello si selezionano per Id; i fori, se il modello Openings e' linkato qui, come riferimenti nel link
    (Selection.SetReferences, Revit 2023+). Senza link si inquadra solo il box del foro."""

    def check_again(self):
        raise ValueError(NOT_IN_THIS_MODEL)

    def refresh(self):
        raise ValueError(NOT_IN_THIS_MODEL)

    def prepare_recreate(self, keys):
        raise ValueError(NOT_IN_THIS_MODEL)

    def recreate(self, keys):
        raise ValueError(NOT_IN_THIS_MODEL)

    def move(self, keys):
        raise ValueError(NOT_IN_THIS_MODEL)

    def delete(self, keys):
        raise ValueError(NOT_IN_THIS_MODEL)

    def openings_links(self):
        """Istanze del link del modello Openings in questo modello (per titolo)."""
        name = self.report.data.get("openings_model")
        return [link for link in DB.FilteredElementCollector(ctx.doc).OfClass(DB.RevitLinkInstance)
                if link.GetLinkDocument() is not None and model_name(link.GetLinkDocument()) == name]

    def select(self, ids, zoom):
        uids = [i[len("opening:"):] for i in ids if isinstance(i, basestring) and i.startswith("opening:")]
        elements = [ctx.doc.GetElement(DB.ElementId(i)) for i in ids if not isinstance(i, basestring)]
        elements = [e for e in elements if e is not None]
        refs = [DB.Reference(e) for e in elements]
        links = self.openings_links()
        found = 0
        for uid in uids:
            for link in links:
                opening = link.GetLinkDocument().GetElement(uid)
                if opening is not None:
                    refs.append(DB.Reference(opening).CreateLinkReference(link))
                    found += 1
                    break
        selection = ctx.uidoc.Selection
        if hasattr(selection, "SetReferences"):
            selection.SetReferences(List[DB.Reference](refs))
        else:  # Revit 2022: solo gli elementi di questo modello
            selection.SetElementIds(List[DB.ElementId]([e.Id for e in elements]))
        if zoom:
            self.zoom_to(elements, uids)
        ctx.uidoc.RefreshActiveView()
        message = "Selected {} elements of this model".format(len(elements))
        if uids:
            message += ", {} of {} openings".format(found, len(uids))
            if not links:
                message += " (the Openings model is not linked here: only zoomed on them)"
        return message + "."

    def zoom_to(self, elements, uids):
        """Inquadra nella vista attiva gli elementi e i box dei fori (salvati nel report, gia' in coordinate di
        questo modello)."""
        lows, highs = [], []
        for element in elements:
            box = element.get_BoundingBox(None)
            if box is not None:
                lows.append(box.Min)
                highs.append(box.Max)
        boxes = dict((r["unique_id"], r["box"]) for r in self.report.rows if r.get("box"))
        for uid in uids:
            if uid in boxes:
                lows.append(boxes[uid][0])
                highs.append(boxes[uid][1])
        if not lows:
            return
        margin = 1.0  # piedi attorno a cio' che si inquadra
        low = DB.XYZ(min(p.X for p in lows) - margin, min(p.Y for p in lows) - margin, min(p.Z for p in lows) - margin)
        high = DB.XYZ(max(p.X for p in highs) + margin, max(p.Y for p in highs) + margin,
                      max(p.Z for p in highs) + margin)
        for view in ctx.uidoc.GetOpenUIViews():
            if view.ViewId == ctx.doc.ActiveView.Id:
                view.ZoomAndCenterRectangle(low, high)


def third_action(ui, holder):
    """Terzo pulsante del riepilogo: Export nel modello Openings, Import nei modelli MEP / ARC (report importato o
    nessun report). holder[0] e' il report mostrato: l'Import lo sostituisce."""
    report = holder[0]
    if not isinstance(report, (ImportedReport, NoImportedReport)):
        def export(folder):
            return {"message": export_report(folder)}
        return ui.ThirdAction(ui.THIRD_EXPORT, "For the MEP / ARC models: choose a folder; writes {}_Openings.json "
                                               "(every opening of the model) and a dated copy.".format(model_name(ctx.doc)),
                              model_value(EXPORT_FOLDER_OPTION), export)

    def do_import(path):
        new = ImportedReport(read_report_file(path), path)
        save_model_value(REPORT_FILE_OPTION, path)
        holder[0] = new
        return {"message": "Imported {}: {} openings cite this model.".format(path, len(new.rows)),
                "title": new.title, "rows": new.totals(), "warning": new.warning()}
    last = model_value(REPORT_FILE_OPTION)
    hint = "Load the report exported from the Openings model (choose the file).{}".format(
        " Last: {}".format(last) if last else "")
    return ui.ThirdAction(ui.THIRD_IMPORT, hint, op.dirname(last) if last else "", do_import)


def show_report(ui, report):
    """Riepilogo con le due schede e il terzo pulsante (Export / Import), poi Text Report o Interactive Report (non
    modali: il comando finisce, la finestra resta aperta; l'Interactive Report lavora tramite ExternalEvent, vedi
    __persistentengine__)."""
    holder = [report]
    imported = isinstance(report, (ImportedReport, NoImportedReport))
    choice = ui.SummaryWindow(bundle_path(SUMMARY_XAML), report.title, report.totals(), report.warning(),
                              third_action(ui, holder), not isinstance(report, NoImportedReport), imported).show()
    report = holder[0]
    if choice == ui.SHOW_TEXT:
        ui.TextReportWindow(bundle_path(TEXT_XAML), report_data(report), revit_owner()).show()
    elif choice == ui.SHOW_WINDOW:
        controller = (ImportController if imported else ReportController)(report, ui)
        ui.ReportWindow(bundle_path(REPORT_XAML), controller.groups(), controller, revit_owner(),
                        bundle_path(TEXT_XAML), readonly=imported).show()


def check_main(scope=None):
    """New Report: ambito (finestra, o dato da Place), controllo e scrittura, poi il report."""
    import OpeningsCheck_ui as ui  # finestre dei report, in openings_modules
    if scope is None:
        scope = ask_check_scope(ui)
    show_report(ui, CheckSession(scope))


def last_main():
    """Last Report. Modello con fori: l'ultimo New Report del modello, senza controllo e senza chiedere l'ambito (se
    su questo computer non ce n'e', tutto il modello con gli stati scritti). Modello senza fori (MEP / ARC): l'ultimo
    report importato, riletto dal disco, e il pulsante Import."""
    import OpeningsCheck_ui as ui
    if not existing_openings():
        show_report(ui, load_imported())
        return
    saved = load_last_report() or {"date": "-", "scope": {"kind": "all", "description":
                                                          "whole model (no New Report on this computer yet)"}}
    show_report(ui, LastReport(saved))
