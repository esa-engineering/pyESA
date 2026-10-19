# -*- coding: utf-8 -*-
"""Openings: comando Place (fori nuovi sugli attraversamenti scelti; gli elementi gia' associati vanno al check)."""
from openings_report import *  # noqa: F401,F403


SKIPPED_LINKS_OPTION = "place_skipped_links"  # ["<modello>::<chiave>|<chiave>"], link non spuntati per modello
SKIPPED_LINKS_MAX = 50


def model_key():
    """Chiave del modello aperto nella configurazione: percorso del centrale o del file, altrimenti il titolo."""
    return doc_path(ctx.doc) or ctx.doc.Title


def load_skipped_links():
    try:
        entries = script.get_config(CONFIG_NAME).get_option(SKIPPED_LINKS_OPTION, []) or []
    except Exception:
        return set()
    prefix = model_key() + "::"
    for entry in entries:
        if entry.startswith(prefix):
            return set(k for k in entry[len(prefix):].split("|") if k)
    return set()


def save_skipped_links(skipped):
    """Ricorda i link non spuntati di questo modello (i link nuovi compaiono spuntati). Errori silenziosi."""
    try:
        config = script.get_config(CONFIG_NAME)
        prefix = model_key() + "::"
        entries = [e for e in (config.get_option(SKIPPED_LINKS_OPTION, []) or []) if not e.startswith(prefix)]
        setattr(config, SKIPPED_LINKS_OPTION, ([prefix + "|".join(sorted(skipped))] + entries)[:SKIPPED_LINKS_MAX])
        script.save_config()
    except Exception:
        pass


def skip_links(skipped):
    """Da qui in poi Place non legge questi link (MEP, muri, solai e strutture)."""
    ctx.skipped_links = frozenset(skipped)
    reset_models()


def missing_families_exit(missing_families):
    forms.alert("Load the opening families first (or choose others in Settings):\n{}".format(
        "\n".join(missing_families)), title=ctx.title, exitscript=True)


def show_result(header, message, summary, openings=(), ids=(), warning="", details=()):
    """Esito in una finestra (niente pannello di output di pyRevit); "Select new openings" li seleziona e li mostra."""
    import OpeningsPlace_ui
    if OpeningsPlace_ui.ResultWindow(bundle_path(RESULT_XAML), header, message, summary, openings, ids, warning,
                                     details).show():
        element_ids = List[DB.ElementId]([DB.ElementId(i) for i in ids])
        ctx.uidoc.Selection.SetElementIds(element_ids)
        ctx.uidoc.ShowElements(element_ids)


def place_main():
    import OpeningsPlace_ui  # finestre di Place&Check, in openings_modules
    symbols, missing_families = family_symbols()
    meps, hosts, ignored, area = [], [], 0, None
    if ctx.uidoc.Selection.GetElementIds().Count:
        # selezione: gli elementi gia' associati a un foro vanno al check, gli altri ai fori nuovi
        openings, linked, meps, hosts, ignored = split_selection()
        if openings or linked:
            chosen = linked_openings(existing_openings(), linked, openings)
            check_main(CheckScope("mep", "selection, already associated ({} openings, {} MEP elements)".format(
                len(openings), len(linked)), set(code_of(o) for o in chosen), linked))
        if not meps and not hosts:
            if not openings and not linked:
                show_result("NOTHING TO DO", "None of the {} selected elements is a pipe, duct, cable tray, "
                            "conduit, opening, wall, floor or structural element.".format(ignored), [])
            return
        if missing_families:
            missing_families_exit(missing_families)
        skip_links(load_skipped_links())  # senza finestra: i link scelti l'ultima volta (il check sopra li legge tutti)
    else:
        if missing_families:
            missing_families_exit(missing_families)
        links, remembered = loaded_links(), load_skipped_links()
        window = OpeningsPlace_ui.PickWindow(bundle_path(PICK_XAML), last_pick(),
                                             "{}\n{}".format(settings_line(), family_line()), links, remembered)
        pick = window.show()
        if pick is None:
            script.exit()
        remember_pick(pick)
        # i link tolti che ora non sono caricati restano tolti quando tornano
        save_skipped_links(window.skipped | (remembered - set(key for key, _ in links)))
        skip_links(window.skipped)
        if pick == OpeningsPlace_ui.PICK_MEPS:
            meps = pick_meps()
        elif pick == OpeningsPlace_ui.PICK_AREA:
            area = pick_area()
        else:
            hosts = pick_hosts()

    stats = {"inside": 0, "existing": 0, "extended": 0, "outside_area": 0, "problems": []}
    items, modes = [], []
    if area is not None:
        # ogni elemento MEP dell'area contro ogni muro, solaio e struttura dell'area (anche nei link); fori solo per
        # gli attraversamenti dentro l'area, anche se l'elemento MEP ne esce
        area_meps, area_hosts = meps_in_area(area), hosts_in_area(area)
        if area_meps and area_hosts:
            with forms.ProgressBar(title="Openings: looking for crossings ({value} of {max_value})") as bar:
                items.extend(find_crossings(area_hosts, stats, only=set(m.text for m in area_meps),
                                            progress=lambda i, n: bar.update_progress(i + 1, n), area=area))
        modes.append("area: {}; MEP elements {}, walls, floors and structural elements {}".format(
            area_text(area), len(area_meps), len(area_hosts)))
    if hosts:
        items.extend(find_crossings(hosts, stats))
        modes.append("walls, floors and structural elements selected: {}".format(len(hosts)))
    if meps:
        # gli elementi gia' scelti sopra hanno gia' dato tutti i loro attraversamenti
        chosen = set(h[0] for h in hosts)
        found = [h for h in hosts_from_meps(meps) if h[0] not in chosen]
        items.extend(find_crossings(found, stats, only=set(m.text for m in meps)))
        modes.append("from {} MEP elements: walls, floors and structural elements found {}".format(
            len(meps), len(found)))
    mode_text = "; ".join(modes)
    items = merge_overlapping(items)
    groups = clusters(items)
    # righe comuni dell'esito: conteggi della ricerca e dettagli (impostazioni, modelli letti, problemi)
    summary = [("Mode", mode_text), ("Crossings found", len(items)),
               ("Already with an opening", stats["existing"]), ("Running inside (no opening)", stats["inside"]),
               ("Extended through layers", stats["extended"])]
    if stats["outside_area"]:
        summary.append(("Outside the area (skipped)", stats["outside_area"]))
    if ignored:
        summary.append(("Selected elements not used", "{} (not MEP, opening, wall, floor or structure, or not "
                                                       "straight)".format(ignored)))
    details = ["Settings: {}.".format(settings_line()), family_line()] + models_lines()
    problems = "\n".join(stats["problems"])
    if not groups:
        show_result("NO NEW OPENINGS NEEDED", "Every crossing found already has an opening, or none was found.",
                    summary, warning="Not read:\n" + problems if problems else "", details=details)
        return
    preview = {"mode": mode_text, "crossings": len(items), "openings": len(groups), "existing": stats["existing"],
               "inside": stats["inside"], "extended": stats["extended"]}
    if not OpeningsPlace_ui.ConfirmWindow(bundle_path(CONFIRM_XAML), preview, stats["problems"]).show():
        script.exit()

    missing = set()
    rows, ids = [], []
    mep_codes = {}  # riferimento MEP -> (MepRef, codici dei fori nuovi); si scrivono solo i MEP del modello
    managed = set(code_of(o) for o in existing_openings())  # sui MEP si tengono solo i codici di fori esistenti
    transaction = DB.Transaction(ctx.doc, ctx.title)
    transaction.Start()
    try:
        for group in groups:
            frame, kind = group[0]["frame"], group[0]["kind"]
            data = sizes(group, frame)
            meps = []
            for item in group:
                if all(m != item["mep"] for m in meps):
                    meps.append(item["mep"])
            instance = create_opening(data, frame, kind, symbols, meps, NEW, missing)
            code = code_of(instance)
            for mep in meps:
                mep_codes.setdefault(mep.text, (mep, []))[1].append(code)
            key, host, model = group[0]["host"]
            ids.append(id_value(instance.Id))
            rows.append(" | ".join([code, "{} {}".format(kind, data["shape"]), size_text(data),
                                    disciplines_text(meps) or "-",
                                    "{} {} ({})".format(host.Category.Name, id_value(host.Id), model),
                                    ", ".join(m.text for m in meps)]))
        for mep, new_codes in [v for v in mep_codes.values() if v[0].is_host]:
            old = [c for c in split_list(get_text(mep.element, ctx.id_param)) if c in managed]
            # con altri fori gia' presenti resta il loro stato, se peggiore
            status = worst([VALID] + ([get_text(mep.element, ctx.status_param) or VALID] if old else []))
            write_mep(mep.element, old + [c for c in new_codes if c not in old], status, missing)
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    linked = len([v for v in mep_codes.values() if not v[0].is_host])
    written = "{} = {} on the new openings{}.".format(
        ctx.status_param, VALID, "" if linked == len(mep_codes) else " and on their MEP elements in this model")
    if linked:
        written += " MEP elements of linked models ({}) are not written: their link is on the opening ({}).".format(
            linked, ctx.id_param)
    warnings = (["Parameters not found (not written): {}.".format(missing_text(missing))] if missing else []) + \
        (["Not read:\n" + problems] if problems else [])
    show_result("OPENINGS CREATED", "{} new openings, in one transaction (one Undo).".format(len(groups)),
                [("Openings created", len(groups))] + summary, rows, ids, "\n\n".join(warnings),
                [written] + details)
