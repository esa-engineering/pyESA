# -*- coding: utf-8 -*-
"""Openings: comando Place (fori nuovi sugli attraversamenti scelti; gli elementi gia' associati vanno al check)."""
from openings_report import *  # noqa: F401,F403


def missing_families_exit(missing_families):
    forms.alert("Load the opening families first (or choose others in Settings):\n{}".format(
        "\n".join(missing_families)), title=ctx.title, exitscript=True)


def place_main():
    import OpeningsPlace_ui  # finestre di Place, in openings_modules
    output = ctx.output
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
                output.print_md("# Openings Place")
                output.print_md("Nothing to do: the **{}** selected elements are not pipes, ducts, cable trays, "
                                "conduits, openings, walls, floors or structural elements.".format(ignored))
            return
        if missing_families:
            missing_families_exit(missing_families)
    else:
        if missing_families:
            missing_families_exit(missing_families)
        pick = OpeningsPlace_ui.PickWindow(bundle_path(PICK_XAML), last_pick(),
                                           "{}\n{}".format(settings_line(), family_line())).show()
        if pick is None:
            script.exit()
        remember_pick(pick)
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
                items.extend(find_crossings(area_hosts, stats, only=set(id_value(m.Id) for m in area_meps),
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
        items.extend(find_crossings(found, stats, only=set(id_value(m.Id) for m in meps)))
        modes.append("from {} MEP elements: walls, floors and structural elements found {}".format(
            len(meps), len(found)))
    mode_text = "; ".join(modes)
    items = merge_overlapping(items)
    groups = clusters(items)
    output.print_md("# Openings Place")
    output.print_md("Settings: {}".format(settings_line()))
    output.print_md("Mode: **{}** | crossings: **{}** | openings to create: **{}** | already with an opening: **{}** | "
                    "MEP running inside (no opening): **{}** | extended through adjacent layers: **{}**".format(
                        mode_text, len(items), len(groups), stats["existing"], stats["inside"], stats["extended"]))
    if stats["outside_area"]:
        output.print_md("Crossings of these MEP elements outside the area (skipped): **{}**".format(
            stats["outside_area"]))
    if ignored:
        output.print_md("Selected elements not used (not MEP, opening, wall, floor or structure, or not straight): "
                        "**{}**".format(ignored))
    for problem in stats["problems"]:
        output.print_md("- {}".format(problem))
    if not groups:
        output.print_md("Nothing to create.")
        return
    preview = {"mode": mode_text, "crossings": len(items), "openings": len(groups), "existing": stats["existing"],
               "inside": stats["inside"], "extended": stats["extended"]}
    if not OpeningsPlace_ui.ConfirmWindow(bundle_path(CONFIRM_XAML), preview, stats["problems"]).show():
        script.exit()

    missing = set()
    rows = []
    mep_codes = {}  # id elemento MEP -> (elemento, codici dei fori nuovi)
    managed = set(code_of(o) for o in existing_openings())  # sui MEP si tengono solo i codici di fori esistenti
    transaction = DB.Transaction(ctx.doc, ctx.title)
    transaction.Start()
    try:
        for group in groups:
            frame, kind = group[0]["frame"], group[0]["kind"]
            data = sizes(group, frame)
            meps = []
            for item in group:
                if all(m.Id != item["mep"].Id for m in meps):
                    meps.append(item["mep"])
            instance = create_opening(data, frame, kind, symbols, meps, NEW, missing)
            code = code_of(instance)
            for mep in meps:
                mep_codes.setdefault(id_value(mep.Id), (mep, []))[1].append(code)
            key, host, model = group[0]["host"]
            rows.append([output.linkify(instance.Id), disciplines_text(meps), categories_text(meps),
                         "{} {} ({})".format(host.Category.Name, id_value(host.Id), model), data["shape"],
                         size_text(data), ", ".join(str(id_value(m.Id)) for m in meps)])
        for mep, new_codes in mep_codes.values():
            old = [c for c in split_list(get_text(mep, ctx.id_param)) if c in managed]
            # con altri fori gia' presenti resta il loro stato, se peggiore
            status = worst([VALID] + ([get_text(mep, ctx.status_param) or VALID] if old else []))
            write_mep(mep, old + [c for c in new_codes if c not in old], status, missing)
        transaction.Commit()
    except Exception:
        if transaction.GetStatus() == DB.TransactionStatus.Started:
            transaction.RollBack()
        raise
    output.print_table(table_data=rows, columns=["Opening", "Discipline", "Categories", "Crossed element", "Shape",
                                                 "Size [mm]", "MEP elements"])
    output.print_md("New openings and their MEP elements: {} = **{}**.".format(ctx.status_param, VALID))
    if missing:
        output.print_md("**Parameters not found (not written):** {}".format("; ".join(sorted(missing))))
