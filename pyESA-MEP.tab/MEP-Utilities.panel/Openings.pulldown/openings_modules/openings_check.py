# -*- coding: utf-8 -*-
"""Openings: controllo dei fori contro i loro elementi MEP, ambito del controllo, divisioni e unioni, ricreazione,
scrittura degli stati e degli Id, scansione degli attraversamenti senza foro (Missing openings)."""
from openings_pick import *  # noqa: F401,F403


def check_opening(instance, extra=(), scanned=None):
    """Confronto di un foro con i suoi elementi MEP. Ritorna il dict del risultato; "elements" ha una riga per ogni
    Id elencato nel foro, con il suo stato (ELEMENT_*). extra: elementi MEP che lo attraversano senza esservi
    associati (scansione del check): se lo attraversano davvero entrano nel foro ("added"). scanned: {id elemento:
    attraversamenti della scansione}, per ricalcolare la profondita' (scanned_plane)."""
    listed = split_list(get_text(instance, ctx.id_param))
    found = []
    result = {"instance": instance, "code": code_of(instance), "keep": [], "notes": [], "state": OK, "new": None,
              "elements": [], "stale": []}

    def element_row(text, mep, state):
        result["elements"].append({"id": text, "element": mep, "name": mep_name(mep) if mep is not None else
                                   "Element", "state": state})

    for text in listed:
        element = element_by_id(text)
        if element is None or not isinstance(element, MEP_CLASSES):
            result["notes"].append("{} no longer exists".format(text))
            result["stale"].append(text)  # Id vecchio: tolto dal foro alla scrittura
            element_row(text, None, ELEMENT_DELETED)
        else:
            found.append(element)
    kind, shape = opening_kind(instance, found)
    frame = opening_frame(instance, kind, shape)
    old = current_data(instance, kind, shape)
    result.update({"kind": kind, "frame": frame, "old": old, "plane": (old["center"], old["depth"])})
    items, all_inside = [], True
    result["misplaced"] = []  # (elemento, attraversamento): fuori dal foro, il foro si puo' spostare su di lui
    listed_set = set(listed)
    # elementi che attraversano il foro senza esservi associati (trovati dalla scansione del check): si associano
    extra = [m for m in extra if str(id_value(m.Id)) not in listed_set]
    result["added"] = []
    for mep in found + extra:
        text = str(id_value(mep.Id))
        if text not in listed_set:
            # non era associato: entra solo se attraversa davvero il foro
            try:
                section = mep_section(mep)
            except Exception:
                continue
            crossing = recross(mep, frame, old)
            if crossing is None:
                continue
            item = crossing_item(mep, section, crossing[0], crossing[1], frame, kind)
            inside = inside_opening(item, frame, old)
            if not inside:
                all_inside = False
                result["misplaced"].append((mep, item))
            result["notes"].append("{} crosses the opening{} but was not associated: added".format(
                text, "" if inside else ", partly outside,"))
            element_row(text, mep, ELEMENT_ADDED if inside else ELEMENT_ADDED_OUT)
            result["added"].append(mep)
            items.append(item)
            result["keep"].append(mep)
            continue
        try:
            section = mep_section(mep)
        except Exception:
            result["notes"].append("{}: section not readable".format(text))
            element_row(text, mep, ELEMENT_UNREADABLE)
            continue
        crossing = recross(mep, frame, old)
        if crossing is None:
            far = recross(mep, frame, old, limit_reach=False)
            if far is None:
                result["notes"].append("{} no longer crosses the opening".format(text))
                element_row(text, mep, ELEMENT_NOT_CROSSING)
            else:
                # ancora parallelo all'attraversamento ma spostato lontano: il foro si puo' spostare su di lui
                result["notes"].append("{} moved out of the opening".format(text))
                result["misplaced"].append((mep, crossing_item(mep, section, far[0], far[1], frame, kind)))
                element_row(text, mep, ELEMENT_MOVED_OUT)
            continue
        item = crossing_item(mep, section, crossing[0], crossing[1], frame, kind)
        if not inside_opening(item, frame, old):
            all_inside = False
            # entro la distanza di unione ma fuori del tutto: uscito, non "in parte fuori"
            gone = not overlaps_opening(item, frame, old)
            result["notes"].append("{} {}".format(text, "moved out of the opening" if gone else
                                                  "partly outside the opening"))
            result["misplaced"].append((mep, item))
            element_row(text, mep, ELEMENT_MOVED_OUT if gone else ELEMENT_PARTLY_OUT)
        else:
            element_row(text, mep, ELEMENT_INSIDE)
        items.append(item)
        result["keep"].append(mep)
    result["found_meps"] = found + result["added"]  # restano associati al foro finche' non lo si ricrea
    result["found"] = len(result["found_meps"])  # per spostare il foro: foro di un solo elemento
    result["items"] = items
    result["groups"] = [items]  # fori che ne prendono il posto, se ricreato (due se MEC ed ELE vanno separati)
    if not items:
        result["state"] = OBSOLETE
        return result
    # piano e profondita': dalla scansione del check se c'e' (strati nuovi o tolti), altrimenti quelli attuali
    result["plane"] = scanned_plane(result["keep"], frame, old, scanned)
    new = rebuilt(items, frame, result["plane"][0], result["plane"][1])
    result["new"] = new
    # elementi elencati spariti, non piu' attraversanti o illeggibili (quelli aggiunti non contano)
    lost = len([m for m in result["keep"] if str(id_value(m.Id)) in listed_set]) < len(listed)
    diffs = differences(old, new)
    result["notes"].extend(diffs)
    if lost or not all_inside:
        # elementi ancora dentro o in parte fuori: To Be Refreshed (se attivo in Settings); Not Associated solo se
        # non ne resta nessuno
        present = (ELEMENT_INSIDE, ELEMENT_PARTLY_OUT, ELEMENT_ADDED, ELEMENT_ADDED_OUT)
        kept = [e for e in result["elements"] if e["state"] in present]
        result["state"] = PARTIAL if kept and ctx.settings["misaligned"] else CHANGED
    elif diffs:
        result["state"] = MISALIGNED if ctx.settings["misaligned"] else CHANGED
    return result


def linked_openings(openings, sel_meps, sel_openings):
    """Fori legati agli elementi scelti: i fori stessi, quelli elencati sui MEP e quelli che elencano i MEP."""
    mep_ids = set(str(id_value(m.Id)) for m in sel_meps)
    codes = set(c for m in sel_meps for c in split_list(get_text(m, ctx.id_param)))
    wanted = set(id_value(o.Id) for o in sel_openings)
    return [o for o in openings if id_value(o.Id) in wanted or code_of(o) in codes or
            mep_ids & set(split_list(get_text(o, ctx.id_param)))]


class CheckScope(object):
    """Fori del check: tutti (ids None) o un insieme di Id di fori, aggiornato dopo ricreazioni. meps: elementi MEP
    scelti, le cui liste si riscrivono anche se non hanno piu' fori. removed: fori cancellati dalle ricreazioni: i MEP
    che li citano ancora si riscrivono anche se il foro non e' piu' nell'ambito. area (make_area): ambito Area, a
    ogni controllo entrano anche i fori e gli elementi MEP comparsi nell'area nel frattempo."""

    def __init__(self, kind, description, ids=None, meps=None, released=None, area=None):
        self.kind, self.description, self.ids, self.meps = kind, description, ids, list(meps or [])
        self.area = area
        self.removed = set()
        # {Id elemento (testo): [codici dei fori cancellati che lo elencavano]}: elementi rimasti senza foro dopo un
        # Delete, tenuti nell'ambito e nel report finche' un foro nuovo non li circonda
        self.released = dict(released or {})

    def release(self, element, codes):
        text = str(id_value(element.Id))
        self.released[text] = sorted(set(self.released.get(text, [])) | set(codes), key=int)
        if all(m.Id != element.Id for m in self.meps):
            self.meps.append(element)

    def add(self, codes):
        """Fori entrati nell'ambito (attraversati da elementi dell'ambito non associati)."""
        if self.ids is not None:
            self.ids |= set(codes)

    def openings(self):
        found = existing_openings()
        if self.ids is None:
            return found
        return [o for o in found if code_of(o) in self.ids or (self.area is not None and opening_in_area(self.area, o))]

    def replace(self, removed, created):
        self.removed |= set(removed)
        if self.ids is not None:
            self.ids = (self.ids - set(removed)) | set(created)


def discipline_checks(results):
    """Divisioni e unioni dei fori secondo le impostazioni attuali di Settings. Gli elementi sono ancora dentro: lo
    stato e' almeno REGROUP (To Be Refreshed), resta Changed (Not Associated) o Partial se lo era gia'.
    - MEC ed ELE separati: un foro con entrambe le discipline, ricreato, diventa due fori.
    - Unione: due o piu' fori sullo stesso piano i cui fori ricalcolati (franco compreso) distano al massimo la
      distanza di unione, ricreati, diventano un foro solo. Vale per elementi avvicinati e per la spunta MEC/ELE
      tolta; con MEC ed ELE separati si uniscono solo fori della stessa disciplina.
    I fori Changed e Partial (elementi tolti o usciti) non entrano nelle unioni: altrimenti un elemento spostato vicino a un
    foro in ordine lo renderebbe Not Associated. Ricreato il foro Changed, il check dopo propone l'unione. Cosi' i
    gruppi hanno solo fori To Be Refreshed e si ricreano tutti insieme o per niente."""
    live = [r for r in results if r["state"] != OBSOLETE and r["items"]]
    if ctx.settings["separate"]:
        for r in live:
            split = [[i for i in r["items"] if i["discipline"] == d] for d in ("MEC", "ELE")]
            split = [g for g in split if g]
            if len(split) > 1:
                r["groups"] = split
                r["state"] = max(r["state"], REGROUP, key=STATE_RANK.get)
                r["notes"].append("MEC and ELE must be separate (Settings): it becomes two openings")
    merge = ctx.settings["merge_mm"] / MM_PER_FT
    candidates = [r for r in live if len(r["groups"]) == 1 and r["state"] not in (CHANGED, PARTIAL)]
    parent = dict((id(r), id(r)) for r in candidates)

    def root(key):
        while parent[key] != key:
            key = parent[key]
        return key

    def disciplines(r):
        return set(i["discipline"] for i in r["items"])

    for i in range(len(candidates)):
        for j in range(i + 1, len(candidates)):
            a, b = candidates[i], candidates[j]
            if ctx.settings["separate"] and (len(disciplines(a)) > 1 or disciplines(a) != disciplines(b)):
                continue
            u, v, w = a["frame"]
            if abs(w.DotProduct(b["frame"][2])) < 0.999:
                continue
            if mm(abs(b["old"]["center"].Subtract(a["old"]["center"]).DotProduct(w))) > MOVE_TOL_MM:
                continue  # non sullo stesso piano: muri o solai diversi
            # fori come andrebbero ricalcolati (misure con il franco): si uniscono se distano <= unione
            d = b["new"]["center"].Subtract(a["new"]["center"])
            ha, hb = [[h / MM_PER_FT for h in half_sizes(r["new"])] for r in (a, b)]
            gap = max(abs(d.DotProduct(u)) - ha[0] - hb[0], abs(d.DotProduct(v)) - ha[1] - hb[1])
            if gap <= merge:
                parent[root(id(b))] = root(id(a))
    groups = {}
    for r in candidates:
        groups.setdefault(root(id(r)), []).append(r)
    for members in groups.values():
        if len(members) < 2:
            continue
        codes = ", ".join(m["code"] for m in members)
        state = max([REGROUP] + [m["state"] for m in members], key=STATE_RANK.get)
        for m in members:
            m["merge_group"] = members
            m["state"] = state
            m["notes"].append("close enough to be one opening: openings {} become one".format(codes))


def alignment_vector(r):
    """Spostamento (nel piano del foro, misure invariate) che riporta il foro sui suoi elementi usciti in parte o del
    tutto; None se nessun elemento e' fuori o lo spostamento e' sotto 1 mm."""
    if not r.get("misplaced"):
        return None
    items = list(r["items"]) + [item for _, item in r["misplaced"] if item not in r["items"]]
    vector = opening_shift(items, r["frame"], r["old"], r["found"] <= 1)
    return vector if vector.GetLength() * MM_PER_FT >= 1.0 else None


def outside_meps(r):
    """Elementi del foro che esistono ancora ma non attraversano piu' il foro (nemmeno entro la portata): Recreate non
    li include nel foro nuovo e perdono l'associazione, ma restano nell'ambito (CheckScope.release): la scansione li
    mostra in Missing openings finche' un foro nuovo non li circonda."""
    return [m for m in r["found_meps"] if all(m.Id != k.Id for k in r["keep"])]


def recreate_plan(r):
    """Cosa fa Recreate su questo foro, in una riga (dettaglio del report e conferma)."""
    if not r["items"]:
        return "Recreate: not possible, no element crosses it any more."
    lost = len(outside_meps(r))
    lost_text = "; {} element(s) outside it lose the association and stay listed in Missing openings where they " \
                "cross a wall or floor".format(lost) if lost else ""
    if r.get("merge_group"):
        others = [m["code"] for m in r["merge_group"] if m is not r]
        return "Recreate: joined with opening(s) {} into one opening{}.".format(", ".join(others), lost_text)
    if len(r["groups"]) > 1:
        return "Recreate: split into {} openings ({}){}.".format(
            len(r["groups"]), " and ".join(group_discipline(g) for g in r["groups"]), lost_text)
    return "Recreate: one new opening on {} element(s){}.".format(len(unique_meps(r["items"])), lost_text)


def recreate_openings(results, keys, symbols, missing):
    """Ricrea i fori con Id in keys sugli elementi che li attraversano ancora (un gruppo da unire si ricrea tutto
    insieme, un foro misto da dividere diventa due fori). Chiamare dentro una transazione. Ritorna (codici creati,
    codici tolti, codici saltati perche' senza elementi, eventi in testo, {codice nuovo: origine})."""
    created, removed, skipped, events, origins, done = [], [], [], [], {}, set()
    for r in results:
        if r["code"] not in keys:
            continue
        if not r["items"]:
            skipped.append(r["code"])
            continue
        if r.get("merge_group"):
            key = id(r["merge_group"])
            if key in done:
                continue
            done.add(key)
            members = r["merge_group"]
            codes = [m["code"] for m in members]
            items = [i for m in members for i in m["items"]]
            data = rebuilt(items, r["frame"], r["plane"][0], max(m["plane"][1] for m in members))
            for m in members:
                removed.append(m["code"])
                ctx.doc.Delete(m["instance"].Id)
            new = code_of(create_opening(data, r["frame"], r["kind"], symbols, unique_meps(items), NEW, missing))
            created.append(new)
            origins[new] = "joined from {}".format(" + ".join(codes))
            ticked = [c for c in codes if c in keys]
            events.append("Openings {} joined into {}{}".format(" + ".join(codes), new, "" if len(ticked) == len(codes)
                                                                else " (the whole group, also the ones not ticked)"))
            continue
        removed.append(r["code"])
        ctx.doc.Delete(r["instance"].Id)
        news = []
        for group in r["groups"]:
            data = rebuilt(group, r["frame"], r["plane"][0], r["plane"][1])
            new = code_of(create_opening(data, r["frame"], r["kind"], symbols, unique_meps(group), NEW, missing))
            created.append(new)
            news.append(new)
            origins[new] = "split from {} ({})".format(r["code"], group_discipline(group)) if len(r["groups"]) > 1 \
                else "recreated from {}".format(r["code"])
        if len(news) > 1:
            events.append("Opening {} split into {}".format(r["code"], " and ".join(
                "{} ({})".format(n, group_discipline(g)) for n, g in zip(news, r["groups"]))))
        else:
            events.append("Opening {} recreated as {}".format(r["code"], news[0]))
    return created, removed, skipped, events, origins


def mep_candidates(reverse):
    """Elementi MEP da tenere allineati: con stato o Id compilati, o citati da un foro."""
    categories = DB.ElementMulticategoryFilter(List[DB.BuiltInCategory](MEP_CATEGORIES))
    found = {}
    for element in DB.FilteredElementCollector(ctx.doc).WherePasses(categories).WhereElementIsNotElementType():
        if not isinstance(element, MEP_CLASSES):
            continue
        text = str(id_value(element.Id))
        if text in reverse or get_text(element, ctx.status_param) or get_text(element, ctx.id_param):
            found[text] = element
    return found


def plan_writes(results, scope):
    """Cosa scrivere dopo il check, e gli Id vecchi trovati.
    - Fori controllati: stato, Id dei loro elementi ancora esistenti (anche se usciti dal foro: restano associati
      finche' il foro non si ricrea), categorie. Gli Id di elementi cancellati (o non piu' MEP) si tolgono.
    - Elementi MEP: Id dei fori che li citano (la lista e' lo specchio dei fori), categoria, stato peggiore tra
      quei fori; nessun foro: Not Associated. Fuori dall'ambito solo i MEP legati ai fori controllati, a quelli
      cancellati dalle ricreazioni, o scelti."""
    checked = dict((r["code"], r) for r in results)
    lists, statuses = {}, {}
    for opening in existing_openings():
        code = code_of(opening)
        if code in checked:
            r = checked[code]
            lists[code] = [str(id_value(m.Id)) for m in r["found_meps"]]
            statuses[code] = WRITTEN[r["state"]]
        else:
            lists[code] = split_list(get_text(opening, ctx.id_param))
            statuses[code] = get_text(opening, ctx.status_param) or VALID
    reverse = {}
    for code, ids in lists.items():
        for text in ids:
            reverse.setdefault(text, []).append(code)
    chosen = set(str(id_value(m.Id)) for m in scope.meps)
    relevant = set(checked) | scope.removed
    meps, stale_meps, added, updated, orphaned = [], [], 0, 0, 0
    for text, mep in mep_candidates(reverse).items():
        current = split_list(get_text(mep, ctx.id_param))
        codes = reverse.get(text, [])
        if scope.ids is not None and text not in chosen and not (set(codes) | set(current)) & relevant:
            continue
        status = worst([statuses[c] for c in codes]) if codes else NOT_ASSOCIATED
        removed = [c for c in current if c not in codes]
        added += len([c for c in codes if c not in current])
        if removed:
            stale_meps.append((mep, removed))
        if sorted(codes) != sorted(current):
            updated += 1
            if current and not codes:
                orphaned += 1
        meps.append((mep, codes, status))
    return {"meps": meps, "stale_meps": stale_meps, "added": added, "updated": updated, "orphaned": orphaned,
            "stale_openings": sum(len(r["stale"]) for r in results)}


# ----------------------------------------------------------------------------------------------------------------
# Scansione degli attraversamenti dell'ambito (New Report): elementi non associati e fori mancanti
# ----------------------------------------------------------------------------------------------------------------
def scan_meps(scope):
    """Elementi MEP del report: quelli elencati nei fori dell'ambito, piu' quelli scelti (selezione MEP) o nell'area
    (ambito Area, riletti a ogni controllo)."""
    meps = {}
    for mep in list(scope.meps) + (meps_in_area(scope.area) if scope.area is not None else []):
        if not mep.IsValidObject:
            continue
        if is_straight_mep(mep):
            meps[id_value(mep.Id)] = mep
    for opening in scope.openings():
        for text in split_list(get_text(opening, ctx.id_param)):
            element = element_by_id(text)
            if element is not None and is_straight_mep(element):
                meps[id_value(element.Id)] = element
    return list(meps.values())


def scan_crossings(scope):
    """Attraversamenti degli elementi MEP del report (scan_meps), come li trova Place: muri, solai e strutture del
    modello e dei link toccati da quegli elementi, con il foro che contiene ogni attraversamento (item["opening"]) o
    None. Solo questi elementi, qualunque sia l'ambito: non si cercano attraversamenti di altri elementi del modello.
    Ambito Area: muri, solai e strutture dell'area (non uno per elemento MEP, che sarebbero migliaia di ricerche) e
    solo gli attraversamenti dentro l'area. Ritorna (attraversamenti, statistiche)."""
    stats = {"inside": 0, "existing": 0, "extended": 0, "outside_area": 0, "problems": []}
    meps = scan_meps(scope)
    if not meps:
        return [], stats
    only = set(id_value(m.Id) for m in meps)
    hosts = hosts_in_area(scope.area) if scope.area is not None else hosts_from_meps(meps)
    if not hosts:
        return [], stats
    with forms.ProgressBar(title="Openings: looking for crossings ({value} of {max_value})") as bar:
        items = find_crossings(hosts, stats, only=only, keep_existing=True,
                               progress=lambda i, n: bar.update_progress(i + 1, n), area=scope.area)
    return items, stats


def unlisted_crossings(items):
    """{codice del foro: {id: elemento MEP}} degli elementi che attraversano un foro senza esservi associati."""
    found = {}
    for item in items:
        opening = item.get("opening")
        if opening is None:
            continue
        mep_id = id_value(item["mep"].Id)
        if str(mep_id) not in split_list(get_text(opening, ctx.id_param)):
            found.setdefault(code_of(opening), {})[mep_id] = item["mep"]
    return found


def handled_by_opening(item, results):
    """L'attraversamento e' gia' coperto dal check di un foro dell'elemento: l'elemento e' entro la portata del foro
    (anche se uscito: Recreate o Move lo sistemano) e l'attraversamento, parallelo al foro, ne tocca lo spessore (se
    e' piu' lungo, per uno strato nuovo, il foro si allunga: scanned_plane). Non e' un foro mancante, altrimenti si
    creerebbero due fori."""
    mep = item["mep"]
    tol = MOVE_TOL_MM / MM_PER_FT
    for r in results:
        if not any(m.Id == mep.Id for m in r["found_meps"]):
            continue
        if recross(mep, r["frame"], r["old"]) is None:
            continue
        w = r["frame"][2]
        if abs(item["frame"][2].DotProduct(w)) < 0.999:
            continue
        half = r["old"]["depth"] / MM_PER_FT / 2.0
        a, b = sorted((item["p0"].Subtract(r["old"]["center"]).DotProduct(w),
                       item["p1"].Subtract(r["old"]["center"]).DotProduct(w)))
        if b >= -half - tol and a <= half + tol:
            return True
    return False


def host_label(item):
    stack, element, model = item["host"]
    return "{} {} ({})".format(element.Category.Name, id_value(element.Id), model)


# campi di una riga Missing openings salvati per il Last Report (solo testo e Id)
MISSING_SAVED = ("key", "label", "host_text", "kind_text", "size", "crossings", "mep_ids")


def missing_openings(items, results):
    """Attraversamenti senza foro, raggruppati come li creerebbe Place. Ritorna le righe del gruppo Missing
    openings. La chiave e' stabile (elementi attraversati + elementi MEP): una spunta data nel Last Report resta
    valida dopo il ricontrollo che precede Create."""
    items = [i for i in items if i.get("opening") is None and not handled_by_opening(i, results)]
    rows = []
    for group in clusters(merge_overlapping(items)):
        frame, kind = group[0]["frame"], group[0]["kind"]
        data = sizes(group, frame)
        stack, host, _ = group[0]["host"]
        meps = unique_meps(group)
        mep_ids = sorted(id_value(m.Id) for m in meps)
        row = {"key": "new:{}:{}".format("+".join("{}_{}".format(a, b) for a, b in stack),
                                         "-".join(str(i) for i in mep_ids)),
               "group": group, "frame": frame, "kind": kind, "data": data, "meps": meps, "mep_ids": mep_ids,
               "host": host if host.Document.Equals(ctx.doc) else None, "host_text": host_label(group[0]),
               "kind_text": "{} {}".format(kind, data["shape"]), "size": size_text(data), "crossings": len(group)}
        row["label"] = "New opening | {} | {} | {}".format(row["kind_text"], row["size"], row["host_text"])
        rows.append(row)
    return rows
