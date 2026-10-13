# -*- coding: utf-8 -*-
__title__ = "Families"
__doc__ = """Version = 1.0
Date    = 13.10.2026
________________________________________________________________
Description:
Creates the four opening families used by Place and Refresh, saves them
as .rfa in a folder you choose and loads them into the active model
(already loaded: overwritten, the openings placed keep their sizes).

- e_GM.PH.PFV_Wall_Rect   box: X Width, Y Depth, Z Height
- e_GM.PH.PFV_Wall_Round  cylinder along Y: Diameter, Depth
- e_GM.PH.PFV_Floor_Rect  box: X Width, Y Height, Z Depth
- e_GM.PH.PFV_Floor_Round cylinder along Z: Diameter, Depth

Generic Model, not hosted, origin at the centre of the opening, size
parameters by instance, symmetric about the origin (EQ). Each family is
flexed after creation: the report says if a parameter does not drive the
geometry.

The template is the Generic Model template of the Revit library; if it
is not found you are asked to pick it.
________________________________________________________________
Author(s): Matteo Maiorano
"""
__author__ = "Matteo Maiorano"

import math
import os

import clr
clr.AddReference("RevitAPI")
from pyrevit import revit, script, DB, forms

# stessi nomi predefiniti di Place, Refresh e Settings: tenerli allineati
FAMILY_WALL_RECT = "e_GM.PH.PFV_Wall_Rect"
FAMILY_WALL_ROUND = "e_GM.PH.PFV_Wall_Round"
FAMILY_FLOOR_RECT = "e_GM.PH.PFV_Floor_Rect"
FAMILY_FLOOR_ROUND = "e_GM.PH.PFV_Floor_Round"
PARAM_WIDTH = "e_DIM_Width_1"
PARAM_HEIGHT = "e_DIM_Height_1"
PARAM_DEPTH = "e_DIM_Depth_1"
PARAM_DIAMETER = "e_DIM_Diameter_1"

MM_PER_FT = 304.8
DEFAULTS_MM = {PARAM_WIDTH: 300.0, PARAM_HEIGHT: 200.0, PARAM_DEPTH: 250.0, PARAM_DIAMETER: 150.0}
DIM_GAP = 300.0 / MM_PER_FT    # distanza delle quote dalla geometria
SPECS = ((FAMILY_WALL_RECT, "wall", "rect"), (FAMILY_WALL_ROUND, "wall", "round"),
         (FAMILY_FLOOR_RECT, "floor", "rect"), (FAMILY_FLOOR_ROUND, "floor", "round"))
TEMPLATE_WORDS = ("generic model", "modello generico")
TEMPLATE_EXCLUDE = ("face", "wall", "ceiling", "floor", "line", "adaptive", "pattern", "level", "roof", "based",
                    "basato", "faccia", "muro", "soffitto", "pavimento", "linea", "adattivo", "modello di")
X, Y, Z = DB.XYZ.BasisX, DB.XYZ.BasisY, DB.XYZ.BasisZ
ORIGIN = DB.XYZ.Zero

doc = revit.doc
app = __revit__.Application  # noqa: F821
output = script.get_output()


class FamilyError(Exception):
    """Errore atteso nella costruzione di una famiglia: il messaggio va nel report."""


def ft(value_mm):
    return value_mm / MM_PER_FT


# ----------------------------------------------------------------------------------------------------------------
# Template, viste e piani del documento famiglia
# ----------------------------------------------------------------------------------------------------------------
def find_template():
    """Template Generic Model non ospitato della libreria di Revit (metrico se c'e'), altrimenti scelto a mano."""
    root = app.FamilyTemplatePath
    found = []
    if root and os.path.isdir(root):
        for folder, _, files in os.walk(root):
            for name in files:
                low = name.lower()
                if low.endswith(".rft") and any(w in low for w in TEMPLATE_WORDS) and \
                        not any(w in low for w in TEMPLATE_EXCLUDE):
                    found.append(os.path.join(folder, name))
    if found:
        found.sort(key=lambda p: (0 if "metric" in p.lower() or "metrico" in p.lower() else 1, len(p)))
        return found[0]
    return forms.pick_file(file_ext="rft", init_dir=root or None,
                           title="Pick the Generic Model family template (not hosted)")


def plan_view(fdoc):
    for view in DB.FilteredElementCollector(fdoc).OfClass(DB.ViewPlan):
        if not view.IsTemplate:
            return view
    raise FamilyError("the template has no plan view")


def front_view(fdoc):
    """Prospetto che guarda lungo Y (Front o Back): vi si vedono di taglio le facce perpendicolari a X e a Z."""
    for view in DB.FilteredElementCollector(fdoc).OfClass(DB.View):
        if not view.IsTemplate and view.ViewType == DB.ViewType.Elevation and abs(view.ViewDirection.Y) > 0.99:
            return view
    raise FamilyError("the template has no front elevation")


def center_plane(fdoc, normal):
    """Piano di riferimento del template per l'origine con la normale data (Center Left/Right, Front/Back)."""
    for plane in DB.FilteredElementCollector(fdoc).OfClass(DB.ReferencePlane):
        if abs(plane.Normal.DotProduct(normal)) > 0.999 and abs(plane.BubbleEnd.DotProduct(normal)) < 1e-6:
            return plane
    return None


def make_center_plane(fdoc, normal, view):
    """Piano per l'origine, se il template non lo ha (per Z: orizzontale, nel prospetto). Bloccato."""
    existing = center_plane(fdoc, normal)
    if existing is not None:
        return existing
    if normal.IsAlmostEqualTo(Z):
        plane = fdoc.FamilyCreate.NewReferencePlane(DB.XYZ(-1, 0, 0), DB.XYZ(1, 0, 0), Y, view)
    elif normal.IsAlmostEqualTo(X):
        plane = fdoc.FamilyCreate.NewReferencePlane(DB.XYZ(0, -1, 0), DB.XYZ(0, 1, 0), Z, view)
    else:
        plane = fdoc.FamilyCreate.NewReferencePlane(DB.XYZ(-1, 0, 0), DB.XYZ(1, 0, 0), Z, view)
    plane.Name = "PFV Center {}".format("Z" if normal.IsAlmostEqualTo(Z) else ("X" if normal.IsAlmostEqualTo(X)
                                                                                else "Y"))
    plane.Pinned = True
    return plane


# ----------------------------------------------------------------------------------------------------------------
# Parametri
# ----------------------------------------------------------------------------------------------------------------
NOT_SHARED = set()  # parametri non trovati nel file dei parametri condivisi: creati come parametri di famiglia


def shared_definitions():
    """{nome: ExternalDefinition} dal file dei parametri condivisi impostato in Revit (vuoto se non c'e')."""
    found = {}
    try:
        shared_file = app.OpenSharedParameterFile()
    except Exception:
        shared_file = None
    if shared_file is None:
        return found
    for group in shared_file.Groups:
        for definition in group.Definitions:
            found.setdefault(definition.Name, definition)
    return found


SHARED = shared_definitions()


def add_length_parameter(fdoc, name):
    """Parametro di istanza di lunghezza: condiviso se il file dei parametri condivisi lo contiene (serve per abachi
    ed etichette), altrimenti parametro di famiglia (annotato nel report)."""
    manager = fdoc.FamilyManager
    existing = manager.get_Parameter(name)
    if existing is not None:
        return existing
    try:
        group = DB.GroupTypeId.Geometry  # Revit 2022+
    except AttributeError:
        group = DB.BuiltInParameterGroup.PG_GEOMETRY
    if name in SHARED:
        return manager.AddParameter(SHARED[name], group, True)
    NOT_SHARED.add(name)
    try:
        return manager.AddParameter(name, group, DB.SpecTypeId.Length, True)
    except AttributeError:
        return manager.AddParameter(name, group, DB.ParameterType.Length, True)


def ensure_type(fdoc):
    manager = fdoc.FamilyManager
    if manager.CurrentType is None:
        manager.NewType("Standard")


# ----------------------------------------------------------------------------------------------------------------
# Geometria e quote
# ----------------------------------------------------------------------------------------------------------------
def rectangle(a, b):
    """Rettangolo 2a x 2b nel piano XY, centrato sull'origine."""
    corners = [DB.XYZ(-a, -b, 0), DB.XYZ(a, -b, 0), DB.XYZ(a, b, 0), DB.XYZ(-a, b, 0)]
    loop = DB.CurveArray()
    for i in range(4):
        loop.Append(DB.Line.CreateBound(corners[i], corners[(i + 1) % 4]))
    profile = DB.CurveArrArray()
    profile.Append(loop)
    return profile


def circle(radius, normal):
    """Cerchio (due archi) centrato sull'origine, nel piano perpendicolare a normal."""
    plane = DB.Plane.CreateByNormalAndOrigin(normal, ORIGIN)
    loop = DB.CurveArray()
    loop.Append(DB.Arc.Create(plane, radius, 0.0, math.pi))
    loop.Append(DB.Arc.Create(plane, radius, math.pi, 2.0 * math.pi))
    profile = DB.CurveArrArray()
    profile.Append(loop)
    return profile


def extrude(fdoc, profile, normal, length):
    """Estrusione dal profilo lungo normal, simmetrica rispetto all'origine."""
    sketch = DB.SketchPlane.Create(fdoc, DB.Plane.CreateByNormalAndOrigin(normal, ORIGIN))
    extrusion = fdoc.FamilyCreate.NewExtrusion(True, profile, sketch, length)
    extrusion.StartOffset = -length / 2.0
    extrusion.EndOffset = length / 2.0
    fdoc.Regenerate()
    return extrusion


def face_refs(extrusion):
    """{(asse, segno): riferimento} delle facce piane dell'estrusione perpendicolari agli assi."""
    options = DB.Options()
    options.ComputeReferences = True
    refs = {}
    for item in extrusion.get_Geometry(options):
        if not isinstance(item, DB.Solid):
            continue
        for face in item.Faces:
            if not isinstance(face, DB.PlanarFace):
                continue
            n = face.FaceNormal
            for axis_name, axis in (("X", X), ("Y", Y), ("Z", Z)):
                d = n.DotProduct(axis)
                if abs(d) > 0.999:
                    refs[(axis_name, 1 if d > 0 else -1)] = face.Reference
    return refs


def ref_array(refs):
    array = DB.ReferenceArray()
    for ref in refs:
        array.Append(ref)
    return array


def dimension_axis(fdoc, view, refs, axis_name, center, param, half_other):
    """Quota complessiva tra le due facce lungo l'asse, con etichetta param, e quota EQ faccia-centro-faccia."""
    low, high = refs.get((axis_name, -1)), refs.get((axis_name, 1))
    if low is None or high is None:
        raise FamilyError("faces along {} not found".format(axis_name))
    offset = half_other + DIM_GAP
    if axis_name == "X":
        line_a = DB.Line.CreateBound(DB.XYZ(-1, -offset, 0), DB.XYZ(1, -offset, 0))
        line_b = DB.Line.CreateBound(DB.XYZ(-1, -offset - DIM_GAP, 0), DB.XYZ(1, -offset - DIM_GAP, 0))
        if view.ViewType == DB.ViewType.Elevation:
            line_a = DB.Line.CreateBound(DB.XYZ(-1, 0, -offset), DB.XYZ(1, 0, -offset))
            line_b = DB.Line.CreateBound(DB.XYZ(-1, 0, -offset - DIM_GAP), DB.XYZ(1, 0, -offset - DIM_GAP))
    elif axis_name == "Y":
        line_a = DB.Line.CreateBound(DB.XYZ(-offset, -1, 0), DB.XYZ(-offset, 1, 0))
        line_b = DB.Line.CreateBound(DB.XYZ(-offset - DIM_GAP, -1, 0), DB.XYZ(-offset - DIM_GAP, 1, 0))
    else:
        line_a = DB.Line.CreateBound(DB.XYZ(-offset, 0, -1), DB.XYZ(-offset, 0, 1))
        line_b = DB.Line.CreateBound(DB.XYZ(-offset - DIM_GAP, 0, -1), DB.XYZ(-offset - DIM_GAP, 0, 1))
    eq = fdoc.FamilyCreate.NewDimension(view, line_a, ref_array([low, center.GetReference(), high]))
    eq.AreSegmentsEqual = True
    overall = fdoc.FamilyCreate.NewDimension(view, line_b, ref_array([low, high]))
    overall.FamilyLabel = param


def circle_reference(extrusion):
    """Riferimento di un arco del cerchio: dallo sketch, altrimenti da uno spigolo circolare del solido."""
    try:
        for curve in extrusion.Sketch.Profile.get_Item(0):
            if isinstance(curve, DB.Arc) and curve.Reference is not None:
                return curve.Reference, curve
    except Exception:
        pass
    options = DB.Options()
    options.ComputeReferences = True
    for item in extrusion.get_Geometry(options):
        if isinstance(item, DB.Solid):
            for edge in item.Edges:
                curve = edge.AsCurve()
                if isinstance(curve, DB.Arc) and edge.Reference is not None:
                    return edge.Reference, curve
    raise FamilyError("the circle has no reference for the diameter dimension")


def build(fdoc, kind, shape):
    """Geometria, parametri e quote di una famiglia. Ritorna {nome parametro: asse misurato nella prova}."""
    ensure_type(fdoc)
    plan, front = plan_view(fdoc), front_view(fdoc)
    params = {}
    for name in ((PARAM_DIAMETER, PARAM_DEPTH) if shape == "round" else (PARAM_WIDTH, PARAM_HEIGHT, PARAM_DEPTH)):
        params[name] = add_length_parameter(fdoc, name)
        fdoc.FamilyManager.Set(params[name], ft(DEFAULTS_MM[name]))
    cx = make_center_plane(fdoc, X, plan)
    cy = make_center_plane(fdoc, Y, plan)
    cz = make_center_plane(fdoc, Z, front)
    fdoc.Regenerate()
    width, height = ft(DEFAULTS_MM[PARAM_WIDTH]), ft(DEFAULTS_MM[PARAM_HEIGHT])
    depth, diameter = ft(DEFAULTS_MM[PARAM_DEPTH]), ft(DEFAULTS_MM[PARAM_DIAMETER])

    if shape == "rect":
        # muro: X larghezza, Y profondita', Z altezza; solaio: X larghezza, Y altezza, Z profondita'
        size_y, size_z = (depth, height) if kind == "wall" else (height, depth)
        extrusion = extrude(fdoc, rectangle(width / 2.0, size_y / 2.0), Z, size_z)
        refs = face_refs(extrusion)
        name_y, name_z = (PARAM_DEPTH, PARAM_HEIGHT) if kind == "wall" else (PARAM_HEIGHT, PARAM_DEPTH)
        dimension_axis(fdoc, plan, refs, "X", cx, params[PARAM_WIDTH], size_y / 2.0)
        dimension_axis(fdoc, plan, refs, "Y", cy, params[name_y], width / 2.0)
        dimension_axis(fdoc, front, refs, "Z", cz, params[name_z], width / 2.0)
        return extrusion, {PARAM_WIDTH: "X", name_y: "Y", name_z: "Z"}

    # tondo: muro, cilindro lungo Y (cerchio nel prospetto); solaio, lungo Z (cerchio in pianta)
    axis, circle_view, depth_view, depth_axis, depth_center = (Y, front, plan, "Y", cy) if kind == "wall" else \
        (Z, plan, front, "Z", cz)
    extrusion = extrude(fdoc, circle(diameter / 2.0, axis), axis, depth)
    refs = face_refs(extrusion)
    reference, arc = circle_reference(extrusion)
    point = arc.Evaluate(0.25, True)
    dim = fdoc.FamilyCreate.NewDiameterDimension(circle_view, reference, point)
    dim.FamilyLabel = params[PARAM_DIAMETER]
    dimension_axis(fdoc, depth_view, refs, depth_axis, depth_center, params[PARAM_DEPTH], diameter / 2.0)
    return extrusion, {PARAM_DIAMETER: "X", PARAM_DEPTH: depth_axis}


def flex(fdoc, extrusion, axes):
    """Prova: ogni parametro portato a 1,5 volte deve cambiare l'ingombro lungo il suo asse, con il centro fermo.
    Ritorna l'elenco dei problemi (vuoto = tutto a posto). I valori tornano ai predefiniti."""
    manager = fdoc.FamilyManager
    problems = []
    for name, axis in axes.items():
        param = manager.get_Parameter(name)
        test = ft(DEFAULTS_MM[name] * 1.5)
        manager.Set(param, test)
        fdoc.Regenerate()
        box = extrusion.get_BoundingBox(None)
        index = "XYZ".index(axis)
        size = [box.Max.X - box.Min.X, box.Max.Y - box.Min.Y, box.Max.Z - box.Min.Z][index]
        middle = [(box.Max.X + box.Min.X) / 2.0, (box.Max.Y + box.Min.Y) / 2.0, (box.Max.Z + box.Min.Z) / 2.0][index]
        if abs(size - test) > ft(0.5):
            problems.append("{} does not drive the size along {} ({:.0f} instead of {:.0f} mm)".format(
                name, axis, size * MM_PER_FT, test * MM_PER_FT))
        if abs(middle) > ft(0.5):
            problems.append("{} moves the centre along {} by {:.0f} mm".format(name, axis, middle * MM_PER_FT))
        manager.Set(param, ft(DEFAULTS_MM[name]))
        fdoc.Regenerate()
    return problems


def create_family(template, folder, name, kind, shape):
    """Crea, prova e salva una famiglia. Ritorna (percorso, problemi)."""
    fdoc = app.NewFamilyDocument(template)
    if fdoc is None:
        raise FamilyError("the template could not be opened")
    try:
        transaction = DB.Transaction(fdoc, "PFV family")
        transaction.Start()
        try:
            extrusion, axes = build(fdoc, kind, shape)
            problems = flex(fdoc, extrusion, axes)
            transaction.Commit()
        except Exception:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            raise
        path = os.path.join(folder, name + ".rfa")
        options = DB.SaveAsOptions()
        options.OverwriteExistingFile = True
        fdoc.SaveAs(path, options)
        return path, problems
    finally:
        fdoc.Close(False)


class LoadOptions(DB.IFamilyLoadOptions):
    """Famiglia gia' nel modello: si sovrascrive (le istanze tengono i loro valori)."""

    def OnFamilyFound(self, familyInUse, overwriteParameterValues):
        overwriteParameterValues.Value = True
        return True

    def OnSharedFamilyFound(self, sharedFamily, familyInUse, source, overwriteParameterValues):
        source.Value = DB.FamilySource.Family
        overwriteParameterValues.Value = True
        return True


def main():
    template = find_template()
    if not template:
        script.exit()
    folder = forms.pick_folder(title="Folder where the opening families (.rfa) are saved")
    if not folder:
        script.exit()

    rows, paths = [], []
    for name, kind, shape in SPECS:
        try:
            path, problems = create_family(template, folder, name, kind, shape)
            paths.append((name, path))
            rows.append([name, "Created" if not problems else "Created, check it", "; ".join(problems) or
                         "flexed: sizes follow the parameters, centre fixed"])
        except Exception as ex:
            rows.append([name, "Not created", str(ex)])

    loaded = []
    if paths and doc is not None and not doc.IsFamilyDocument:
        transaction = DB.Transaction(doc, "Load PFV families")
        transaction.Start()
        try:
            for name, path in paths:
                result = doc.LoadFamily(path, LoadOptions())
                ok = result[0] if isinstance(result, tuple) else result
                loaded.append("{}: {}".format(name, "loaded" if ok else "not loaded (already up to date?)"))
            transaction.Commit()
        except Exception as ex:
            if transaction.GetStatus() == DB.TransactionStatus.Started:
                transaction.RollBack()
            loaded.append("loading failed: {}".format(ex))

    output.print_md("# Opening families")
    output.print_md("Template: `{}`  \nFolder: `{}`".format(template, folder))
    output.print_table(table_data=rows, columns=["Family", "Result", "Check"])
    if loaded:
        output.print_md("**Model:** {}".format(" | ".join(loaded)))
    if NOT_SHARED:
        output.print_md("**Not in the shared parameter file, created as family parameters** (they work for Place "
                        "and Refresh, but cannot be scheduled or tagged): {}. Set the shared parameter file in "
                        "Revit (Manage > Shared Parameters) and run Families again.".format(
                            ", ".join(sorted(NOT_SHARED))))


main()
