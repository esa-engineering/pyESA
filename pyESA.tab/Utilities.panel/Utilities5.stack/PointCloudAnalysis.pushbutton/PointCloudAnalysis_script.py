# -*- coding: utf-8 -*-
__title__ = "PointCloud\nAnalysis"
__doc__ = """Compare a linked point cloud with selected model faces.

For each face a sampling grid is built and the average signed
distance between the cloud points and the surface is computed,
then displayed as a colored Analysis Visualization map on the
active view.

Positive values = cloud in front of the surface,
negative values = cloud behind the surface.

Note: the sampling grid is uniform on planar faces only."""
__author__ = "Antonio Miano"

import os
import math
import codecs
import datetime
import clr

clr.AddReference('RevitAPI')
clr.AddReference('RevitAPIUI')
clr.AddReference('PresentationFramework')
clr.AddReference('PresentationCore')
clr.AddReference('WindowsBase')

from Autodesk.Revit.DB import (
    FilteredElementCollector,
    Transaction,
    Element,
    Face,
    Plane,
    UV,
    XYZ,
    Color,
    UnitUtils,
    LabelUtils,
    PointCloudInstance,
)
from Autodesk.Revit.DB.PointClouds import PointCloudFilterFactory
from Autodesk.Revit.DB.Analysis import (
    SpatialFieldManager,
    AnalysisResultSchema,
    FieldDomainPointsByUV,
    FieldValues,
    ValueAtPoint,
    AnalysisDisplayStyle,
    AnalysisDisplayColoredSurfaceSettings,
    AnalysisDisplayColorSettings,
    AnalysisDisplayColorEntry,
    AnalysisDisplayStyleColorSettingsType,
    AnalysisDisplayLegendSettings,
)
from Autodesk.Revit.UI.Selection import ObjectType
from Autodesk.Revit.Exceptions import OperationCanceledException
from System.Collections.Generic import List

from System.Windows import Window
from System.Windows.Markup import XamlReader
from System.IO import StreamReader

from pyrevit import forms, script

# Revit >= 2021 uses ForgeTypeId units, older versions DisplayUnitType
try:
    from Autodesk.Revit.DB import SpecTypeId, UnitTypeId
    USE_FORGE_UNITS = True
except ImportError:
    from Autodesk.Revit.DB import UnitType, DisplayUnitType
    USE_FORGE_UNITS = False

uidoc = __revit__.ActiveUIDocument
doc = uidoc.Document

SCHEMA_NAME = "PointCloud Deviation"
STYLE_NAME = "ESA PointCloud Deviation"

QUALITY_ORDER = ['fast', 'medium', 'high', 'all']
QUALITY_POINTS = {'fast': 10, 'medium': 200, 'high': 1000}
# 'all': no practical cap and a tiny average spacing, so GetPoints
# effectively returns every cloud point inside the cell box
ALL_POINTS_CAP = 1000000
ALL_AVG_DIST_M = 0.0001

# Defaults in meters, converted to document units for the form
DEFAULT_GRID_M = 0.3
DEFAULT_OFFSET_M = 0.1
DEFAULT_QUALITY = 'high'

# Color gradients: name -> RGB stops from scale min to scale max
GRADIENTS = [
    ('Blue > White > Red', [(0, 62, 255), (255, 255, 255), (255, 32, 0)]),
    ('Blue > Red', [(0, 62, 255), (255, 32, 0)]),
    ('Green > Yellow > Red', [(0, 160, 60), (255, 220, 0), (255, 32, 0)]),
    ('Red > Yellow > Green', [(255, 32, 0), (255, 220, 0), (0, 160, 60)]),
    ('Grayscale', [(40, 40, 40), (230, 230, 230)]),
]
GRADIENT_MAP = dict(GRADIENTS)
DEFAULT_GRADIENT = GRADIENTS[0][0]
HISTOGRAM_MIN_BINS = 5
HISTOGRAM_MAX_BINS = 30


# ---------------------------------------------------------------- units

def get_length_unit():
    if USE_FORGE_UNITS:
        return doc.GetUnits().GetFormatOptions(SpecTypeId.Length).GetUnitTypeId()
    return doc.GetUnits().GetFormatOptions(UnitType.UT_Length).DisplayUnits


LENGTH_UNIT = get_length_unit()


def to_internal(value):
    return UnitUtils.ConvertToInternalUnits(value, LENGTH_UNIT)


def from_internal(value):
    return UnitUtils.ConvertFromInternalUnits(value, LENGTH_UNIT)


def meters_to_internal(value):
    if USE_FORGE_UNITS:
        return UnitUtils.ConvertToInternalUnits(value, UnitTypeId.Meters)
    return UnitUtils.ConvertToInternalUnits(value, DisplayUnitType.DUT_METERS)


def length_unit_label():
    try:
        if USE_FORGE_UNITS:
            return LabelUtils.GetLabelForUnit(LENGTH_UNIT)
        return LabelUtils.GetLabelFor(LENGTH_UNIT)
    except Exception:
        return ""


# ---------------------------------------------------------------- form

class AnalysisParamsForm(Window):
    """Dialog for grid spacing, offset, quality and color gradient."""

    def __init__(self):
        Window.__init__(self)
        self.result = None
        self._load_xaml()

    def _load_xaml(self):
        script_dir = os.path.dirname(__file__)
        xaml_path = os.path.join(script_dir, 'PointCloudAnalysisForm.xaml')

        reader = StreamReader(xaml_path)
        try:
            root = XamlReader.Load(reader.BaseStream)
        finally:
            reader.Close()

        self.Content = root.Content
        self.Title = root.Title
        self.Height = root.Height
        self.Width = root.Width
        self.MinHeight = root.MinHeight
        self.MinWidth = root.MinWidth
        self.WindowStartupLocation = root.WindowStartupLocation
        self.ResizeMode = root.ResizeMode
        self.ShowInTaskbar = root.ShowInTaskbar

        self.lbl_units = root.FindName('lbl_units')
        self.txt_grid = root.FindName('txt_grid')
        self.txt_offset = root.FindName('txt_offset')
        self.cmb_quality = root.FindName('cmb_quality')
        self.cmb_gradient = root.FindName('cmb_gradient')
        self.btn_ok = root.FindName('btn_ok')
        self.btn_cancel = root.FindName('btn_cancel')

        unit = length_unit_label()
        if unit:
            self.lbl_units.Text = "Length values in document units ({0})".format(unit)

        self.txt_grid.Text = "{0:g}".format(from_internal(meters_to_internal(DEFAULT_GRID_M)))
        self.txt_offset.Text = "{0:g}".format(from_internal(meters_to_internal(DEFAULT_OFFSET_M)))

        for quality in QUALITY_ORDER:
            self.cmb_quality.Items.Add(quality)
        self.cmb_quality.SelectedItem = DEFAULT_QUALITY

        for gradient_name, _ in GRADIENTS:
            self.cmb_gradient.Items.Add(gradient_name)
        self.cmb_gradient.SelectedItem = DEFAULT_GRADIENT

        self.btn_ok.Click += self.OnOk
        self.btn_cancel.Click += self.OnCancel

    @staticmethod
    def _parse_positive(text):
        try:
            value = float(text.strip().replace(',', '.'))
        except (ValueError, AttributeError):
            return None
        if value <= 0:
            return None
        return value

    def OnOk(self, sender, args):
        grid = self._parse_positive(self.txt_grid.Text)
        offset = self._parse_positive(self.txt_offset.Text)
        if grid is None or offset is None:
            forms.alert("Grid spacing and offset must be positive numbers.",
                        title="PointCloud Analysis")
            return
        self.result = {
            'grid': to_internal(grid),
            'offset': to_internal(offset),
            'quality': self.cmb_quality.SelectedItem,
            'gradient': self.cmb_gradient.SelectedItem,
        }
        self.Close()

    def OnCancel(self, sender, args):
        self.Close()


# ---------------------------------------------------------------- geometry helpers

def frange_centers(start, end, step):
    """Yield cell-center values from start to end with the given step."""
    value = start + step * 0.5
    while value < end:
        yield value
        value += step


def transform_plane(xform, plane):
    """Return the plane mapped through the given transform."""
    normal = xform.OfVector(plane.Normal).Normalize()
    origin = xform.OfPoint(plane.Origin)
    return Plane.CreateByNormalAndOrigin(normal, origin)


def box_planes(bmin, bmax):
    """Six inward-facing planes of a world-aligned box."""
    return [
        Plane.CreateByNormalAndOrigin(XYZ.BasisX, bmin),
        Plane.CreateByNormalAndOrigin(XYZ.BasisX.Negate(), bmax),
        Plane.CreateByNormalAndOrigin(XYZ.BasisY, bmin),
        Plane.CreateByNormalAndOrigin(XYZ.BasisY.Negate(), bmax),
        Plane.CreateByNormalAndOrigin(XYZ.BasisZ, bmin),
        Plane.CreateByNormalAndOrigin(XYZ.BasisZ.Negate(), bmax),
    ]


def cell_planes(face, uvp, grid, offset):
    """Six inward-facing planes of a box centered on the face cell,
    oriented along the face normal (works on any face orientation)."""
    deriv = face.ComputeDerivatives(uvp)
    normal = face.ComputeNormal(uvp)
    tang_u = deriv.BasisX.Normalize()
    tang_v = normal.CrossProduct(tang_u).Normalize()
    center = deriv.Origin
    half = grid * 0.5
    return [
        Plane.CreateByNormalAndOrigin(tang_u, center.Subtract(tang_u.Multiply(half))),
        Plane.CreateByNormalAndOrigin(tang_u.Negate(), center.Add(tang_u.Multiply(half))),
        Plane.CreateByNormalAndOrigin(tang_v, center.Subtract(tang_v.Multiply(half))),
        Plane.CreateByNormalAndOrigin(tang_v.Negate(), center.Add(tang_v.Multiply(half))),
        Plane.CreateByNormalAndOrigin(normal, center.Subtract(normal.Multiply(offset))),
        Plane.CreateByNormalAndOrigin(normal.Negate(), center.Add(normal.Multiply(offset))),
    ]


def get_cloud_points(pcl, planes, avg_dist, max_points):
    plane_list = List[Plane](planes)
    pc_filter = PointCloudFilterFactory.CreateMultiPlaneFilter(plane_list)
    return pcl.GetPoints(pc_filter, avg_dist, max_points)


def count_cloud_points(pcl, planes, avg_dist, max_points):
    count = 0
    for _ in get_cloud_points(pcl, planes, avg_dist, max_points):
        count += 1
    return count


def face_model_bbox(face, margin):
    """Approximate world bounding box of a face, expanded by margin."""
    bb = face.GetBoundingBox()
    samples = []
    for u_par in (bb.Min.U, (bb.Min.U + bb.Max.U) * 0.5, bb.Max.U):
        for v_par in (bb.Min.V, (bb.Min.V + bb.Max.V) * 0.5, bb.Max.V):
            try:
                samples.append(face.Evaluate(UV(u_par, v_par)))
            except Exception:
                continue
    xs = [p.X for p in samples]
    ys = [p.Y for p in samples]
    zs = [p.Z for p in samples]
    bmin = XYZ(min(xs) - margin, min(ys) - margin, min(zs) - margin)
    bmax = XYZ(max(xs) + margin, max(ys) + margin, max(zs) + margin)
    return bmin, bmax


def detect_filter_space(pcl, total_transform, face, margin):
    """Return the transform to apply to filter planes (or None).

    The Revit API is ambiguous on whether PointCloudInstance.GetPoints
    applies the filter in model or in cloud coordinates: the same box is
    tested in both spaces around the first face, and the space that
    returns more points wins. With an identity transform both are equal
    and None (model space) is returned.
    """
    if total_transform.IsIdentity:
        return None
    bmin, bmax = face_model_bbox(face, margin)
    planes_model = box_planes(bmin, bmax)
    inverse = total_transform.Inverse
    planes_cloud = [transform_plane(inverse, p) for p in planes_model]
    probe_dist = meters_to_internal(0.05)
    count_model = count_cloud_points(pcl, planes_model, probe_dist, 100)
    count_cloud = count_cloud_points(pcl, planes_cloud, probe_dist, 100)
    if count_cloud > count_model:
        return inverse
    return None


# ---------------------------------------------------------------- analysis

def quality_settings(quality, grid):
    """Return (max points per cell, average sampling distance)."""
    if quality in QUALITY_POINTS:
        max_points = QUALITY_POINTS[quality]
        return max_points, grid / math.sqrt(max_points)
    return ALL_POINTS_CAP, meters_to_internal(ALL_AVG_DIST_M)


def analyze_face(pcl, total_transform, to_cloud, face, grid, offset,
                 avg_dist, max_points):
    """Sample the face on a UV grid and return per-cell UVs, centers
    (model XYZ), average signed distances (internal units) and cloud
    point counts."""
    uvs = []
    centers = []
    values = []
    counts = []
    cells_total = 0
    cells_empty = 0

    bb = face.GetBoundingBox()
    for u_par in frange_centers(bb.Min.U, bb.Max.U, grid):
        for v_par in frange_centers(bb.Min.V, bb.Max.V, grid):
            uvp = UV(u_par, v_par)
            try:
                if not face.IsInside(uvp):
                    continue
                planes = cell_planes(face, uvp, grid, offset)
            except Exception:
                continue
            cells_total += 1

            if to_cloud is not None:
                planes = [transform_plane(to_cloud, p) for p in planes]

            distances = []
            for cloud_pt in get_cloud_points(pcl, planes, avg_dist, max_points):
                model_pt = total_transform.OfPoint(
                    XYZ(cloud_pt.X, cloud_pt.Y, cloud_pt.Z))
                projection = face.Project(model_pt)
                if projection is None:
                    continue
                normal = face.ComputeNormal(projection.UVPoint)
                distances.append(
                    normal.DotProduct(model_pt.Subtract(projection.XYZPoint)))

            if distances:
                uvs.append(uvp)
                centers.append(face.Evaluate(uvp))
                values.append(sum(distances) / len(distances))
                counts.append(len(distances))
            else:
                cells_empty += 1

    return uvs, centers, values, counts, cells_total, cells_empty


# ---------------------------------------------------------------- AVF display

def get_result_schema_index(sfm):
    for idx in sfm.GetRegisteredResults():
        if sfm.GetResultSchema(idx).Name == SCHEMA_NAME:
            return idx
    schema = AnalysisResultSchema(
        SCHEMA_NAME,
        "Average signed distance between point cloud and surface")
    label = length_unit_label()
    if label:
        try:
            schema.SetUnits(List[str]([label]), List[float]([1.0]))
            schema.CurrentUnits = 0
        except Exception:
            pass
    return sfm.RegisterResult(schema)


def build_style_settings(stops):
    surface_settings = AnalysisDisplayColoredSurfaceSettings()
    surface_settings.ShowGridLines = False

    color_settings = AnalysisDisplayColorSettings()
    color_settings.ColorSettingsType = \
        AnalysisDisplayStyleColorSettingsType.GradientColor
    color_settings.MinColor = Color(stops[0][0], stops[0][1], stops[0][2])
    color_settings.MaxColor = Color(stops[-1][0], stops[-1][1], stops[-1][2])
    if len(stops) > 2:
        entries = List[AnalysisDisplayColorEntry](
            [AnalysisDisplayColorEntry(Color(s[0], s[1], s[2]))
             for s in stops[1:-1]])
        color_settings.SetIntermediateColors(entries)

    legend_settings = AnalysisDisplayLegendSettings()
    legend_settings.ShowLegend = True
    legend_settings.ShowDataDescription = False
    legend_settings.NumberOfSteps = 10
    return surface_settings, color_settings, legend_settings


def update_display_style(stops):
    """Get or create the display style, always re-applying the settings
    so gradient changes and legend visibility take effect on every run."""
    surface_settings, color_settings, legend_settings = \
        build_style_settings(stops)
    for style in FilteredElementCollector(doc).OfClass(AnalysisDisplayStyle):
        if style.Name == STYLE_NAME:
            style.SetColoredSurfaceSettings(surface_settings)
            style.SetColorSettings(color_settings)
            style.SetLegendSettings(legend_settings)
            return style
    return AnalysisDisplayStyle.CreateAnalysisDisplayStyle(
        doc, STYLE_NAME, surface_settings, color_settings, legend_settings)


def show_results(view, face_results, stops):
    """Draw the AVF colored map on the view. Values in document units."""
    t = Transaction(doc, "PointCloud Analysis")
    t.Start()
    try:
        sfm = SpatialFieldManager.GetSpatialFieldManager(view)
        if sfm is None:
            sfm = SpatialFieldManager.CreateSpatialFieldManager(view, 1)

        # Remove primitives left over from previous runs
        sfm.Clear()

        schema_idx = get_result_schema_index(sfm)
        view.AnalysisDisplayStyleId = update_display_style(stops).Id

        for reference, uvs, values in face_results:
            prim_idx = sfm.AddSpatialFieldPrimitive(reference)
            domain_pts = FieldDomainPointsByUV(List[UV](uvs))
            value_list = List[ValueAtPoint](
                [ValueAtPoint(List[float]([v])) for v in values])
            sfm.UpdateSpatialFieldPrimitive(
                prim_idx, domain_pts, FieldValues(value_list), schema_idx)
        t.Commit()
    except Exception:
        if t.HasStarted():
            t.RollBack()
        raise


# ---------------------------------------------------------------- histogram

def gradient_hex(stops, t):
    """Interpolate the gradient stops at t in [0, 1], return a hex color."""
    if t <= 0:
        rgb = stops[0]
    elif t >= 1:
        rgb = stops[-1]
    else:
        scaled = t * (len(stops) - 1)
        i = int(scaled)
        frac = scaled - i
        a, b = stops[i], stops[i + 1]
        rgb = [int(round(a[k] + (b[k] - a[k]) * frac)) for k in range(3)]
    return '#{0:02x}{1:02x}{2:02x}'.format(rgb[0], rgb[1], rgb[2])


def histogram_bins(values, lo, hi):
    """Number of histogram bins via the Freedman-Diaconis rule
    (Sturges as fallback when the IQR is degenerate), clamped so the
    bar labels stay readable."""
    n = len(values)
    if n < 2:
        return HISTOGRAM_MIN_BINS
    ordered = sorted(values)
    q1 = ordered[int(round(0.25 * (n - 1)))]
    q3 = ordered[int(round(0.75 * (n - 1)))]
    iqr = q3 - q1
    if iqr > 1e-12:
        bin_width = 2.0 * iqr / (n ** (1.0 / 3.0))
        bins = int(math.ceil((hi - lo) / bin_width))
    else:
        bins = int(math.ceil(math.log(n, 2))) + 1
    return max(HISTOGRAM_MIN_BINS, min(HISTOGRAM_MAX_BINS, bins))


def print_histogram(output, values, counts, stops, unit):
    """Bar chart in the pyRevit output window: deviation bins on the
    x axis, number of cloud points per bin on the y axis (each cell
    weighted by the cloud points it contains), point count above bars.

    Rendered as plain HTML/CSS (no Chart.js: the pyRevit chart API
    loads it from a CDN, which fails offline or behind a proxy)."""
    lo = min(values)
    hi = max(values)
    if hi - lo < 1e-9:
        hi = lo + 1e-9
    n_bins = histogram_bins(values, lo, hi)
    width = (hi - lo) / n_bins

    point_bins = [0] * n_bins
    cell_bins = [0] * n_bins
    for value, n_points in zip(values, counts):
        idx = int((value - lo) / width)
        if idx < 0:
            idx = 0
        elif idx >= n_bins:
            idx = n_bins - 1
        point_bins[idx] += n_points
        cell_bins[idx] += 1

    peak = max(point_bins)
    if peak == 0:
        return
    plot_height = 420
    label_height = 22
    max_bar_height = plot_height - label_height
    bar_width = 100.0 / n_bins

    bars = []
    for i, n_points in enumerate(point_bins):
        bin_lo = lo + width * i
        bin_hi = bin_lo + width
        center = bin_lo + width * 0.5
        height = int(round(max_bar_height * float(n_points) / peak))
        if n_points > 0 and height < 3:
            height = 3
        color = gradient_hex(stops, (center - lo) / (hi - lo))
        label = '{0}'.format(n_points) if n_points > 0 else '&nbsp;'
        tooltip = "{0:.3f} .. {1:.3f} {2}: {3} point(s) in {4} cell(s)".format(
            bin_lo, bin_hi, unit, n_points, cell_bins[i])
        bars.append(
            '<div title="{0}" style="display:inline-block;'
            'vertical-align:bottom;width:{1:.2f}%;margin:0;padding:0;">'
            '<div style="font-size:15px;line-height:{2}px;'
            'text-align:center;color:#444;white-space:nowrap;'
            'overflow:visible;">{3}</div>'
            '<div style="height:{4}px;background:{5};'
            'margin:0 1px;"></div>'
            '</div>'.format(
                tooltip, bar_width, label_height, label, height, color))

    html = (
        '<div style="margin-top:15px;">'
        '<p style="font-weight:bold;font-size:21px;margin-bottom:12px;">'
        'Deviation distribution [{unit}] - {total_pts} cloud points '
        'in {total_cells} cells</p>'
        '<div style="height:{plot_h}px;font-size:0;line-height:0;'
        'border-bottom:2px solid #999;border-left:2px solid #999;">'
        '{bars}</div>'
        '<div style="font-size:19px;color:#666;overflow:hidden;">'
        '<span style="float:left;">{lo:.3f}</span>'
        '<span style="float:right;">{hi:.3f}</span>'
        '<span style="display:block;text-align:center;">{mid:.3f}</span>'
        '</div>'
        '<p style="font-size:17px;color:#666;margin-top:8px;">'
        'Average deviation [{unit}] on the x axis, number of cloud '
        'points on the y axis. Hover the bars for bin details.</p>'
        '</div>').format(
            unit=unit,
            total_pts=sum(point_bins),
            total_cells=len(values),
            plot_h=plot_height,
            bars=''.join(bars),
            lo=lo,
            hi=hi,
            mid=(lo + hi) * 0.5)
    output.print_html(html)


# ---------------------------------------------------------------- CSV export

def get_id_value(eid):
    """Return the integer value of an ElementId (works in Revit 2022-2026+)."""
    try:
        return eid.Value
    except AttributeError:
        return eid.IntegerValue


def export_csv(rows, unit):
    """Write the analyzed cells to a CSV file in the temp folder and
    return its path. Semicolon separator and decimal comma (Italian
    Excel conventions)."""
    stamp = datetime.datetime.now().strftime('%Y%m%d_%H%M%S')
    folder = os.environ.get('TEMP') or os.path.expanduser('~')
    path = os.path.join(folder, 'PointCloudAnalysis_{0}.csv'.format(stamp))

    def num(value):
        return '{0:.4f}'.format(value).replace('.', ',')

    lines = ['ElementId;Face;Cell X [{0}];Cell Y [{0}];Cell Z [{0}];'
             'Avg deviation [{0}];Cloud points'.format(unit)]
    for row in rows:
        lines.append('{0};{1};{2};{3};{4};{5};{6}'.format(
            row[0], row[1], num(row[2]), num(row[3]), num(row[4]),
            num(row[5]), row[6]))

    stream = codecs.open(path, 'w', 'utf-8-sig')
    try:
        stream.write('\r\n'.join(lines))
    finally:
        stream.close()
    return path


def print_csv_button(output, path):
    html = (
        '<div style="margin-top:15px;">'
        '<a href="file:///{href}" style="display:inline-block;'
        'background:#2D5A8A;color:#ffffff;font-weight:bold;'
        'font-size:17px;padding:10px 20px;border-radius:4px;'
        'text-decoration:none;">Open CSV (analyzed cells)</a>'
        '<p style="font-size:14px;color:#666;margin-top:6px;">'
        'File saved to: {path}</p>'
        '</div>').format(href=path.replace('\\', '/'), path=path)
    output.print_html(html)


# ---------------------------------------------------------------- main

def pick_point_cloud():
    clouds = list(FilteredElementCollector(doc)
                  .OfClass(PointCloudInstance)
                  .WhereElementIsNotElementType())
    if not clouds:
        forms.alert("No point cloud found in the document.",
                    title="PointCloud Analysis")
        script.exit()
    if len(clouds) == 1:
        return clouds[0]
    labels = {}
    for cloud in clouds:
        name = Element.Name.GetValue(cloud)
        labels["{0}  (id {1})".format(name, cloud.Id)] = cloud
    choice = forms.SelectFromList.show(
        sorted(labels.keys()),
        title="PointCloud Analysis",
        button_name="Use this point cloud",
        multiselect=False)
    if not choice:
        script.exit()
    return labels[choice]


def pick_faces():
    try:
        references = uidoc.Selection.PickObjects(
            ObjectType.Face, "Select the faces to analyze, then press Finish")
    except OperationCanceledException:
        script.exit()
    faces = []
    skipped = 0
    for reference in references:
        element = doc.GetElement(reference)
        geo_obj = None
        try:
            geo_obj = element.GetGeometryObjectFromReference(reference)
        except Exception:
            pass
        if isinstance(geo_obj, Face):
            faces.append((reference, geo_obj))
        else:
            skipped += 1
    if skipped:
        forms.alert(
            "{0} selected face(s) skipped (faces of linked models are "
            "not supported).".format(skipped),
            title="PointCloud Analysis")
    if not faces:
        script.exit()
    return faces


def main():
    view = doc.ActiveView
    output = script.get_output()
    unit = length_unit_label()

    pcl = pick_point_cloud()
    faces = pick_faces()

    form = AnalysisParamsForm()
    form.ShowDialog()
    if form.result is None:
        script.exit()
    params = form.result

    grid = params['grid']
    offset = params['offset']
    stops = GRADIENT_MAP[params['gradient']]
    max_points, avg_dist = quality_settings(params['quality'], grid)

    total_transform = pcl.GetTotalTransform()
    to_cloud = detect_filter_space(
        pcl, total_transform, faces[0][1], offset + grid)

    print("PointCloud Analysis")
    print("Point cloud: {0}".format(Element.Name.GetValue(pcl)))
    if params['quality'] == 'all':
        print("Quality: all (every cloud point in each cell)")
    else:
        print("Quality: {0} ({1} points per cell max)".format(
            params['quality'], max_points))
    if to_cloud is not None:
        print("Cloud filter applied in point cloud coordinates "
              "(non-identity cloud transform detected).")
    print("")

    face_results = []
    all_values = []
    all_counts = []
    csv_rows = []
    total_cells = 0
    empty_cells = 0
    with forms.ProgressBar(title="Analyzing face {value} of {max_value}",
                           cancellable=True) as progress:
        for i, (reference, face) in enumerate(faces):
            if progress.cancelled:
                script.exit()
            progress.update_progress(i + 1, len(faces))
            uvs, centers, values, counts, n_cells, n_empty = analyze_face(
                pcl, total_transform, to_cloud, face,
                grid, offset, avg_dist, max_points)
            total_cells += n_cells
            empty_cells += n_empty
            if uvs:
                display_values = [from_internal(v) for v in values]
                all_values.extend(display_values)
                all_counts.extend(counts)
                face_results.append((reference, uvs, display_values))
                elem_id = get_id_value(reference.ElementId)
                for center, value, count in zip(
                        centers, display_values, counts):
                    csv_rows.append((
                        elem_id, i + 1,
                        from_internal(center.X),
                        from_internal(center.Y),
                        from_internal(center.Z),
                        value, count))

    if not face_results:
        forms.alert(
            "No cloud points found near the selected faces.\n"
            "Try increasing the offset or check the point cloud position.",
            title="PointCloud Analysis")
        script.exit()

    try:
        show_results(view, face_results, stops)
    except Exception as ex:
        forms.alert(
            "Could not display the analysis on the active view:\n{0}\n\n"
            "Try again from a 3D view.".format(ex),
            title="PointCloud Analysis")
        script.exit()

    print("Faces analyzed: {0}".format(len(faces)))
    print("Cells evaluated: {0} (no cloud points in {1})".format(
        total_cells, empty_cells))
    print("Cloud points analyzed: {0}".format(sum(all_counts)))
    print("Deviation [{0}]  min: {1:.3f}   max: {2:.3f}   avg: {3:.3f}".format(
        unit, min(all_values), max(all_values),
        sum(all_values) / len(all_values)))

    print_histogram(output, all_values, all_counts, stops, unit)

    try:
        csv_path = export_csv(csv_rows, unit)
        print_csv_button(output, csv_path)
    except Exception as ex:
        print("CSV export failed: {0}".format(ex))


main()
