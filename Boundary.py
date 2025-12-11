# -*- coding: utf-8 -*-
"""Pick a point, find floor below, then:
   - If there are NO non-tree planting instances in that floor footprint:
       Create Area Boundary lines offset 5 mm inside the floor perimeter.
   - If there ARE non-tree planting instances:
       Create an Area Boundary representing the floor area MINUS the
       union of plant spreads (inverse of plant area, scalloped where
       plants touch the edge), and DO NOT draw the simple floor-outline
       boundary.
   - In ALL cases:
       For any planting instance whose family/type/header contains 'tree'
       and has type parameter 'Trunk Diameter', draw a circle as an
       Area Boundary at the base with that diameter + 100 mm.
   - Also:
       Place an Area at the picked point in the Area Plan and, if a
       mix_name is provided, set the Area's Name parameter to mix_name.

   Update behaviour:
       - All created Area Boundary lines are given a dedicated line style
         named 'BM Mix Boundary'.
       - On each run, the script finds the current floor footprint,
         defines a 2D region around it (exact floor extents, no margin),
         and deletes any existing 'BM Mix Boundary' Area Boundary lines
         in that region before drawing the new ones. This allows easy
         updating when geometry/planting changes.
"""

import math
from System.Collections.Generic import List

from pyrevit import revit, DB, script
from Autodesk.Revit.UI import TaskDialog
from Autodesk.Revit.UI.Selection import ObjectSnapTypes

logger = script.get_logger()

# -------------------------- Tunables -----------------------------

SPREAD_PARAM_NAME = "Spread"         # Type param that stores plant diameter (ft)
TREE_KEYWORD = "tree"                # Look for this in name/family/header
TRUNK_PARAM_NAME = "Trunk Diameter"  # Tree type param for trunk diameter (ft)

SEARCH_MARGIN_M = 1.5                # Search a bit outside the floor boundary
PLANT_OFFSET_M = 0.20                # Extra outward offset on plant radius
FLOOR_OFFSET_MM = 5.0                # Floor boundary offset inward (mm) when no plants

TREE_EXTRA_DIAM_MM = 100.0           # Extra diameter added to Trunk Diameter circle

FT_PER_M = 1.0 / 0.3048

# Line style name used to mark boundaries created by this tool
MIX_BOUNDARY_STYLE_NAME = "BM Mix Boundary"


# ----------------- General helpers -------------------------------

def get_any_3d_view(doc):
    """Return any non-template 3D view, or None if not found."""
    views = DB.FilteredElementCollector(doc).OfClass(DB.View3D).ToElements()
    for v in views:
        if not v.IsTemplate:
            return v
    return None


def get_floor_below_point(doc, pickpoint):
    """Use a vertical raycast to find the nearest Floor below the picked point."""
    view3d = get_any_3d_view(doc)
    if not view3d:
        logger.error("No 3D view found for raycasting.")
        return None

    floor_filter = DB.ElementClassFilter(DB.Floor)
    ref_intersector = DB.ReferenceIntersector(
        floor_filter,
        DB.FindReferenceTarget.Element,
        view3d
    )

    origin = pickpoint + DB.XYZ(0, 0, 10.0)
    direction = DB.XYZ(0, 0, -1.0)

    result = ref_intersector.FindNearest(origin, direction)
    if result is None:
        logger.warning("No floor found below picked point.")
        return None

    ref = result.GetReference()
    floor = doc.GetElement(ref.ElementId)
    return floor


def get_top_planar_face(floor):
    """Return the upward-facing planar top face of the floor, or None."""
    opt = DB.Options()
    opt.ComputeReferences = True
    opt.DetailLevel = DB.ViewDetailLevel.Fine

    geom_elem = floor.get_Geometry(opt)
    if geom_elem is None:
        return None

    top_face = None
    top_z = None

    for gobj in geom_elem:
        solid = gobj if isinstance(gobj, DB.Solid) else None
        if not solid:
            continue
        if solid.Volume <= 0:
            continue

        for face in solid.Faces:
            planar = face if isinstance(face, DB.PlanarFace) else None
            if not planar:
                continue

            normal = planar.FaceNormal
            if normal.Z > 0.9:
                face_z = planar.Origin.Z
                if top_face is None or face_z > top_z:
                    top_face = planar
                    top_z = face_z

    return top_face


def flatten_curve_to_view_z(curve, target_z):
    """Translate the curve vertically so its first endpoint sits at target_z."""
    try:
        p0 = curve.GetEndPoint(0)
    except:
        return curve

    delta_z = target_z - p0.Z
    if abs(delta_z) < 1e-6:
        return curve

    transform = DB.Transform.CreateTranslation(DB.XYZ(0.0, 0.0, delta_z))
    return curve.CreateTransformed(transform)


# --------------- Offset floor loop (5 mm inwards) ----------------

def _sample_points_on_curveloop(curveloop, samples_per_curve):
    pts = []
    for crv in curveloop:
        for i in range(samples_per_curve):
            t = float(i) / float(samples_per_curve)
            try:
                pt = crv.Evaluate(t, True)
                pts.append(pt)
            except:
                pts.append(crv.GetEndPoint(0))
                pts.append(crv.GetEndPoint(1))
                break
    return pts


def _compute_centroid(points):
    if not points:
        return DB.XYZ(0, 0, 0)
    x = y = z = 0.0
    count = float(len(points))
    for p in points:
        x += p.X
        y += p.Y
        z += p.Z
    return DB.XYZ(x / count, y / count, z / count)


def _average_radius(curveloop, centroid, samples_per_curve):
    pts = _sample_points_on_curveloop(curveloop, samples_per_curve)
    if not pts:
        return 0.0
    total = 0.0
    for p in pts:
        v = p - centroid
        total += v.GetLength()
    return total / float(len(pts))


def offset_loop_inwards(curveloop, offset_dist_feet, normal, samples_per_curve=6):
    """Offset CurveLoop inward by approx offset_dist_feet (pick smaller-radius of +/- offset)."""
    orig_pts = _sample_points_on_curveloop(curveloop, samples_per_curve)
    centroid = _compute_centroid(orig_pts)

    candidates = []
    try:
        plus = DB.CurveLoop.CreateViaOffset(curveloop, offset_dist_feet, normal)
        candidates.append(plus)
    except:
        pass
    try:
        minus = DB.CurveLoop.CreateViaOffset(curveloop, -offset_dist_feet, normal)
        candidates.append(minus)
    except:
        pass

    if not candidates:
        return curveloop
    if len(candidates) == 1:
        return candidates[0]

    best_loop = None
    best_rad = None
    for cl in candidates:
        rad = _average_radius(cl, centroid, samples_per_curve)
        if best_loop is None or rad < best_rad:
            best_loop = cl
            best_rad = rad
    return best_loop


# -------------------- Line style / deletion ----------------------

def get_mix_boundary_linestyle(doc):
    """
    Get or create a dedicated GraphicsStyle (line style) for mix boundaries,
    under the Area Scheme Lines category.
    """
    cats = doc.Settings.Categories
    try:
        cat = cats.get_Item(DB.BuiltInCategory.OST_AreaSchemeLines)
    except:
        cat = None

    if cat is None:
        return None

    # Look for existing subcategory with desired name
    for sub in cat.SubCategories:
        if sub.Name == MIX_BOUNDARY_STYLE_NAME:
            try:
                return sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
            except:
                return None

    # Create new subcategory
    try:
        new_sub = cats.NewSubcategory(cat, MIX_BOUNDARY_STYLE_NAME)
        # Optional: tweak graphics (colour / weight) here if you want
        return new_sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
    except:
        return None


def delete_existing_boundaries_in_region(doc, view, minx, maxx, miny, maxy, boundary_style):
    """
    Delete Area Boundary curves in this view, within the XY region, that use
    the mix boundary line style. If boundary_style is None, delete all
    AreaScheme/AreaBoundary curves in that region (last-resort fallback).
    """
    ids_to_delete = List[DB.ElementId]()

    collector = DB.FilteredElementCollector(doc, view.Id).OfClass(DB.CurveElement)

    bic_scheme = int(DB.BuiltInCategory.OST_AreaSchemeLines)
    try:
        bic_boundary = int(DB.BuiltInCategory.OST_AreaBoundaryLines)
    except:
        bic_boundary = None

    for ce in collector:
        if ce is None:
            continue
        cat = ce.Category
        if not cat:
            continue
        cid = cat.Id.IntegerValue
        if cid != bic_scheme and (bic_boundary is None or cid != bic_boundary):
            continue

        # If we have a specific style, only delete that style
        if boundary_style is not None:
            try:
                ls = ce.LineStyle
            except:
                ls = None
            if ls is None or ls.Id != boundary_style.Id:
                continue

        bb = ce.get_BoundingBox(view)
        if bb is None:
            bb = ce.get_BoundingBox(None)
        if bb is None:
            continue

        cmin = bb.Min
        cmax = bb.Max

        if cmax.X < minx or cmin.X > maxx or cmax.Y < miny or cmin.Y > maxy:
            continue

        ids_to_delete.Add(ce.Id)

    if ids_to_delete.Count > 0:
        logger.info("Deleting %s existing mix boundary curves in region.", ids_to_delete.Count)
        doc.Delete(ids_to_delete)


# -------------------- Plant helpers ------------------------------

def is_tree_symbol(sym):
    """Return True if symbol looks like a tree (by name / Scheduling Header)."""
    try:
        tname = (sym.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString() or "").lower()
        fname = (sym.get_Parameter(DB.BuiltInParameter.SYMBOL_FAMILY_NAME_PARAM).AsString() or "").lower()
        hdr = ""
        p = sym.LookupParameter("Scheduling Header")
        if p and p.HasValue:
            hdr = (p.AsString() or "").lower()
        s = tname + " " + fname + " " + hdr
        return TREE_KEYWORD in s
    except:
        return False


def collect_plants(doc, poly_pts, search_margin_ft, floor_z):
    """
    Collect non-tree planting instances near polygon; return list of (center, radius_ft).
    Plants are treated as overlapping regardless of their actual Z – their
    centres are projected onto floor_z.
    """
    if not poly_pts:
        return []

    minx = min(p.X for p in poly_pts)
    maxx = max(p.X for p in poly_pts)
    miny = min(p.Y for p in poly_pts)
    maxy = max(p.Y for p in poly_pts)

    # Very tall vertical band so elevation doesn't exclude anything
    outline = DB.Outline(DB.XYZ(minx - search_margin_ft,
                                miny - search_margin_ft,
                                floor_z - 1000.0),
                         DB.XYZ(maxx + search_margin_ft,
                                maxy + search_margin_ft,
                                floor_z + 1000.0))
    bbfilter = DB.BoundingBoxIntersectsFilter(outline)

    discs = []
    col = DB.FilteredElementCollector(doc) \
             .OfCategory(DB.BuiltInCategory.OST_Planting) \
             .WhereElementIsNotElementType()

    for inst in col:
        try:
            try:
                if not bbfilter.PassesFilter(doc, inst.Id):
                    continue
            except:
                pass

            loc = inst.Location
            if not hasattr(loc, "Point") or loc.Point is None:
                continue
            raw_center = loc.Point

            sym = doc.GetElement(inst.GetTypeId())
            if is_tree_symbol(sym):
                # trees handled separately (trunk circles)
                continue

            p = sym.LookupParameter(SPREAD_PARAM_NAME)
            if not p or not p.HasValue:
                continue

            diam_ft = p.AsDouble()
            r = max(0.5 * diam_ft, 0.05)

            # Project plant centre to floor Z so discs always overlap the floor solid
            c_flat = DB.XYZ(raw_center.X, raw_center.Y, floor_z)
            discs.append((c_flat, r))
        except:
            continue

    return discs


def build_disc_solid(center, radius):
    """Create a thin extrusion solid of a full circle at 'center' with 'radius'."""
    loop = DB.CurveLoop()
    loop.Append(DB.Arc.Create(center, radius, 0.0, math.pi,
                              DB.XYZ.BasisX, DB.XYZ.BasisY))
    loop.Append(DB.Arc.Create(center, radius, math.pi, 2.0 * math.pi,
                              DB.XYZ.BasisX, DB.XYZ.BasisY))
    return DB.GeometryCreationUtilities.CreateExtrusionGeometry(
        [loop], DB.XYZ.BasisZ, 0.1
    )


# --------------- Plant area boundary creation --------------------

def create_plant_area_boundary(doc, view, outer_loop, target_z, sketch_plane, boundary_style):
    """
    Build Area Boundary lines for the floor area MINUS the union of plant discs.
    Returns True if any plant-based boundary was created, otherwise False.
    """
    # Poly points for plant search (use start points of each segment)
    poly_pts = [crv.GetEndPoint(0) for crv in outer_loop]
    if len(poly_pts) < 3:
        return False

    floor_z = poly_pts[0].Z

    discs = collect_plants(doc, poly_pts, SEARCH_MARGIN_M * FT_PER_M, floor_z)
    if not discs:
        return False

    # Union all disc solids (radius expanded by PLANT_OFFSET_M)
    union_solid = None
    off_ft = PLANT_OFFSET_M * FT_PER_M
    for (c, r) in discs:
        try:
            s = build_disc_solid(c, r + max(0.0, off_ft))
        except:
            continue
        if union_solid is None:
            union_solid = s
        else:
            try:
                union_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
                    union_solid, s, DB.BooleanOperationsType.Union)
            except:
                pass

    if union_solid is None:
        return False

    # Solid from floor loop (polygon)
    try:
        poly_solid = DB.GeometryCreationUtilities.CreateExtrusionGeometry(
            [outer_loop], DB.XYZ.BasisZ, 0.1)
    except:
        poly_solid = None

    if poly_solid is None:
        return False

    # FINAL TARGET REGION:
    #   floor area MINUS (union of plant discs)
    try:
        plantmix_solid = DB.BooleanOperationsUtils.ExecuteBooleanOperation(
            poly_solid, union_solid, DB.BooleanOperationsType.Difference)
    except:
        plantmix_solid = None

    if plantmix_solid is None:
        return False

    # Find top Z of plantmix_solid and project edges to target_z
    z_vals = []
    for e in plantmix_solid.Edges:
        cr = e.AsCurve()
        p0 = cr.GetEndPoint(0)
        p1 = cr.GetEndPoint(1)
        z_vals.append(p0.Z)
        z_vals.append(p1.Z)

    if not z_vals:
        return False

    max_z = max(z_vals)
    tol = 1e-6
    z_shift = target_z - max_z
    xform = DB.Transform.CreateTranslation(DB.XYZ(0, 0, z_shift))

    created = 0
    for e in plantmix_solid.Edges:
        cr = e.AsCurve()
        p0 = cr.GetEndPoint(0)
        p1 = cr.GetEndPoint(1)
        # Only edges lying on the top face of final region
        if abs(p0.Z - max_z) < tol and abs(p1.Z - max_z) < tol:
            try:
                proj = cr.CreateTransformed(xform)
                # Skip very short segments
                try:
                    length = proj.ApproximateLength
                except:
                    length = proj.Length
                if length < (1.0 / 48.0):  # ~6 mm
                    continue
                line = doc.Create.NewAreaBoundaryLine(sketch_plane, proj, view)
                if boundary_style is not None:
                    try:
                        line.LineStyle = boundary_style
                    except:
                        pass
                created += 1
            except:
                pass

    logger.info("Created %s plant-inverse area boundary segments.", created)
    return created > 0


# --------------- Tree trunk circles (area boundaries) ------------

def draw_tree_trunks(doc, view, outer_loop, target_z, sketch_plane, boundary_style):
    """
    For any planting instance whose symbol looks like a tree and whose
    type has parameter 'Trunk Diameter', draw a circle (two arcs) with
    that diameter + 100 mm as Area Boundary lines at the tree base.
    """
    poly_pts = [crv.GetEndPoint(0) for crv in outer_loop]
    if not poly_pts:
        return

    minx = min(p.X for p in poly_pts)
    maxx = max(p.X for p in poly_pts)
    miny = min(p.Y for p in poly_pts)
    maxy = max(p.Y for p in poly_pts)
    floor_z = poly_pts[0].Z

    margin_ft = SEARCH_MARGIN_M * FT_PER_M

    outline = DB.Outline(DB.XYZ(minx - margin_ft,
                                miny - margin_ft,
                                floor_z - 1000.0),
                         DB.XYZ(maxx + margin_ft,
                                maxy + margin_ft,
                                floor_z + 1000.0))
    bbfilter = DB.BoundingBoxIntersectsFilter(outline)

    col = DB.FilteredElementCollector(doc) \
             .OfCategory(DB.BuiltInCategory.OST_Planting) \
             .WhereElementIsNotElementType()

    created = 0

    for inst in col:
        try:
            try:
                if not bbfilter.PassesFilter(doc, inst.Id):
                    continue
            except:
                pass

            loc = inst.Location
            if not hasattr(loc, "Point") or loc.Point is None:
                continue
            pt = loc.Point

            sym = doc.GetElement(inst.GetTypeId())
            if not is_tree_symbol(sym):
                continue

            p = sym.LookupParameter(TRUNK_PARAM_NAME)
            if not p or not p.HasValue:
                continue

            diam_ft = p.AsDouble()
            if diam_ft <= 0.0:
                continue

            # add 100 mm to diameter
            extra_diam_ft = TREE_EXTRA_DIAM_MM / 304.8
            radius_ft = 0.5 * (diam_ft + extra_diam_ft)

            center = DB.XYZ(pt.X, pt.Y, target_z)

            arc1 = DB.Arc.Create(center, radius_ft, 0.0, math.pi,
                                 DB.XYZ.BasisX, DB.XYZ.BasisY)
            arc2 = DB.Arc.Create(center, radius_ft, math.pi, 2.0 * math.pi,
                                 DB.XYZ.BasisX, DB.XYZ.BasisY)

            line1 = doc.Create.NewAreaBoundaryLine(sketch_plane, arc1, view)
            if boundary_style is not None:
                try:
                    line1.LineStyle = boundary_style
                except:
                    pass

            line2 = doc.Create.NewAreaBoundaryLine(sketch_plane, arc2, view)
            if boundary_style is not None:
                try:
                    line2.LineStyle = boundary_style
                except:
                    pass

            created += 2

        except:
            continue

    logger.info("Created %s area-boundary arcs for tree trunks.", created)


# --------------- Main area boundary creation ---------------------

def create_area_boundaries_from_floor(doc, view, floor):
    """Create area boundary lines based on presence/absence of planting, plus tree trunks."""

    top_face = get_top_planar_face(floor)
    if top_face is None:
        TaskDialog.Show("Area Boundary From Floor",
                        "Could not find a horizontal top face for the selected floor.")
        return False

    loops = top_face.GetEdgesAsCurveLoops()
    if loops is None or loops.Count == 0:
        TaskDialog.Show("Area Boundary From Floor",
                        "Could not get perimeter loops for the selected floor.")
        return False

    # Choose outer loop as polygon – the one with greatest length
    outer_loop = None
    max_len = 0.0
    for cl in loops:
        length = 0.0
        for crv in cl:
            try:
                length += crv.ApproximateLength
            except:
                length += crv.Length
        if length > max_len:
            max_len = length
            outer_loop = cl

    if outer_loop is None:
        TaskDialog.Show("Area Boundary From Floor",
                        "Failed to determine outer perimeter for the selected floor.")
        return False

    poly_pts = [crv.GetEndPoint(0) for crv in outer_loop]
    if len(poly_pts) < 3:
        TaskDialog.Show("Area Boundary From Floor",
                        "Floor perimeter is too simple or invalid for boundary creation.")
        return False

    # Compute 2D region for deletion: exactly the floor extents (no margin),
    # so we don't touch mix boundaries outside this floor.
    minx = min(p.X for p in poly_pts)
    maxx = max(p.X for p in poly_pts)
    miny = min(p.Y for p in poly_pts)
    maxy = max(p.Y for p in poly_pts)

    reg_minx = minx
    reg_maxx = maxx
    reg_miny = miny
    reg_maxy = maxy

    # Ensure the view has a sketch plane
    sketch_plane = view.SketchPlane
    if sketch_plane is None:
        gen_level = view.GenLevel
        if gen_level:
            z = gen_level.Elevation
        else:
            z = 0.0
        plane = DB.Plane.CreateByNormalAndOrigin(DB.XYZ.BasisZ, DB.XYZ(0.0, 0.0, z))
        sketch_plane = DB.SketchPlane.Create(doc, plane)
        view.SketchPlane = sketch_plane

    if view.GenLevel:
        target_z = view.GenLevel.Elevation
    else:
        target_z = sketch_plane.GetPlane().Origin.Z

    # Get or create dedicated mix boundary line style
    boundary_style = get_mix_boundary_linestyle(doc)

    # First, delete any existing mix boundaries in this region
    delete_existing_boundaries_in_region(
        doc, view, reg_minx, reg_maxx, reg_miny, reg_maxy, boundary_style
    )

    # Try: create plant-inverse area boundary
    has_plant_boundary = create_plant_area_boundary(
        doc, view, outer_loop, target_z, sketch_plane, boundary_style
    )

    # Always draw tree trunk circles as area boundaries
    draw_tree_trunks(doc, view, outer_loop, target_z, sketch_plane, boundary_style)

    if has_plant_boundary:
        logger.info("Plant-inverse boundary created; skipping floor-outline boundary.")
        return True

    # No plants: default to floor outline offset 5 mm inside
    offset_dist_ft = FLOOR_OFFSET_MM / 304.8
    created_count = 0

    for cl in loops:
        offset_cl = offset_loop_inwards(cl, offset_dist_ft, DB.XYZ.BasisZ)
        for crv in offset_cl:
            flat_curve = flatten_curve_to_view_z(crv, target_z)
            line = doc.Create.NewAreaBoundaryLine(sketch_plane, flat_curve, view)
            if boundary_style is not None:
                try:
                    line.LineStyle = boundary_style
                except:
                    pass
            created_count += 1

    logger.info("Created %s floor-outline area boundary segments (offset 5mm inside).",
                created_count)

    TaskDialog.Show(
        "Area Boundary From Floor",
        "No non-tree plants found in this floor.\n"
        "Created {0} area boundary segments (offset 5 mm inside the floor edge).\n"
        "Tree trunk circles (area boundaries) were drawn where applicable."
        .format(created_count)
    )

    return True


# ---------------------------- Entry ------------------------------

def main(mix_name=None):
    uidoc = revit.uidoc
    doc = revit.doc
    view = doc.ActiveView

    if view.ViewType != DB.ViewType.AreaPlan:
        TaskDialog.Show("Area Boundary From Floor",
                        "Active view must be an Area Plan to create area boundary lines.")
        return

    try:
        pick = uidoc.Selection.PickPoint(
            ObjectSnapTypes.None,
            "Pick a point to find the floor below and place the Area."
        )
    except Exception:
        # User cancelled
        return

    floor = get_floor_below_point(doc, pick)
    if floor is None:
        TaskDialog.Show("Area Boundary From Floor",
                        "No floor was found below the picked point.")
        return

    t = DB.Transaction(doc, "Create / Update Area Boundaries and Area")
    t.Start()
    try:
        success = create_area_boundaries_from_floor(doc, view, floor)

        # Place an Area at the picked point and set Name if mix_name is given
        try:
            uv = DB.UV(pick.X, pick.Y)
            area = doc.Create.NewArea(view, uv)
            if mix_name and area:
                name_param = area.LookupParameter("Name")
                if name_param and name_param.StorageType == DB.StorageType.String:
                    name_param.Set(mix_name)
        except Exception as e:
            logger.error("Failed to create and/or name Area: %s", e)

        t.Commit()
    except Exception as e:
        logger.error("Error while creating boundaries / area: %s", e)
        t.RollBack()
        TaskDialog.Show("Area Boundary From Floor",
                        "Failed to create boundaries and Area.\n\n{0}".format(e))


if __name__ == "__main__":
    # Standalone use from a pyRevit button (no mix name)
    main()
