# coding: utf-8
"""Native Area Boundary drawing workflow for the Mix Schedule Editor.

This module intentionally owns the asynchronous Revit event workflow used by
script.py's "Draw Area" button. Revit posted commands execute after the current
API context returns, so DocumentChanged only records newly added boundary line
ids and Idling performs all model modifications once Revit returns to idle.
"""

from pyrevit import DB, forms
from Autodesk.Revit.UI import RevitCommandId, PostableCommand

try:
    from System import EventHandler
    from Autodesk.Revit.DB.Events import DocumentChangedEventArgs
    from Autodesk.Revit.UI.Events import IdlingEventArgs
    HAS_TYPED_REVIT_EVENTS = True
except Exception:
    EventHandler = None
    DocumentChangedEventArgs = None
    IdlingEventArgs = None
    HAS_TYPED_REVIT_EVENTS = False

_DRAW_AREA_SESSION = None
_DRAW_AREA_EVENTS_SUBSCRIBED = False
_DRAW_AREA_DOC_CHANGED_HANDLER = None
_DRAW_AREA_IDLING_HANDLER = None
_DRAW_AREA_EVENTS_UIAPP = None
DRAW_AREA_MAX_IDLE_WITHOUT_LINES = 25
DRAW_AREA_MAX_IDLE_WITH_LINES = 1
MIX_BOUNDARY_STYLE_NAME = u'BM Mix Boundary'
AREA_NAME_PARAM = u'Name'


def _to_unicode(value):
    if value is None:
        return u''
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u''


def _get_element_id_int(eid):
    if eid is None:
        return None
    try:
        return eid.IntegerValue
    except Exception:
        pass
    try:
        return eid.Value
    except Exception:
        pass
    try:
        return int(eid)
    except Exception:
        return None


def is_area_plan_view(view):
    """Return True when view can host native Area Boundary lines and Areas."""
    try:
        return isinstance(view, DB.ViewPlan) and view.ViewType == DB.ViewType.AreaPlan
    except Exception:
        try:
            return view.ViewType == DB.ViewType.AreaPlan
        except Exception:
            return False


def get_mix_boundary_linestyle(doc):
    """Get or create the BM Mix Boundary line style for Area Boundary lines."""
    cats = doc.Settings.Categories
    try:
        cat = cats.get_Item(DB.BuiltInCategory.OST_AreaSchemeLines)
    except Exception:
        cat = None
    if cat is None:
        return None

    for sub in cat.SubCategories:
        try:
            if sub.Name == MIX_BOUNDARY_STYLE_NAME:
                return sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
        except Exception:
            pass

    try:
        new_sub = cats.NewSubcategory(cat, MIX_BOUNDARY_STYLE_NAME)
        return new_sub.GetGraphicsStyle(DB.GraphicsStyleType.Projection)
    except Exception:
        return None


def _is_area_boundary_element(elem, view_id):
    """Best-effort check for Area Boundary / Area Scheme line elements."""
    if elem is None:
        return False
    try:
        if elem.OwnerViewId != view_id:
            return False
    except Exception:
        pass
    try:
        cat = elem.Category
        if cat is not None and cat.Id == DB.ElementId(DB.BuiltInCategory.OST_AreaSchemeLines):
            return True
    except Exception:
        pass
    try:
        bic = int(DB.BuiltInCategory.OST_AreaSchemeLines)
        if elem.Category is not None and _get_element_id_int(elem.Category.Id) == bic:
            return True
    except Exception:
        pass
    return False


def _collect_area_boundary_ids_in_view(doc, view_id):
    ids = set()
    try:
        elems = (DB.FilteredElementCollector(doc, view_id)
                 .OfCategory(DB.BuiltInCategory.OST_AreaSchemeLines)
                 .WhereElementIsNotElementType()
                 .ToElements())
        for elem in elems:
            try:
                ids.add(_get_element_id_int(elem.Id))
            except Exception:
                pass
    except Exception:
        pass
    return ids


def _get_curve_endpoints_xy(curve):
    try:
        p0 = curve.GetEndPoint(0)
        p1 = curve.GetEndPoint(1)
        return (p0.X, p0.Y), (p1.X, p1.Y)
    except Exception:
        return None, None


def _point_key(pt, tol=0.001):
    return (int(round(pt[0] / tol)), int(round(pt[1] / tol)))


def _loop_area(points):
    area = 0.0
    if len(points) < 3:
        return 0.0
    for i in range(len(points)):
        x1, y1 = points[i]
        x2, y2 = points[(i + 1) % len(points)]
        area += x1 * y2 - x2 * y1
    return area / 2.0


def _point_in_poly(pt, poly):
    x, y = pt
    inside = False
    n = len(poly)
    if n < 3:
        return False
    j = n - 1
    for i in range(n):
        xi, yi = poly[i]
        xj, yj = poly[j]
        if ((yi > y) != (yj > y)):
            try:
                x_cross = (xj - xi) * (y - yi) / (yj - yi) + xi
                if x < x_cross:
                    inside = not inside
            except Exception:
                pass
        j = i
    return inside


def _centroid_or_bbox_point(poly):
    area = _loop_area(poly)
    if abs(area) > 1e-9:
        cx = 0.0
        cy = 0.0
        for i in range(len(poly)):
            x0, y0 = poly[i]
            x1, y1 = poly[(i + 1) % len(poly)]
            cross = x0 * y1 - x1 * y0
            cx += (x0 + x1) * cross
            cy += (y0 + y1) * cross
        try:
            cx = cx / (6.0 * area)
            cy = cy / (6.0 * area)
            if _point_in_poly((cx, cy), poly):
                return (cx, cy)
        except Exception:
            pass
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    return ((min(xs) + max(xs)) / 2.0, (min(ys) + max(ys)) / 2.0)


def _sample_points_for_loop(poly):
    xs = [p[0] for p in poly]
    ys = [p[1] for p in poly]
    minx, maxx = min(xs), max(xs)
    miny, maxy = min(ys), max(ys)
    points = [_centroid_or_bbox_point(poly)]
    for ix in range(1, 4):
        for iy in range(1, 4):
            pt = (minx + (maxx - minx) * ix / 4.0,
                  miny + (maxy - miny) * iy / 4.0)
            if _point_in_poly(pt, poly):
                points.append(pt)
    return points



def _sample_points_from_boundary_ids(doc, id_ints):
    """Return practical placement candidates around newly drawn boundary lines.

    This supports the common workflow where the user draws only the missing
    side(s) of a region that is otherwise bounded by existing Area Boundary
    lines. In that case the newly drawn curves are not a closed loop by
    themselves, but Revit can still place an Area if we sample inside the
    completed region near the new work.
    """
    pts = []
    for id_int in id_ints:
        try:
            elem = doc.GetElement(DB.ElementId(id_int))
            curve = elem.GeometryCurve
        except Exception:
            continue
        p0, p1 = _get_curve_endpoints_xy(curve)
        if p0 is not None:
            pts.append(p0)
        if p1 is not None:
            pts.append(p1)

    if not pts:
        return []

    xs = [pt[0] for pt in pts]
    ys = [pt[1] for pt in pts]
    minx, maxx = min(xs), max(xs)
    miny, maxy = min(ys), max(ys)
    width = maxx - minx
    height = maxy - miny
    pad = max(width, height) * 0.25
    if pad < 1.0:
        pad = 1.0
    if width < 0.1:
        minx -= pad
        maxx += pad
    if height < 0.1:
        miny -= pad
        maxy += pad

    candidates = []
    seen = set()

    def _add(pt):
        key = _point_key(pt)
        if key not in seen:
            seen.add(key)
            candidates.append(pt)

    _add(((minx + maxx) / 2.0, (miny + maxy) / 2.0))
    for ix in range(1, 6):
        for iy in range(1, 6):
            _add((minx + (maxx - minx) * ix / 6.0,
                  miny + (maxy - miny) * iy / 6.0))
    return candidates


def _style_boundary_ids(doc, ids, style):
    if style is None:
        return
    for id_int in ids:
        try:
            elem = doc.GetElement(DB.ElementId(id_int))
            elem.LineStyle = style
        except Exception:
            pass

def _build_closed_loops_from_boundary_ids(doc, id_ints):
    """Build simple endpoint-connected loops from new Area Boundary curves."""
    segments = []
    for id_int in id_ints:
        try:
            elem = doc.GetElement(DB.ElementId(id_int))
            curve = elem.GeometryCurve
        except Exception:
            continue
        p0, p1 = _get_curve_endpoints_xy(curve)
        if p0 is None or p1 is None:
            continue
        if _point_key(p0) == _point_key(p1):
            continue
        segments.append([p0, p1, False])

    loops = []
    for seg in segments:
        if seg[2]:
            continue
        seg[2] = True
        loop = [seg[0], seg[1]]
        start_key = _point_key(seg[0])
        end_key = _point_key(seg[1])
        changed = True
        while changed and end_key != start_key:
            changed = False
            for other in segments:
                if other[2]:
                    continue
                ok0 = _point_key(other[0])
                ok1 = _point_key(other[1])
                if ok0 == end_key:
                    other[2] = True
                    loop.append(other[1])
                    end_key = ok1
                    changed = True
                    break
                if ok1 == end_key:
                    other[2] = True
                    loop.append(other[0])
                    end_key = ok0
                    changed = True
                    break
        if end_key == start_key and len(loop) >= 4:
            loop = loop[:-1]
            if abs(_loop_area(loop)) > 1e-6:
                loops.append(loop)
    loops.sort(key=lambda pts: abs(_loop_area(pts)), reverse=True)
    return loops


def _area_name_matches(area, mix_name):
    expected = _to_unicode(mix_name).strip()
    try:
        p = area.LookupParameter(AREA_NAME_PARAM)
        if p:
            current = _to_unicode(p.AsString()).strip()
            if current == expected:
                return True
            current = _to_unicode(p.AsValueString()).strip()
            if current == expected:
                return True
    except Exception:
        pass
    try:
        current = _to_unicode(area.Name).strip()
        if current == expected:
            return True
    except Exception:
        pass
    return False


def _try_set_parameter(param, value):
    if param is None:
        return False
    try:
        if param.IsReadOnly:
            return False
    except Exception:
        pass
    try:
        param.Set(value)
        return True
    except Exception:
        return False


def _set_area_name(area, mix_name, set_area_name_callback=None):
    """Set an Area name robustly across Revit versions/templates.

    The callback from script.py uses its existing set_param helper, but that
    helper intentionally swallows failures. Always verify the result and then
    fall back to common Area/SpatialElement name parameters if needed.
    """
    mix_name = _to_unicode(mix_name).strip()
    if not mix_name:
        return False

    if set_area_name_callback is not None:
        try:
            set_area_name_callback(area, mix_name)
        except Exception:
            pass
        if _area_name_matches(area, mix_name):
            return True

    # Area name is usually the SpatialElement/Room name built-in parameter.
    for bip_name in ('ROOM_NAME', 'SPACE_NAME'):
        try:
            bip = getattr(DB.BuiltInParameter, bip_name)
            if _try_set_parameter(area.get_Parameter(bip), mix_name):
                if _area_name_matches(area, mix_name):
                    return True
                return True
        except Exception:
            pass

    try:
        if _try_set_parameter(area.LookupParameter(AREA_NAME_PARAM), mix_name):
            return True
    except Exception:
        pass

    return _area_name_matches(area, mix_name)


def _progress_update(progress_bar, value, max_value):
    if progress_bar is None:
        return
    try:
        progress_bar.update_progress(value, max_value)
    except Exception:
        pass



def _make_revit_event_handler(callback, args_type):
    """Create a typed .NET event handler when Revit/IronPython requires one."""
    if HAS_TYPED_REVIT_EVENTS and EventHandler is not None and args_type is not None:
        try:
            return EventHandler[args_type](callback)
        except Exception:
            pass
    return callback


def _exception_message(ex):
    msg = _to_unicode(ex)
    try:
        inner = ex.InnerException
    except Exception:
        inner = None
    if inner is not None:
        inner_msg = _to_unicode(inner)
        if inner_msg and inner_msg not in msg:
            msg = msg + u'\n' + inner_msg
    return msg


class DrawAreaSession(object):
    """State kept alive while Revit's posted Area Boundary command runs."""

    def __init__(self, doc, uidoc, uiapp, view_id, mix_name, existing_ids, reopen_callback, set_area_name_callback):
        self.doc = doc
        self.uidoc = uidoc
        self.uiapp = uiapp
        self.view_id = view_id
        self.mix_name = mix_name
        self.existing_ids = existing_ids or set()
        self.reopen_callback = reopen_callback
        self.set_area_name_callback = set_area_name_callback
        self.added_ids = set()
        self.idle_attempts = 0
        self.idle_with_lines = 0
        self.doc_changed_handler = None
        self.idling_handler = None
        self.completed = False


def _ensure_draw_area_events(uiapp):
    """Subscribe persistent Revit handlers once for this module instance.

    Re-subscribing after the modal editor has been reopened from a previous
    Idling run can throw a generic Revit target-invocation exception. Keeping
    one handler pair installed avoids the repeated add/remove cycle; the
    handlers are inert whenever _DRAW_AREA_SESSION is None.
    """
    global _DRAW_AREA_EVENTS_SUBSCRIBED
    global _DRAW_AREA_DOC_CHANGED_HANDLER
    global _DRAW_AREA_IDLING_HANDLER
    global _DRAW_AREA_EVENTS_UIAPP

    if _DRAW_AREA_EVENTS_SUBSCRIBED and _DRAW_AREA_EVENTS_UIAPP is uiapp:
        return

    doc_handler = _make_revit_event_handler(
        _draw_area_document_changed,
        DocumentChangedEventArgs
    )
    idling_handler = _make_revit_event_handler(
        _draw_area_idling,
        IdlingEventArgs
    )

    added_doc_handler = False
    try:
        uiapp.Application.DocumentChanged += doc_handler
        added_doc_handler = True
        uiapp.Idling += idling_handler
    except Exception:
        if added_doc_handler:
            try:
                uiapp.Application.DocumentChanged -= doc_handler
            except Exception:
                pass
        raise

    _DRAW_AREA_DOC_CHANGED_HANDLER = doc_handler
    _DRAW_AREA_IDLING_HANDLER = idling_handler
    _DRAW_AREA_EVENTS_UIAPP = uiapp
    _DRAW_AREA_EVENTS_SUBSCRIBED = True


def _draw_area_unsubscribe(session):
    # Handlers are intentionally kept subscribed for the module lifetime. They
    # immediately return when _DRAW_AREA_SESSION is None, and avoiding repeated
    # event add/remove cycles prevents the second Draw Area run subscription
    # failure seen after returning to the modal Mix Editor.
    return


def _draw_area_reopen_editor(session):
    try:
        if session is not None and session.reopen_callback is not None:
            session.reopen_callback(session.doc)
    except Exception as ex:
        forms.alert(u'Could not reopen the Mix Schedule Editor:\n{0}'.format(_exception_message(ex)),
                    title='Draw Area')


def _draw_area_finish(session, reopen=True):
    global _DRAW_AREA_SESSION
    if session is None:
        return
    _draw_area_unsubscribe(session)
    session.completed = True
    if _DRAW_AREA_SESSION is session:
        _DRAW_AREA_SESSION = None
    if reopen:
        _draw_area_reopen_editor(session)


def _draw_area_document_changed(sender, args):
    """Record Area Boundary ids only; never modify the model in this event."""
    session = _DRAW_AREA_SESSION
    if session is None or session.completed:
        return
    try:
        if args.GetDocument() != session.doc:
            return
    except Exception:
        return
    try:
        added = args.GetAddedElementIds()
    except Exception:
        return
    for eid in added:
        try:
            elem = session.doc.GetElement(eid)
        except Exception:
            elem = None
        if _is_area_boundary_element(elem, session.view_id):
            id_int = _get_element_id_int(eid)
            if id_int is not None:
                session.added_ids.add(id_int)


def _draw_area_apply_results(session, progress_bar=None):
    doc = session.doc
    view = doc.GetElement(session.view_id)
    if not is_area_plan_view(view):
        forms.alert(u'The active Area Plan is no longer available. The boundary lines were left in place.',
                    title='Draw Area')
        return

    _progress_update(progress_bar, 1, 4)

    ids = set(session.added_ids)
    if not ids:
        current = _collect_area_boundary_ids_in_view(doc, session.view_id)
        ids = current.difference(session.existing_ids)
    if not ids:
        return

    loops = _build_closed_loops_from_boundary_ids(doc, ids)
    _progress_update(progress_bar, 2, 4)
    if not loops:
        # The newly drawn lines might be intentionally open because they close
        # a region together with existing Area Boundary lines. Style the new
        # lines, then ask Revit to place an Area at sampled points around the
        # new work; Revit will accept a point if the total boundary network is
        # enclosed even when the new lines are not a closed loop by themselves.
        placed_count = 0
        candidates = _sample_points_from_boundary_ids(doc, ids)
        t = DB.Transaction(doc, 'Place Drawn Mix Area')
        try:
            t.Start()
            style = get_mix_boundary_linestyle(doc)
            _style_boundary_ids(doc, ids, style)
            for x, y in candidates:
                try:
                    area = doc.Create.NewArea(view, DB.UV(x, y))
                    if area is not None:
                        _set_area_name(area, session.mix_name, session.set_area_name_callback)
                        placed_count += 1
                        break
                except Exception:
                    pass
            t.Commit()
        except Exception:
            try:
                if t.HasStarted() and not t.HasEnded():
                    t.RollBack()
            except Exception:
                pass

        if placed_count <= 0:
            forms.alert(u'Area Boundary lines were drawn and styled, but Revit did not find an enclosed '
                        u'Area at the sampled points. The lines were left in place; if they connect to '
                        u'existing boundaries, place the Area manually or draw a little farther into the region.',
                        title='Draw Area')
        return

    placed_count = 0
    failed_loops = 0
    t = DB.Transaction(doc, 'Place Drawn Mix Area')
    try:
        t.Start()
        style = get_mix_boundary_linestyle(doc)
        _style_boundary_ids(doc, ids, style)

        loop_index = 0
        loop_count = len(loops)
        total_steps = 3 + loop_count + 1
        _progress_update(progress_bar, 3, total_steps)

        for loop in loops:
            loop_index += 1
            placed = False
            for x, y in _sample_points_for_loop(loop):
                try:
                    area = doc.Create.NewArea(view, DB.UV(x, y))
                    if area is not None:
                        _set_area_name(area, session.mix_name, session.set_area_name_callback)
                        placed_count += 1
                        placed = True
                        break
                except Exception:
                    pass
            if not placed:
                failed_loops += 1
            _progress_update(progress_bar, 3 + loop_index, total_steps)

        t.Commit()
        _progress_update(progress_bar, total_steps, total_steps)
    except Exception as ex:
        try:
            if t.HasStarted() and not t.HasEnded():
                t.RollBack()
        except Exception:
            pass
        forms.alert(u'Area Boundary lines were drawn, but the Area could not be placed:\n{0}\n\n'
                    u'The boundary lines were left in place.'.format(ex),
                    title='Draw Area')
        return

    if placed_count <= 0:
        forms.alert(u'Area Boundary lines were drawn, but Revit did not accept any sampled placement point. '
                    u'The boundary lines were left in place; please place the Area manually.',
                    title='Draw Area')
    elif failed_loops > 0:
        forms.alert(u'Placed {0} Area(s), but {1} detected loop(s) could not receive an Area. '
                    u'Those boundary lines were left in place.'.format(placed_count, failed_loops),
                    title='Draw Area')


def _draw_area_idling(sender, args):
    """Process recorded ids after Revit returns to idle from the posted command."""
    # Once Revit gives us an Idling callback, ask it to keep raising Idling
    # without waiting for extra user input. This cannot force the first idle
    # after the native command, but it avoids an additional click/key press once
    # Revit has yielded back to the API.
    try:
        args.SetRaiseWithoutDelay()
    except Exception:
        pass

    session = _DRAW_AREA_SESSION
    if session is None or session.completed:
        return

    if session.added_ids:
        session.idle_with_lines += 1
        if session.idle_with_lines < DRAW_AREA_MAX_IDLE_WITH_LINES:
            return
    else:
        session.idle_attempts += 1
        if session.idle_attempts < DRAW_AREA_MAX_IDLE_WITHOUT_LINES:
            return

    try:
        progress_ctx = None
        try:
            progress_ctx = forms.ProgressBar(
                title='Draw Area: processing boundary ({value} of {max_value})',
                cancellable=False
            )
        except Exception:
            progress_ctx = None

        if progress_ctx is None:
            # If the pyRevit progress UI is unavailable for any reason, still
            # complete the Revit work and reopen the editor.
            _draw_area_apply_results(session, None)
        else:
            with progress_ctx as progress_bar:
                _draw_area_apply_results(session, progress_bar)
    finally:
        _draw_area_finish(session, reopen=True)


def start_draw_area_session(doc, uidoc, uiapp, view, mix_name, close_callback, reopen_callback, set_area_name_callback=None):
    """Validate, subscribe event handlers, close editor, and post Area Boundary."""
    global _DRAW_AREA_SESSION

    mix_name = _to_unicode(mix_name).strip()
    if not mix_name:
        forms.alert(u'This mix does not have a name, so an Area cannot be named.',
                    title='Draw Area')
        return False

    if not is_area_plan_view(view):
        forms.alert(u'Draw Area must be run from an Area Plan view.\n\n'
                    u'Open the relevant Area Plan, expand the mix, and click Draw Area again.',
                    title='Draw Area')
        return False

    if uiapp is None:
        forms.alert(u'Could not access the Revit UI application to start Draw Area.',
                    title='Draw Area')
        return False

    if _DRAW_AREA_SESSION is not None:
        try:
            _draw_area_finish(_DRAW_AREA_SESSION, reopen=False)
        except Exception:
            _DRAW_AREA_SESSION = None

    existing_ids = _collect_area_boundary_ids_in_view(doc, view.Id)
    session = DrawAreaSession(doc, uidoc, uiapp, view.Id, mix_name, existing_ids, reopen_callback, set_area_name_callback)

    try:
        _ensure_draw_area_events(uiapp)
    except Exception as ex:
        forms.alert(u'Could not subscribe to Revit events for Draw Area:\n{0}'.format(_exception_message(ex)),
                    title='Draw Area')
        return False

    _DRAW_AREA_SESSION = session

    try:
        if close_callback is not None:
            close_callback()
    except Exception:
        pass

    try:
        cmd_id = RevitCommandId.LookupPostableCommandId(PostableCommand.AreaBoundary)
        if cmd_id is None:
            raise Exception('Area Boundary command id was not found.')
        uiapp.PostCommand(cmd_id)
    except Exception as ex:
        _draw_area_finish(session, reopen=False)
        forms.alert(u'Could not start Revit Area Boundary drawing command:\n{0}'.format(_exception_message(ex)),
                    title='Draw Area')
        _draw_area_reopen_editor(session)
        return False

    return True
