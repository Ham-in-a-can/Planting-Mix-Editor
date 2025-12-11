# coding: utf-8
"""pyRevit Mix Schedule Editor + Area Colours + Filled Region strips.

Reads and edits Generic Annotation instances for planting mix schedules.
Family: bm_massed_mix (configurable via FAMILY_NAME).

Adds:
- Link to an Area Color Fill Scheme (MASSED PLANTING by default).
- Per-mix colour swatch in the UI.
- Palette + colour picker per mix.
- Writes colours back to the colour scheme on Apply.
- Creates/updates FilledRegionTypes "bm_planting_<MixName>" and a strip in
  the target view at (0,0).
"""

import os
import sys
import imp
import clr

from pyrevit import revit, DB, script, forms

from Autodesk.Revit.DB import *
from Autodesk.Revit.UI import *
from Autodesk.Revit.DB import BuiltInParameter


# WPF / .NET
clr.AddReference('PresentationCore')
clr.AddReference('PresentationFramework')
clr.AddReference('WindowsBase')
clr.AddReference('System.Windows.Forms')
clr.AddReference('System.Drawing')

import System
from System.Windows import (
    Thickness, Visibility, CornerRadius, GridLength, GridUnitType,
    HorizontalAlignment, VerticalAlignment, FontWeights,
    SizeToContent, WindowStartupLocation, Window, TextWrapping, Point
)
from System.Windows.Controls import (
    StackPanel, Grid, TextBlock, Button, Border,
    ColumnDefinition, TextBox, Orientation, Image, WrapPanel,
    ScrollViewer, ScrollBarVisibility
)
from System.Windows.Input import Cursors
import System.Windows.Media as Media
from System.IO import FileStream, FileMode

from System.Windows.Markup import XamlReader
from System.Windows.Media.Imaging import BitmapImage
from System import Uri, DateTime
import System.Windows.Forms as WinForms
import System.Drawing as Drawing
from System.Collections.Generic import List

try:
    from Autodesk.Revit.UI import ColorSelectionDialog
    HAS_REVIT_COLOR_DIALOG = True
except Exception:
    HAS_REVIT_COLOR_DIALOG = False

LOGGER = script.get_logger()

# -----------------------------
# Global configuration
# -----------------------------
FAMILY_NAME = 'bm_massed_mix'            # Generic Annotation family to target
MAX_SPECIES = 15                         # Max species rows supported

# UI behaviour for long Botanical / Common names
# NAME_FADE_FRACTION = fraction (0–1) of cell width used for fade at right edge
NAME_FADE_FRACTION = 0.33
NAME_TOOLTIP_MIN_CHARS = 28
NAME_FADE_WIDTH = 75      # pixels of fade overlay at right edge



# Fallback drafting view name for mix schedules
DEFAULT_MIX_DRAFTING_VIEW_NAME = 'Massed Planting Mixes'

# Icon used for the "Duplicate mix" button (in the pushbutton folder)
DUPLICATE_ICON_NAME = 'duplicate_icon_hover.png'

PARAM_MIX_NAME = 'Mix Name'
PARAM_NUM_SPECIES = 'Num_Species'

# Row parameter name templates
PARAM_SPECIES_CODE_TEMPLATE  = 'S{0}_Code'   # S1_Code, S2_Code, ...
PARAM_SPECIES_PCT_TEMPLATE   = 'S{0}_Pct'    # S1_Pct, S2_Pct, ...
PARAM_SPECIES_SPACE_TEMPLATE = 'S{0}_Space'  # S1_Space, S2_Space, ...
PARAM_SPECIES_BOT_TEMPLATE   = 'S{0}_Bot'    # S1_Bot, S2_Bot, ...
PARAM_SPECIES_COM_TEMPLATE   = 'S{0}_Com'    # S1_Com, S2_Com, ...
PARAM_SPECIES_GRADE_TEMPLATE = 'S{0}_Grade'  # S1_Grade, S2_Grade, ...

# Area name parameter (instance) for all Areas
AREA_NAME_PARAM = 'Name'

# --- Area Color Fill Scheme configuration ---
COLOR_SCHEME_NAME = u'MASSED PLANTING'                 # Name shown in Edit Color Scheme
COLOR_SCHEME_CATEGORY = DB.BuiltInCategory.OST_Areas   # Category for scheme

# Default area colours for quick selection
DEFAULT_AREA_COLORS = [
    (u'Soft Green',   DB.Color(198, 224, 180)),
    (u'Mid Green',    DB.Color(169, 209, 142)),
    (u'Olive',        DB.Color(143, 188, 143)),
    (u'Blue',         DB.Color(189, 215, 238)),
    (u'Yellow',       DB.Color(255, 242, 204)),
    (u'Orange',       DB.Color(252, 213, 180)),
    (u'Red',          DB.Color(244, 199, 195)),
    (u'Grey',         DB.Color(217, 217, 217)),
    (u'Dark Green',   DB.Color(0,   97,  0)),
    (u'Deep Blue',    DB.Color(0,   112, 192)),
    (u'Purple',       DB.Color(112, 48,  160)),
    (u'Brown',        DB.Color(150, 75,  0)),
    (u'Light Grey',   DB.Color(242, 242, 242)),
]

# Filled Region strip configuration (schedule colour bars)
FILLED_REGION_WIDTH_MM = 20.0
FILLED_REGION_HEIGHT_MM = 4.0

# Debug: set True while trying to understand filled region behaviour
DEBUG_FILLED_REGION = False

#Path to plant library script
CREATE_PLANT_SCRIPT_PATH = r"C:\Users\hamishc\OneDrive - Boffa Miskell\Desktop\Shared_Dynamo\ToolBar\OnServer\BoffaTestTools.extension\BoffaTools.tab\Test.panel\Create Plant.pushbutton\script.py"


# Path to XAML file (placed next to this script)
try:
    SCRIPT_DIR = os.path.dirname(__file__)
except Exception:
    try:
        from pyrevit import script as _script_mod
        SCRIPT_DIR = os.path.dirname(_script_mod.get_script_path())
    except Exception:
        SCRIPT_DIR = os.getcwd()

XAML_FILE = os.path.join(SCRIPT_DIR, 'mix_schedules.xaml')
DUPLICATE_ICON_PATH = os.path.join(SCRIPT_DIR, DUPLICATE_ICON_NAME)


# -----------------------------
# Helpers
# -----------------------------

# Debug log for filled region operations (shown in a popup on Apply)
FR_DEBUG_LINES = []


def fr_debug(msg):
    """Collect debug info for filled regions and also send to pyRevit logger."""
    if not DEBUG_FILLED_REGION:
        return
    try:
        LOGGER.info(msg)
    except Exception:
        pass
    try:
        FR_DEBUG_LINES.append(msg)
    except Exception:
        pass
from Autodesk.Revit.DB import BuiltInParameter  # if not already imported


def _get_element_id_int(eid):
    """Return a best-effort integer value for a Revit ElementId across versions."""
    if eid is None:
        return None

    try:
        return eid.IntegerValue
    except Exception:
        pass

    # Revit 2026+ exposes ElementId.Value instead of IntegerValue
    try:
        return eid.Value
    except Exception:
        pass

    try:
        return int(eid)
    except Exception:
        return None


def _ensure_elementid_integervalue_alias():
    """Expose ElementId.IntegerValue on newer Revit versions that use Value."""
    try:
        # If the attribute already exists, nothing to do.
        if hasattr(DB.ElementId, 'IntegerValue'):
            return

        def _get_integer_value(self):
            try:
                return self.Value
            except Exception:
                try:
                    return int(self)
                except Exception:
                    return None

        # Attach a Python property as a compatibility alias.
        try:
            DB.ElementId.IntegerValue = property(_get_integer_value)
        except Exception:
            # Fallback: attach a simple attribute if property assignment fails.
            setattr(DB.ElementId, 'IntegerValue', _get_integer_value)
    except Exception:
        # Best-effort; if we cannot patch, leave the environment unchanged.
        pass


# Ensure ElementId exposes IntegerValue for downstream scripts like Boundary.py
_ensure_elementid_integervalue_alias()


def get_element_name(elem):
    """Safely get an element/type name across Revit/IronPython versions."""
    # Try standard Name property
    try:
        name = elem.Name
        if name:
            return name
    except Exception:
        pass

    # Fallback to SYMBOL_NAME_PARAM
    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            name = p.AsString()
            if name:
                return name
    except Exception:
        pass

    return u'<no name>'

from Autodesk.Revit.UI import TaskDialog  # already imported via *, but explicit is fine


def show_fr_debug_popup():
    """Show a TaskDialog (or forms.alert) with filled region debug info."""
    if not DEBUG_FILLED_REGION:
        return
    if not FR_DEBUG_LINES:
        return

    msg = u'\n'.join(FR_DEBUG_LINES)
    try:
        TaskDialog.Show('Filled Region Debug', msg)
    except Exception:
        try:
            forms.alert(msg, title='Filled Region Debug')
        except Exception:
            pass

    # clear for next run
    del FR_DEBUG_LINES[:]


def get_element_name(elem):
    """Safely get an element/type name across Revit/IronPython versions."""
    # Try standard Name property
    try:
        name = elem.Name
        if name:
            return name
    except Exception:
        pass

    # Fallback to SYMBOL_NAME_PARAM
    try:
        p = elem.get_Parameter(BuiltInParameter.SYMBOL_NAME_PARAM)
        if p:
            name = p.AsString()
            if name:
                return name
    except Exception:
        pass

    return u'<no name>'


def _to_unicode(value):
    if value is None:
        return u''
    try:
        if isinstance(value, unicode):
            return value
    except Exception:
        pass
    try:
        return unicode(value)
    except Exception:
        try:
            return str(value)
        except Exception:
            return u''


def get_param(element, name):
    """Safe parameter getter returning a Python string or None."""
    try:
        p = element.LookupParameter(name)
    except Exception:
        p = None
    if not p or not p.HasValue:
        return None
    try:
        if p.StorageType == DB.StorageType.String:
            return p.AsString()
        elif p.StorageType == DB.StorageType.Integer:
            return str(p.AsInteger())
        elif p.StorageType == DB.StorageType.Double:
            try:
                return p.AsValueString()
            except Exception:
                return str(p.AsDouble())
        else:
            return p.AsValueString()
    except Exception:
        return None


def set_param(element, name, value):
    """Safe parameter setter that handles string/int/double where reasonable.

    NOTE: For Double length parameters we often want to set feet directly.
    For spacing we handle that explicitly in on_apply; here we assume the
    caller passes feet for Double types.
    """
    try:
        p = element.LookupParameter(name)
    except Exception:
        p = None
    if not p:
        return
    try:
        if p.StorageType == DB.StorageType.String:
            p.Set(value if value is not None else u'')
        elif p.StorageType == DB.StorageType.Integer:
            if value in (None, u'', ''):
                ival = 0
            else:
                try:
                    ival = int(value)
                except Exception:
                    ival = 0
            p.Set(ival)
        elif p.StorageType == DB.StorageType.Double:
            if value in (None, u'', ''):
                dval = 0.0
            else:
                try:
                    dval = float(value)
                except Exception:
                    try:
                        dval = float(str(value).replace(',', '.'))
                    except Exception:
                        dval = 0.0
            p.Set(dval)
        else:
            if hasattr(p, 'SetValueString'):
                try:
                    p.SetValueString(str(value))
                except Exception:
                    pass
    except Exception:
        # swallow errors rather than killing the UI
        pass


def copy_param_between_elements(src_elem, dst_elem, param_name):
    """Copy a single parameter value from src_elem to dst_elem by name."""
    try:
        p_src = src_elem.LookupParameter(param_name)
        p_dst = dst_elem.LookupParameter(param_name)
    except Exception:
        p_src = None
        p_dst = None
    if not p_src or not p_dst:
        return

    try:
        st = p_src.StorageType
        if st == DB.StorageType.String:
            p_dst.Set(p_src.AsString())
        elif st == DB.StorageType.Integer:
            p_dst.Set(p_src.AsInteger())
        elif st == DB.StorageType.Double:
            p_dst.Set(p_src.AsDouble())
        else:
            try:
                val = p_src.AsValueString()
                if val is not None and hasattr(p_dst, 'SetValueString'):
                    p_dst.SetValueString(val)
            except Exception:
                pass
    except Exception:
        pass


# ---- Color fill helpers (Area colour swatches) ----
def _get_entry_color(entry):
    """Try to read a Revit.DB.Color from a ColorFillSchemeEntry."""
    if entry is None:
        return None
    for attr in ('BackgroundColor', 'Color', 'ForegroundColor'):
        try:
            if hasattr(entry, attr):
                c = getattr(entry, attr)
                if c is not None:
                    return c
        except Exception:
            continue
    return None


def _set_entry_color(entry, color):
    """Try to set a Revit.DB.Color on a ColorFillSchemeEntry."""
    if entry is None or color is None:
        return
    for attr in ('BackgroundColor', 'Color', 'ForegroundColor'):
        try:
            if hasattr(entry, attr):
                setattr(entry, attr, color)
                return
        except Exception:
            continue


def _get_color_entry_keys(entry):
    """Return (value_key, caption_key) strings for a ColorFillSchemeEntry."""
    val_key = u''
    cap_key = u''

    try:
        st = entry.StorageType
    except Exception:
        st = None

    if st == DB.StorageType.String:
        try:
            val_key = entry.GetStringValue()
        except Exception:
            val_key = u''
    elif st == DB.StorageType.Integer:
        try:
            ival = entry.GetIntegerValue()
            val_key = str(ival)
        except Exception:
            val_key = u''
    elif st == DB.StorageType.Double:
        try:
            dval = entry.GetDoubleValue()
            val_key = str(dval)
        except Exception:
            val_key = u''
    else:
        try:
            eid = entry.GetElementIdValue()
            eid_int = _get_element_id_int(eid)
            if eid_int not in (None, -1):
                val_key = str(eid_int)
        except Exception:
            val_key = u''

    val_key = _to_unicode(val_key).strip()

    try:
        cap_key = _to_unicode(entry.Caption).strip()
    except Exception:
        cap_key = u''

    return val_key, cap_key


def _dbcolor_to_media_brush(db_color):
    """Convert Revit.DB.Color to a WPF SolidColorBrush for UI preview."""
    try:
        if db_color is None:
            return Media.SolidColorBrush(Media.Colors.LightGray)
        r = db_color.Red
        g = db_color.Green
        b = db_color.Blue
        c = Media.Color.FromRgb(r, g, b)
        return Media.SolidColorBrush(c)
    except Exception:
        return Media.SolidColorBrush(Media.Colors.LightGray)


def pick_color_with_revit_dialog(initial_dbcolor):
    """Try Revit's ColorSelectionDialog, then WinForms ColorDialog. Returns DB.Color or None."""
    if HAS_REVIT_COLOR_DIALOG:
        try:
            dlg = ColorSelectionDialog()
            if initial_dbcolor is not None:
                try:
                    dlg.SelectedColor = initial_dbcolor
                except Exception:
                    pass
            try:
                ColorSelectionDialog.Show(dlg)
            except Exception:
                try:
                    dlg.Show()
                except Exception:
                    pass
            col = dlg.SelectedColor
            if col is not None and col.IsValidObject:
                return col
        except Exception:
            pass

    # Fallback: WinForms
    try:
        cd = WinForms.ColorDialog()
    except Exception:
        return None

    if initial_dbcolor is not None:
        try:
            cd.Color = Drawing.Color.FromArgb(
                initial_dbcolor.Red,
                initial_dbcolor.Green,
                initial_dbcolor.Blue
            )
        except Exception:
            pass

    result = cd.ShowDialog()
    if result != WinForms.DialogResult.OK:
        return None

    col = cd.Color
    try:
        return DB.Color(col.R, col.G, col.B)
    except Exception:
        return None


def pick_area_color_with_palette(initial_dbcolor):
    """Small palette window + 'More colours...' button. Returns DB.Color or None."""
    selected = {'color': None}

    win = Window()
    win.Title = u'Select Area Colour'
    win.SizeToContent = SizeToContent.WidthAndHeight
    win.WindowStartupLocation = WindowStartupLocation.CenterScreen
    win.ResizeMode = 0  # NoResize

    root = StackPanel()
    root.Margin = Thickness(10)
    win.Content = root

    label = TextBlock()
    label.Text = u'Quick colours'
    label.FontSize = 12
    label.Margin = Thickness(0, 0, 0, 6)
    root.Children.Add(label)

    wrap = WrapPanel()
    wrap.Margin = Thickness(0, 0, 0, 8)
    root.Children.Add(wrap)

    for name, dbcol in DEFAULT_AREA_COLORS:
        sw = Border()
        sw.Width = 22
        sw.Height = 22
        sw.Margin = Thickness(2)
        sw.CornerRadius = CornerRadius(3)
        sw.BorderThickness = Thickness(1)
        sw.BorderBrush = Media.Brushes.Gray
        sw.Background = _dbcolor_to_media_brush(dbcol)
        sw.ToolTip = name

        def on_click(sender, args, dbcol=dbcol):
            selected['color'] = dbcol
            try:
                win.DialogResult = True
            except Exception:
                pass
            win.Close()

        sw.MouseLeftButtonUp += on_click
        wrap.Children.Add(sw)

    more_btn = Button()
    more_btn.Content = u'More colours...'
    more_btn.Margin = Thickness(0, 4, 0, 0)
    more_btn.Padding = Thickness(6, 2, 6, 2)

    def on_more(sender, args):
        col = pick_color_with_revit_dialog(initial_dbcolor)
        if col is not None:
            selected['color'] = col
            try:
                win.DialogResult = True
            except Exception:
                pass
        else:
            try:
                win.DialogResult = False
            except Exception:
                pass
        win.Close()

    more_btn.Click += on_more
    root.Children.Add(more_btn)

    try:
        result = win.ShowDialog()
    except Exception:
        result = None

    if result:
        return selected['color']
    return None


# ---- Filled Region helpers ----
def _sanitize_mix_name_for_type(mix_name):
    base = _to_unicode(mix_name).strip()
    if not base:
        base = u'Unnamed'
    for ch in u'<>:"/\|?*':
        base = base.replace(ch, u'_')
    return base


def _get_template_filled_region_type(doc):
    """Return a FilledRegionType to use as a template."""
    try:
        col = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType)
    except Exception:
        return None

    for t in col:
        try:
            if _to_unicode(t.Name) == u'bm_planting_template':
                return t
        except Exception:
            continue

    for t in col:
        return t

    return None


def _set_filled_region_type_color(fr_type, color):
    """Set the foreground colour on a FilledRegionType, like the working test script."""
    if fr_type is None or color is None:
        return

    # Try to turn the foreground pattern on, if supported
    try:
        if hasattr(fr_type, "IsForegroundPatternVisible"):
            fr_type.IsForegroundPatternVisible = True
    except Exception:
        pass

    before_color = None
    try:
        before_color = fr_type.ForegroundPatternColor
    except Exception:
        pass

    try:
        fr_type.ForegroundPatternColor = color

        if DEBUG_FILLED_REGION:
            before_txt = (
                u'<None>' if before_color is None
                else u'RGB({0},{1},{2})'.format(
                    before_color.Red,
                    before_color.Green,
                    before_color.Blue
                )
            )
            after_txt = u'RGB({0},{1},{2})'.format(color.Red, color.Green, color.Blue)
            fr_debug(
                u'FR: Updated ForegroundPatternColor on type "{0}" (Id {1}) '
                u'from {2} to {3}.'.format(
                    get_element_name(fr_type),
                    _get_element_id_int(fr_type.Id),
                    before_txt,
                    after_txt
                )
            )
    except Exception as ex:
        fr_debug(
            u'FR: Failed to set ForegroundPatternColor on type "{0}" (Id {1}): {2}'
            .format(
                get_element_name(fr_type),
                _get_element_id_int(fr_type.Id),
                ex
            )
        )



def _ensure_filled_region_for_mix(doc, mix_name, color, host_view):
    """Ensure a bm_planting_<MixName> FilledRegionType + one strip exist.

    - Reuses an existing FilledRegionType if it exists.
    - Only duplicates the template when needed.
    - If duplication fails with "name already in use", tries to pick up the
      existing type anyway.
    - Logs everything via fr_debug so we can show it in a popup.
    """
    if color is None or host_view is None:
        fr_debug(
            u'FR: Skipping mix "{0}" because color or host_view is None.'
            .format(mix_name)
        )
        return

    safe_name = _sanitize_mix_name_for_type(mix_name)
    type_name = u'bm_planting_{0}'.format(safe_name)

    fr_debug(
        u'FR: Ensure type "{0}" for mix "{1}" in view "{2}" (Id {3}) RGB({4},{5},{6})'
        .format(
            type_name,
            mix_name,
            _to_unicode(host_view.Name),
            _get_element_id_int(host_view.Id),
            color.Red, color.Green, color.Blue
        )
    )

    # --- 1. Try to find an existing FilledRegionType with this name ---
    fr_type = None
    try:
        col_types = DB.FilteredElementCollector(doc).OfClass(DB.FilledRegionType)
    except Exception:
        col_types = []

    for t in col_types:
        try:
            if get_element_name(t) == type_name:
                fr_type = t
                fr_debug(
                    u'FR: Found existing FilledRegionType "{0}" (Id {1}).'
                    .format(type_name, _get_element_id_int(t.Id))
                )
                break
        except Exception:
            continue

    # --- 2. If none found, try duplicating a template ---
    if fr_type is None:
        template = _get_template_filled_region_type(doc)
        if template is None:
            fr_debug(
                u'FR: No FilledRegionType template found for "{0}".'
                .format(type_name)
            )
            return

        try:
            new_type_id = template.Duplicate(type_name)
            fr_type = doc.GetElement(new_type_id)
            fr_debug(
                u'FR: Duplicated template "{0}" to new type "{1}" (Id {2}).'
                .format(
                    get_element_name(template),
                    type_name,
                    _get_element_id_int(fr_type.Id)
                )
            )
        except Exception as ex:
            # Most common: "The name is already in use for this element type"
            fr_debug(
                u'FR: Duplicate failed for "{0}" from template "{1}": {2}'
                .format(type_name, get_element_name(template), ex)
            )
            # Try to resolve an already-existing type with this name
            try:
                all_types = DB.FilteredElementCollector(doc).OfClass(DB.ElementType)
            except Exception:
                all_types = []

            for et in all_types:
                try:
                    if get_element_name(et) == type_name and isinstance(et, DB.FilledRegionType):
                        fr_type = et
                        fr_debug(
                            u'FR: Resolved existing FilledRegionType "{0}" (Id {1}) '
                            u'after duplicate failure.'
                            .format(type_name, _get_element_id_int(et.Id))
                        )
                        break
                except Exception:
                    continue

    # --- 3. If we still don't have a type, bail out for this mix ---
    if fr_type is None:
        fr_debug(
            u'FR: Could not get FilledRegionType for "{0}". Giving up for this mix.'
            .format(type_name)
        )
        return

    # --- 4. Update the type's colour ---
    _set_filled_region_type_color(fr_type, color)

    # --- 5. Ensure there is a strip in the host view ---
    if not (isinstance(host_view, DB.ViewPlan) or
            isinstance(host_view, DB.ViewDrafting) or
            isinstance(host_view, DB.ViewSection) or
            isinstance(host_view, DB.ViewDetail)):
        fr_debug(
            u'FR: Host view "{0}" (Id {1}) is not detail-capable; skipping strip.'
            .format(_to_unicode(host_view.Name), _get_element_id_int(host_view.Id))
        )
        return

    # Check if a strip of this type already exists in this view
    try:
        col_fr = DB.FilteredElementCollector(doc, host_view.Id).OfClass(DB.FilledRegion)
    except Exception:
        col_fr = []
    for fr in col_fr:
        try:
            if _get_element_id_int(fr.GetTypeId()) == _get_element_id_int(fr_type.Id):
                fr_debug(
                    u'FR: Existing strip already present for type "{0}".'
                    .format(type_name)
                )
                return
        except Exception:
            continue

    # Create a simple rectangle at (0,0)
    try:
        width_ft = FILLED_REGION_WIDTH_MM / 304.8
        height_ft = FILLED_REGION_HEIGHT_MM / 304.8

        p0 = DB.XYZ(0.0, 0.0, 0.0)
        p1 = DB.XYZ(width_ft, 0.0, 0.0)
        p2 = DB.XYZ(width_ft, height_ft, 0.0)
        p3 = DB.XYZ(0.0, height_ft, 0.0)

        curves = List[DB.Curve]()
        curves.Add(DB.Line.CreateBound(p0, p1))
        curves.Add(DB.Line.CreateBound(p1, p2))
        curves.Add(DB.Line.CreateBound(p2, p3))
        curves.Add(DB.Line.CreateBound(p3, p0))

        loop = DB.CurveLoop.Create(curves)
        loops = List[DB.CurveLoop]()
        loops.Add(loop)

        fr = DB.FilledRegion.Create(doc, fr_type.Id, host_view.Id, loops)
        fr_debug(
            u'FR: Created strip Id {0} of type "{1}" in view "{2}".'
            .format(_get_element_id_int(fr.Id), type_name, _to_unicode(host_view.Name))
        )
    except Exception as ex:
        fr_debug(
            u'FR: Exception while creating strip for "{0}": {1}'
            .format(type_name, ex)
        )

# ---- Percent conversions ----
def pct_raw_to_display(raw):
    """Convert stored decimal or percent-ish value to '50%' style string."""
    s = _to_unicode(raw).strip()
    if not s:
        return u''
    s = s.replace(u'%', u'').replace(u',', u'.')
    cleaned = []
    for ch in s:
        if ch.isdigit() or ch in u'.-':
            cleaned.append(ch)
    s_num = u''.join(cleaned).strip()
    if not s_num:
        return u''
    try:
        val = float(s_num)
    except Exception:
        return u''
    if val <= 1.0 + 1e-9:
        perc = val * 100.0
    else:
        perc = val
    int_perc = int(round(perc))
    if abs(perc - int_perc) < 1e-6:
        num_str = _to_unicode(int_perc)
    else:
        num_str = (u'%.2f' % perc).rstrip('0').rstrip('.')
    return num_str + u'%'


def pct_display_to_raw(display):
    """Convert user display like '50%' or '50' or '0.5' into decimal string '0.5'."""
    s = _to_unicode(display).strip()
    if not s:
        return u''
    s = s.replace(u'%', u'').replace(u',', u'.')
    cleaned = []
    for ch in s:
        if ch.isdigit() or ch in u'.-':
            cleaned.append(ch)
    s_num = u''.join(cleaned).strip()
    if not s_num:
        return u''
    try:
        val = float(s_num)
    except Exception:
        return u''
    if val > 1.0 + 1e-9:
        raw = val / 100.0
    else:
        raw = val
    raw_str = (u'%.4f' % raw).rstrip('0').rstrip('.')
    if not raw_str:
        raw_str = u'0'
    return raw_str


# ---- Spacing conversions ----
def space_raw_to_display(raw):
    """Convert stored mm-like value string to 'Xm' display string."""
    s = _to_unicode(raw).strip()
    if not s:
        return u''
    cleaned = []
    for ch in s:
        if ch.isdigit() or ch in u'.,-':
            cleaned.append(ch)
    s_num = u''.join(cleaned).replace(u',', u'.').strip()
    if not s_num:
        return u''
    try:
        mm = float(s_num)
    except Exception:
        return u''
    m_val = mm / 1000.0
    int_m = int(round(m_val))
    if abs(m_val - int_m) < 1e-6:
        num_str = _to_unicode(int_m)
    else:
        num_str = (u'%.3f' % m_val).rstrip('0').rstrip('.')
    return num_str + u'm'


def space_display_to_raw(display):
    """Convert user display like '1.5m' or '1.5' into mm string '1500'."""
    s = _to_unicode(display).strip().lower()
    if not s:
        return u''
    s = s.replace(u'mm', u'')
    s = s.replace(u'm', u'')
    cleaned = []
    for ch in s:
        if ch.isdigit() or ch in u'.,-':
            cleaned.append(ch)
    s_num = u''.join(cleaned).replace(u',', u'.').strip()
    if not s_num:
        return u''
    try:
        m_val = float(s_num)
    except Exception:
        return u''
    mm_val = m_val * 1000.0
    mm_int = int(round(mm_val))
    return _to_unicode(mm_int)


# -----------------------------
# Data models
# -----------------------------
class SpeciesRow(object):
    def __init__(self, index, code, pct, spacing, bot, com, grade):
        self.index = index
        self.code = code or u''
        self.pct = pct or u''
        self.spacing = spacing or u''
        self.bot = bot or u''
        self.com = com or u''
        self.grade = grade or u''


class MixModel(object):
    """In-memory representation of a single mix schedule annotation instance."""
    def __init__(self, element):
        self.element = element
        self.element_id = element.Id

        # UI header references
        self.title_block = None
        self.header_grid = None
        self.summary_block = None

        self.mix_name = get_param(element, PARAM_MIX_NAME) or u'(unnamed mix)'

        try:
            ns = get_param(element, PARAM_NUM_SPECIES)
            self.num_species = int(ns) if ns not in (None, u'', '') else 0
        except Exception:
            self.num_species = 0

        if self.num_species < 0:
            self.num_species = 0
        if self.num_species > MAX_SPECIES:
            self.num_species = MAX_SPECIES

        self.rows = []
        self._load_rows()

        # UI references
        self.panel = None        # outer Border
        self.body_panel = None   # StackPanel with table content
        self.arrow = None        # TextBlock arrow icon
        self.is_expanded = False

        self.dirty = False

        # Color scheme / Area color linkage
        self.area_color_entry = None          # ColorFillSchemeEntry associated with this mix (if any)
        self.area_color_dbcolor = None        # Original Revit.DB.Color from scheme
        self.area_color_new_dbcolor = None    # Pending new color to push on Apply
        self.area_color_border = None         # WPF Border used for colour swatch UI

    def _load_rows(self):
        self.rows = []
        for i in range(1, self.num_species + 1):
            code_name  = PARAM_SPECIES_CODE_TEMPLATE.format(i)
            pct_name   = PARAM_SPECIES_PCT_TEMPLATE.format(i)
            space_name = PARAM_SPECIES_SPACE_TEMPLATE.format(i)
            bot_name   = PARAM_SPECIES_BOT_TEMPLATE.format(i)
            com_name   = PARAM_SPECIES_COM_TEMPLATE.format(i)
            grade_name = PARAM_SPECIES_GRADE_TEMPLATE.format(i)

            code_raw  = get_param(self.element, code_name)
            pct_raw   = get_param(self.element, pct_name)
            space_raw = get_param(self.element, space_name)
            bot       = get_param(self.element, bot_name)
            com       = get_param(self.element, com_name)
            grade     = get_param(self.element, grade_name)

            pct_display = pct_raw_to_display(pct_raw)
            space_display = space_raw_to_display(space_raw)

            self.rows.append(SpeciesRow(i, code_raw, pct_display, space_display, bot, com, grade))

    def remove_row_at(self, index):
        if index < 0 or index >= len(self.rows):
            return
        del self.rows[index]
        for idx, row in enumerate(self.rows):
            row.index = idx + 1
        self.num_species = len(self.rows)
        self.dirty = True

    def add_row(self):
        if len(self.rows) >= MAX_SPECIES:
            return
        new_index = len(self.rows) + 1
        self.rows.append(SpeciesRow(new_index, u'', u'', u'', u'', u'', u''))
        self.num_species = len(self.rows)
        self.dirty = True


# -----------------------------
# Ghost cleanup (not used heavily with this version, but kept)
# -----------------------------
_GHOST_CLEANUP_STATE = None


def _on_idling_cleanup(sender, args):
    """Run once on Idling to delete ghost Area + boundaries."""
    global _GHOST_CLEANUP_STATE
    state = _GHOST_CLEANUP_STATE
    if not state:
        try:
            sender.Idling -= _on_idling_cleanup
        except Exception:
            pass
        return

    doc = state.get('doc')
    if doc is None:
        _GHOST_CLEANUP_STATE = None
        try:
            sender.Idling -= _on_idling_cleanup
        except Exception:
            pass
        return

    t = DB.Transaction(doc, 'Cleanup Mix Editor ghost Areas')
    try:
        t.Start()

        area_id = state.get('area_id')
        if area_id:
            try:
                doc.Delete(area_id)
            except Exception:
                pass

        for bid in state.get('boundary_ids') or []:
            try:
                doc.Delete(bid)
            except Exception:
                pass

        sketch_plane_id = state.get('sketch_plane_id')
        if sketch_plane_id:
            try:
                doc.Delete(sketch_plane_id)
            except Exception:
                pass

        t.Commit()
    except Exception:
        try:
            t.RollBack()
        except Exception:
            pass

    _GHOST_CLEANUP_STATE = None
    try:
        sender.Idling -= _on_idling_cleanup
    except Exception:
        pass


def _schedule_ghost_cleanup(doc, area_id, boundary_ids, sketch_plane_id):
    """Store ghost IDs and hook Idling so cleanup runs once after close."""
    global _GHOST_CLEANUP_STATE
    if (area_id is None and (not boundary_ids) and sketch_plane_id is None):
        return

    _GHOST_CLEANUP_STATE = {
        'doc': doc,
        'area_id': area_id,
        'boundary_ids': list(boundary_ids or []),
        'sketch_plane_id': sketch_plane_id
    }

    try:
        uiapp = __revit__   # pyRevit gives this UIApplication global
        uiapp.Idling += _on_idling_cleanup
    except Exception:
        _GHOST_CLEANUP_STATE = None


# -----------------------------
# Window controller
# -----------------------------
class MixWindowController(object):
    def __init__(self, doc):
        self.doc = doc
        self.mixes = []
        self.window = None
        self.stack_panel = None
        self.apply_button = None
        self.close_button = None
        self.current_expanded = None
        self._mix_symbol = None
        self._target_view_id = None

        # For manual double-click detection on header
        self._last_header_mix = None
        self._last_header_click_time = None

        # Pending Area renames: list of (old_name, new_name)
        self._area_renames = []

        # Ghost Area + boundary for seeding names
        self._ghost_area_id = None
        self._ghost_boundary_ids = []
        self._ghost_sketch_plane_id = None

        # Colour scheme / swatch state
        self._color_scheme = None               # DB.ColorFillScheme or None
        self._color_entries_by_key = {}         # u"Mix Name" -> ColorFillSchemeEntry

        self._load_mixes()
        self._load_color_scheme()
        for mix in self.mixes:
            self._attach_color_to_mix(mix)
        self._build_window()

    # ---- Plant Library ----
    def _run_create_plant_script(self):
        """Run the Create Plant pushbutton script from Mix Editor."""
        if not os.path.exists(CREATE_PLANT_SCRIPT_PATH):
            forms.alert(
                u"Could not find Create Plant script at:\n{0}".format(CREATE_PLANT_SCRIPT_PATH),
                title="Add Plant (library)"
            )
            return

        try:
            # Load as a uniquely-named module so it doesn't clash with this script.py
            # Loading it will execute its top-level code, just like pressing its own button.
            imp.load_source("create_plant_pushbutton_script", CREATE_PLANT_SCRIPT_PATH)
        except Exception as ex:
            forms.alert(
                u"Error while running Create Plant script:\n{0}".format(ex),
                title="Add Plant (library)"
            )


    # ---- Revit data ----
    def _load_mixes(self):
        col = (DB.FilteredElementCollector(self.doc)
               .OfClass(DB.FamilyInstance)
               .OfCategory(DB.BuiltInCategory.OST_GenericAnnotation))

        view_counts = {}

        for fi in col:
            try:
                fam = fi.Symbol.Family
            except Exception:
                fam = None
            if fam and fam.Name == FAMILY_NAME:
                self.mixes.append(MixModel(fi))
                try:
                    view_id = fi.OwnerViewId
                except Exception:
                    view_id = None
                view_key = _get_element_id_int(view_id)
                if view_key not in (None, -1):
                    key = view_key
                    if key in view_counts:
                        view_counts[key] += 1
                    else:
                        view_counts[key] = 1

        self.mixes.sort(key=lambda m: (m.mix_name or u'').lower())

        self._target_view_id = None
        if view_counts:
            max_count = max(view_counts.values())
            candidate_keys = [k for k, v in view_counts.items() if v == max_count]
            chosen = None
            if len(candidate_keys) == 1:
                chosen = candidate_keys[0]
            else:
                for key in candidate_keys:
                    try:
                        vid = DB.ElementId(key)
                        v = self.doc.GetElement(vid)
                        if isinstance(v, DB.View) and v.Name == DEFAULT_MIX_DRAFTING_VIEW_NAME:
                            chosen = key
                            break
                    except Exception:
                        continue
                if chosen is None:
                    chosen = candidate_keys[0]
            try:
                self._target_view_id = DB.ElementId(chosen)
            except Exception:
                self._target_view_id = None

    def _load_color_scheme(self):
        """Find the Color Fill Scheme for the chosen category and name."""
        self._color_scheme = None
        self._color_entries_by_key = {}

        doc = self.doc

        try:
            schemes = list(DB.FilteredElementCollector(doc).OfClass(DB.ColorFillScheme))
        except Exception:
            schemes = []

        if not schemes:
            LOGGER.warning('No ColorFillScheme elements found in document.')
            return

        cat_id = None
        try:
            cat = doc.Settings.Categories.get_Item(COLOR_SCHEME_CATEGORY)
            if cat:
                cat_id = cat.Id
        except Exception:
            try:
                cat = doc.Settings.Categories.get_Item(str(COLOR_SCHEME_CATEGORY))
                if cat:
                    cat_id = cat.Id
            except Exception:
                pass

        if cat_id is not None:
            by_cat = []
            for scheme in schemes:
                try:
                    if (_get_element_id_int(scheme.CategoryId)
                            == _get_element_id_int(cat_id)):
                        by_cat.append(scheme)
                except Exception:
                    pass
            if by_cat:
                schemes = by_cat

        if not schemes:
            LOGGER.warning(
                'No ColorFillScheme found for category: {0}'.format(COLOR_SCHEME_CATEGORY)
            )
            return

        chosen = None
        if COLOR_SCHEME_NAME:
            for scheme in schemes:
                try:
                    if scheme.Name == COLOR_SCHEME_NAME:
                        chosen = scheme
                        break
                except Exception:
                    pass

        if chosen is None:
            chosen = schemes[0]

        self._color_scheme = chosen

        if not self._color_scheme:
            return

        lookup = {}

        entries = []
        try:
            entries = list(self._color_scheme.GetEntries())
        except Exception:
            try:
                entries = list(self._color_scheme.Entries)
            except Exception:
                entries = []

        for entry in entries:
            v_key, c_key = _get_color_entry_keys(entry)
            for key in (v_key, c_key):
                if key and key not in lookup:
                    lookup[key] = entry

        self._color_entries_by_key = lookup

        LOGGER.info(
            'Loaded color scheme "{0}" for category {1}. {2} entries mapped.'
            .format(self._color_scheme.Name, COLOR_SCHEME_CATEGORY, len(self._color_entries_by_key))
        )

    def _attach_color_to_mix(self, mix):
        """Link a mix to a ColorFillSchemeEntry (if any) based on its name."""
        scheme = self._color_scheme
        mapping = self._color_entries_by_key
        if scheme is None or not mapping or mix is None:
            return

        key = _to_unicode(mix.mix_name).strip()
        if not key:
            return

        entry = mapping.get(key)

        if entry is None:
            lk = key.lower()
            for k, v in mapping.items():
                try:
                    if _to_unicode(k).strip().lower() == lk:
                        entry = v
                        break
                except Exception:
                    continue

        mix.area_color_entry = entry
        mix.area_color_new_dbcolor = None

        if entry is not None:
            mix.area_color_dbcolor = _get_entry_color(entry)
        else:
            mix.area_color_dbcolor = None

    def _get_mix_symbol(self):
        """Find and cache the FamilySymbol for the bm_massed_mix generic annotation."""
        if getattr(self, '_mix_symbol', None) is not None:
            return self._mix_symbol
        col = (DB.FilteredElementCollector(self.doc)
               .OfClass(DB.FamilySymbol)
               .OfCategory(DB.BuiltInCategory.OST_GenericAnnotation))
        for sym in col:
            try:
                fam = sym.Family
            except Exception:
                fam = None
            if fam and fam.Name == FAMILY_NAME:
                self._mix_symbol = sym
                return sym
        self._mix_symbol = None
        return None

    def _get_target_view(self):
        """Determine which view to place new mix instances + filled regions into."""
        try:
            if self._target_view_id is not None:
                v = self.doc.GetElement(self._target_view_id)
                if isinstance(v, DB.View):
                    return v
        except Exception:
            pass

        if self.mixes:
            try:
                inst = self.mixes[0].element
                owner = self.doc.GetElement(inst.OwnerViewId)
                if isinstance(owner, DB.View):
                    self._target_view_id = owner.Id
                    return owner
            except Exception:
                pass

        try:
            views = (DB.FilteredElementCollector(self.doc)
                     .OfClass(DB.ViewDrafting))
            for v in views:
                if v.Name == DEFAULT_MIX_DRAFTING_VIEW_NAME:
                    self._target_view_id = v.Id
                    return v
        except Exception:
            pass

        try:
            return self.doc.ActiveView
        except Exception:
            return None

    def _record_mix_rename(self, old_name, new_name):
        """Record that a mix name changed; will be applied to Areas on Apply."""
        old_name = _to_unicode(old_name).strip()
        new_name = _to_unicode(new_name).strip()
        if not old_name or not new_name or old_name == new_name:
            return
        self._area_renames.append((old_name, new_name))

    def _generate_copy_name(self, base_name):
        """Generate a unique 'base_name Copy' or 'base_name Copy X' name."""
        base_name_u = _to_unicode(base_name).strip()
        if not base_name_u:
            base_name_u = u'(unnamed mix)'
        existing = set((_to_unicode(m.mix_name).strip() for m in self.mixes))
        candidate = base_name_u + u' Copy'
        if candidate not in existing:
            return candidate
        i = 2
        while True:
            candidate = u'{0} Copy {1}'.format(base_name_u, i)
            if candidate not in existing:
                return candidate
            i += 1

    # ---- UI construction ----
    def _build_window(self):
        if not os.path.exists(XAML_FILE):
            forms.alert('XAML file not found:\n{0}'.format(XAML_FILE))
            return

        stream = FileStream(XAML_FILE, FileMode.Open)
        try:
            self.window = XamlReader.Load(stream)
        finally:
            stream.Close()

        # Fix width so long text cannot force the window to grow;
        # content will clip within the table instead.
        try:
            self.window.SizeToContent = SizeToContent.Height
        except Exception:
            pass

        self.stack_panel = self.window.FindName('MixStackPanel')
        self.apply_button = self.window.FindName('ApplyButton')
        self.close_button = self.window.FindName('CloseButton')

        if self.apply_button is not None:
            self.apply_button.Click += self.on_apply
        if self.close_button is not None:
            self.close_button.Click += self.on_close

        if self.stack_panel is not None:
            self._refresh_stack_panel()

    def _refresh_stack_panel(self):
        if self.stack_panel is None:
            return
        self.stack_panel.Children.Clear()
        self.current_expanded = None
        for mix in self.mixes:
            panel = self._create_mix_panel(mix)
            mix.panel = panel
            self.stack_panel.Children.Add(panel)
        create_entry = self._create_new_mix_button()
        self.stack_panel.Children.Add(create_entry)

    def _create_new_mix_button(self):
        border = Border()
        border.Margin = Thickness(0, 0, 0, 6)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(6)
        border.BorderBrush = Media.Brushes.LightGray
        border.Background = Media.Brushes.WhiteSmoke
        border.Cursor = Cursors.Hand

        inner_stack = StackPanel()
        inner_stack.Orientation = Orientation.Horizontal
        inner_stack.Margin = Thickness(6)
        border.Child = inner_stack

        plus = TextBlock()
        plus.Text = u'+'
        plus.FontSize = 14
        plus.Foreground = Media.Brushes.ForestGreen
        plus.Margin = Thickness(0, 0, 4, 0)
        inner_stack.Children.Add(plus)

        label = TextBlock()
        label.Text = u'Create New Mix'
        label.FontSize = 13
        label.Foreground = Media.Brushes.ForestGreen
        inner_stack.Children.Add(label)

        border.MouseLeftButtonUp += self.on_create_new_mix

        return border

    def _create_mix_panel(self, mix):
        border = Border()
        border.Margin = Thickness(0, 0, 0, 6)
        border.BorderThickness = Thickness(1)
        border.CornerRadius = CornerRadius(6)
        border.BorderBrush = Media.Brushes.LightGray
        border.Background = Media.Brushes.White

        outer_stack = StackPanel()
        border.Child = outer_stack

        header_grid = Grid()
        header_grid.Margin = Thickness(6)
        header_grid.Cursor = Cursors.Hand

        header_grid.ColumnDefinitions.Add(ColumnDefinition())
        header_grid.ColumnDefinitions.Add(ColumnDefinition())
        header_grid.ColumnDefinitions.Add(ColumnDefinition())

        header_grid.ColumnDefinitions[0].Width = GridLength(20)
        header_grid.ColumnDefinitions[1].Width = GridLength(1, GridUnitType.Star)
        header_grid.ColumnDefinitions[2].Width = GridLength(24)

        arrow = TextBlock()
        arrow.Text = u'▸'
        arrow.FontSize = 14
        arrow.VerticalAlignment = VerticalAlignment.Center
        arrow.Foreground = Media.Brushes.DimGray
        Grid.SetColumn(arrow, 0)
        header_grid.Children.Add(arrow)

        header_stack = StackPanel()
        header_stack.Orientation = Orientation.Vertical
        Grid.SetColumn(header_stack, 1)
        header_grid.Children.Add(header_stack)

        title = TextBlock()
        title.Text = mix.mix_name
        title.FontSize = 14
        title.VerticalAlignment = VerticalAlignment.Center
        title.Foreground = Media.Brushes.Black
        header_stack.Children.Add(title)

        summary = TextBlock()
        summary.Text = u''
        summary.FontSize = 10
        summary.Margin = Thickness(0, 1, 0, 0)
        summary.Foreground = Media.Brushes.DimGray
        summary.Visibility = Visibility.Collapsed
        header_stack.Children.Add(summary)

        dup_border = Border()
        dup_border.HorizontalAlignment = HorizontalAlignment.Right
        dup_border.VerticalAlignment = VerticalAlignment.Center
        dup_border.Padding = Thickness(2)
        dup_border.Background = Media.Brushes.Transparent
        Grid.SetColumn(dup_border, 2)

        dup_img = Image()
        dup_img.Width = 16
        dup_img.Height = 16
        dup_img.Stretch = Media.Stretch.Uniform
        dup_img.Opacity = 0.5

        try:
            if os.path.exists(DUPLICATE_ICON_PATH):
                bmp = BitmapImage()
                bmp.BeginInit()
                bmp.UriSource = Uri(DUPLICATE_ICON_PATH)
                bmp.EndInit()
                dup_img.Source = bmp
        except Exception:
            pass

        dup_border.Child = dup_img
        dup_border.Tag = mix
        dup_border.MouseEnter += self.on_duplicate_icon_mouse_enter
        dup_border.MouseLeave += self.on_duplicate_icon_mouse_leave
        dup_border.MouseLeftButtonUp += self.on_duplicate_mix

        header_grid.Children.Add(dup_border)

        body = StackPanel()
        body.Margin = Thickness(24, 0, 6, 6)
        body.Visibility = Visibility.Collapsed

        mix.body_panel = body
        mix.arrow = arrow
        mix.header_grid = header_grid
        mix.title_block = title
        mix.summary_block = summary

        header_grid.Tag = mix
        header_grid.MouseLeftButtonUp += self.on_header_mouse_left_button_up

        outer_stack.Children.Add(header_grid)
        outer_stack.Children.Add(body)

        self._render_mix_body(mix)

        if mix.is_expanded:
            body.Visibility = Visibility.Visible
            arrow.Text = u'▾'

        return border

    def _update_mix_percent_summary(self, mix):
        """Update the header percent summary for a mix."""
        sb = getattr(mix, 'summary_block', None)
        if sb is None:
            return

        total = 0.0
        for row in mix.rows:
            raw = pct_display_to_raw(row.pct)
            if not raw:
                continue
            try:
                val = float(raw) * 100.0
            except Exception:
                continue
            total += val

        diff = total - 100.0
        if abs(diff) < 0.5:
            sb.Text = u''
            sb.Visibility = Visibility.Collapsed
            return

        abs_diff = abs(diff)
        int_diff = int(round(abs_diff))
        if abs(abs_diff - int_diff) < 1e-6:
            diff_str = u'{0}%'.format(int_diff)
        else:
            diff_str = (u'%.1f%%' % abs_diff)

        if diff > 0:
            sb.Text = diff_str + u' surplus'
            sb.Foreground = Media.Brushes.ForestGreen
        else:
            sb.Text = diff_str + u' deficit'
            sb.Foreground = Media.Brushes.IndianRed

        sb.Visibility = Visibility.Visible

    def _mix_has_area(self, mix):
        """Return True if there is an Area whose Name matches this mix name."""
        if mix is None:
            return False

        name = getattr(mix, 'mix_name', None)
        if not name:
            return False

        name = _to_unicode(name).strip()
        if not name:
            return False

        try:
            areas = (DB.FilteredElementCollector(self.doc)
                     .OfCategory(DB.BuiltInCategory.OST_Areas)
                     .WhereElementIsNotElementType())
        except Exception:
            return False

        for area in areas:
            nm = get_param(area, AREA_NAME_PARAM)
            if not nm:
                continue
            if _to_unicode(nm).strip() == name:
                return True

        return False


    def _render_mix_body(self, mix):
        body = mix.body_panel
        if body is None:
            return
        body.Children.Clear()

        # --- Area colour row ---
        color_grid = Grid()
        color_grid.Margin = Thickness(0, 0, 0, 4)
        color_grid.ColumnDefinitions.Add(ColumnDefinition())
        color_grid.ColumnDefinitions.Add(ColumnDefinition())
        color_grid.ColumnDefinitions.Add(ColumnDefinition())
        color_grid.ColumnDefinitions[0].Width = GridLength(1, GridUnitType.Auto)
        color_grid.ColumnDefinitions[1].Width = GridLength(80)
        color_grid.ColumnDefinitions[2].Width = GridLength(80)


        color_label = TextBlock()
        color_label.Text = u'Area colour'
        color_label.FontSize = 11
        color_label.Margin = Thickness(0, 0, 8, 0)
        color_label.Foreground = Media.Brushes.DimGray
        color_label.VerticalAlignment = VerticalAlignment.Center
        Grid.SetColumn(color_label, 0)
        color_grid.Children.Add(color_label)

        swatch = Border()
        swatch.Width = 60
        swatch.Height = 16
        swatch.CornerRadius = CornerRadius(3)
        swatch.BorderThickness = Thickness(1)
        swatch.BorderBrush = Media.Brushes.Gray
        swatch.Margin = Thickness(0, 0, 0, 0)
        swatch.HorizontalAlignment = HorizontalAlignment.Left

        if mix.area_color_new_dbcolor is not None:
            brush = _dbcolor_to_media_brush(mix.area_color_new_dbcolor)
        else:
            brush = _dbcolor_to_media_brush(mix.area_color_dbcolor)
        swatch.Background = brush

        # If no Area exists for this mix, show "Not Placed" in the swatch
        if not self._mix_has_area(mix):
            status_tb = TextBlock()
            status_tb.Text = u'Not Placed'
            status_tb.FontSize = 10
            # darker grey than the main label
            status_tb.Foreground = Media.Brushes.DarkSlateGray
            status_tb.HorizontalAlignment = HorizontalAlignment.Center
            status_tb.VerticalAlignment = VerticalAlignment.Center
            status_tb.Margin = Thickness(0)
            swatch.Child = status_tb


        if mix.area_color_entry is not None and self._color_scheme is not None:
            swatch.Cursor = Cursors.Hand
            swatch.Opacity = 1.0
            swatch.Tag = mix
            swatch.MouseLeftButtonUp += self.on_area_color_click
        else:
            swatch.Cursor = Cursors.Arrow
            swatch.Opacity = 0.4

        Grid.SetColumn(swatch, 1)
        color_grid.Children.Add(swatch)

        # --- Place Area button next to colour swatch ---
        place_btn = Button()
        place_btn.Content = u'Place Area'
        place_btn.FontSize = 10
        place_btn.Margin = Thickness(6, 0, 0, 0)
        place_btn.Padding = Thickness(4, 0, 4, 0)
        place_btn.HorizontalAlignment = HorizontalAlignment.Left
        place_btn.Tag = mix
        place_btn.Click += self.on_place_area_click
        Grid.SetColumn(place_btn, 2)
        color_grid.Children.Add(place_btn)

        mix.area_color_border = swatch
        body.Children.Add(color_grid)


        # --- Header row for species table ---
        header_grid = Grid()
        for _ in range(7):  # Code, Percent, Spacing, Bot, Com, Grade, minus
            header_grid.ColumnDefinitions.Add(ColumnDefinition())

        header_grid.ColumnDefinitions[0].Width = GridLength(60)   # Code
        header_grid.ColumnDefinitions[1].Width = GridLength(45)   # Percent
        header_grid.ColumnDefinitions[2].Width = GridLength(45)   # Spacing
        header_grid.ColumnDefinitions[3].Width = GridLength(1, GridUnitType.Star)  # Bot
        header_grid.ColumnDefinitions[4].Width = GridLength(1, GridUnitType.Star)  # Com
        header_grid.ColumnDefinitions[5].Width = GridLength(45)   # Grade
        header_grid.ColumnDefinitions[6].Width = GridLength(26)   # Minus

        labels = [u'Code', u'Percent', u'Spacing',
                  u'Botanical Name', u'Common Name', u'Grade', u'']
        for idx, text in enumerate(labels):
            tb = TextBlock()
            tb.Text = text
            tb.FontSize = 11
            tb.Margin = Thickness(0, 0, 4, 2)
            tb.Foreground = Media.Brushes.DimGray
            tb.FontWeight = FontWeights.SemiBold
            Grid.SetColumn(tb, idx)
            header_grid.Children.Add(tb)
        body.Children.Add(header_grid)

        # --- Data rows ---
        for row_index, row in enumerate(mix.rows):
            row_grid = Grid()
            for _ in range(7):
                row_grid.ColumnDefinitions.Add(ColumnDefinition())

            row_grid.ColumnDefinitions[0].Width = GridLength(60)
            row_grid.ColumnDefinitions[1].Width = GridLength(45)
            row_grid.ColumnDefinitions[2].Width = GridLength(45)
            row_grid.ColumnDefinitions[3].Width = GridLength(1, GridUnitType.Star)
            row_grid.ColumnDefinitions[4].Width = GridLength(1, GridUnitType.Star)
            row_grid.ColumnDefinitions[5].Width = GridLength(45)
            row_grid.ColumnDefinitions[6].Width = GridLength(26)

            code_box = TextBox()
            code_box.Text = row.code
            code_box.Margin = Thickness(0, 0, 4, 2)
            code_box.Tag = (mix, row_index, 'code')
            code_box.TextChanged += self.on_cell_changed
            Grid.SetColumn(code_box, 0)
            row_grid.Children.Add(code_box)

            pct_box = TextBox()
            pct_box.Text = row.pct
            pct_box.Margin = Thickness(0, 0, 4, 2)
            pct_box.Tag = (mix, row_index, 'pct')
            pct_box.TextChanged += self.on_cell_changed
            pct_box.LostFocus += self.on_pct_lost_focus
            pct_box.GotKeyboardFocus += self.on_textbox_got_keyboard_focus
            pct_box.PreviewMouseLeftButtonDown += self.on_textbox_preview_mouse_left_button_down
            Grid.SetColumn(pct_box, 1)
            row_grid.Children.Add(pct_box)

            space_box = TextBox()
            space_box.Text = row.spacing
            space_box.Margin = Thickness(0, 0, 4, 2)
            space_box.Tag = (mix, row_index, 'spacing')
            space_box.TextChanged += self.on_cell_changed
            space_box.LostFocus += self.on_space_lost_focus
            space_box.GotKeyboardFocus += self.on_textbox_got_keyboard_focus
            space_box.PreviewMouseLeftButtonDown += self.on_textbox_preview_mouse_left_button_down
            Grid.SetColumn(space_box, 2)
            row_grid.Children.Add(space_box)

            # Botanical name cell: outer border for grid line, inner grid for text + fade
            bot_border = Border()
            bot_border.Margin = Thickness(0, 0, 4, 2)
            bot_border.BorderThickness = Thickness(1)
            bot_border.BorderBrush = Media.Brushes.LightGray
            bot_border.Padding = Thickness(0)
            bot_border.HorizontalAlignment = HorizontalAlignment.Stretch

            bot_inner = Grid()
            bot_border.Child = bot_inner

            bot_box = TextBox()
            bot_box.Text = row.bot
            bot_box.Margin = Thickness(0)
            bot_box.BorderThickness = Thickness(0)   # border comes from outer Border
            bot_box.Background = Media.Brushes.Transparent
            bot_box.Tag = (mix, row_index, 'bot')
            bot_box.TextChanged += self.on_cell_changed

            bot_box.TextWrapping = TextWrapping.NoWrap
            bot_box.SetValue(
                ScrollViewer.HorizontalScrollBarVisibilityProperty,
                ScrollBarVisibility.Hidden
            )
            bot_box.VerticalContentAlignment = VerticalAlignment.Center
            bot_box.HorizontalAlignment = HorizontalAlignment.Stretch

            # Tooltip when text is long enough that it might be clipped
            if row.bot and len(row.bot) >= NAME_TOOLTIP_MIN_CHARS:
                bot_box.ToolTip = row.bot
            else:
                bot_box.ToolTip = None

            bot_inner.Children.Add(bot_box)

            # Fade overlay: inside the cell border, drawn on top of the text
            if (row.bot and len(row.bot) >= NAME_TOOLTIP_MIN_CHARS
                    and NAME_FADE_WIDTH > 0 and NAME_FADE_FRACTION > 0.0):
                try:
                    fade = Border()
                    fade.Width = NAME_FADE_WIDTH
                    fade.HorizontalAlignment = HorizontalAlignment.Right
                    # small inset so we don't cover the border line
                    fade.Margin = Thickness(0, 1, 1, 1)
                    fade.BorderThickness = Thickness(0)
                    fade.IsHitTestVisible = False

                    fade_brush = Media.LinearGradientBrush()
                    fade_brush.StartPoint = Point(0, 0)
                    fade_brush.EndPoint = Point(1, 0)

                    fade_start = 1.0 - float(NAME_FADE_FRACTION)
                    if fade_start < 0.0:
                        fade_start = 0.0
                    if fade_start > 1.0:
                        fade_start = 1.0

                    # Transparent on the left, light veil on the right
                    fade_brush.GradientStops.Add(
                        Media.GradientStop(
                            Media.Color.FromArgb(0, 255, 255, 255), 0.0
                        )
                    )
                    fade_brush.GradientStops.Add(
                        Media.GradientStop(
                            Media.Color.FromArgb(0, 255, 255, 255), fade_start
                        )
                    )
                    fade_brush.GradientStops.Add(
                        Media.GradientStop(
                            Media.Color.FromArgb(200, 245, 245, 245), 1.0
                        )
                    )
                    fade.Background = fade_brush

                    # Add AFTER the TextBox so the visual order is:
                    # border (outer) -> text -> fade -> border stroke again
                    bot_inner.Children.Add(fade)
                except Exception:
                    pass

            Grid.SetColumn(bot_border, 3)
            row_grid.Children.Add(bot_border)

            # Common name cell: same pattern as Botanical
            com_border = Border()
            com_border.Margin = Thickness(0, 0, 4, 2)
            com_border.BorderThickness = Thickness(1)
            com_border.BorderBrush = Media.Brushes.LightGray
            com_border.Padding = Thickness(0)
            com_border.HorizontalAlignment = HorizontalAlignment.Stretch

            com_inner = Grid()
            com_border.Child = com_inner

            com_box = TextBox()
            com_box.Text = row.com
            com_box.Margin = Thickness(0)
            com_box.BorderThickness = Thickness(0)
            com_box.Background = Media.Brushes.Transparent
            com_box.Tag = (mix, row_index, 'com')
            com_box.TextChanged += self.on_cell_changed

            com_box.TextWrapping = TextWrapping.NoWrap
            com_box.SetValue(
                ScrollViewer.HorizontalScrollBarVisibilityProperty,
                ScrollBarVisibility.Hidden
            )
            com_box.VerticalContentAlignment = VerticalAlignment.Center
            com_box.HorizontalAlignment = HorizontalAlignment.Stretch

            if row.com and len(row.com) >= NAME_TOOLTIP_MIN_CHARS:
                com_box.ToolTip = row.com
            else:
                com_box.ToolTip = None

            com_inner.Children.Add(com_box)

            if (row.com and len(row.com) >= NAME_TOOLTIP_MIN_CHARS
                    and NAME_FADE_WIDTH > 0 and NAME_FADE_FRACTION > 0.0):
                try:
                    fade2 = Border()
                    fade2.Width = NAME_FADE_WIDTH
                    fade2.HorizontalAlignment = HorizontalAlignment.Right
                    fade2.Margin = Thickness(0, 1, 1, 1)
                    fade2.BorderThickness = Thickness(0)
                    fade2.IsHitTestVisible = False

                    fade_brush2 = Media.LinearGradientBrush()
                    fade_brush2.StartPoint = Point(0, 0)
                    fade_brush2.EndPoint = Point(1, 0)

                    fade_start2 = 1.0 - float(NAME_FADE_FRACTION)
                    if fade_start2 < 0.0:
                        fade_start2 = 0.0
                    if fade_start2 > 1.0:
                        fade_start2 = 1.0

                    fade_brush2.GradientStops.Add(
                        Media.GradientStop(
                            Media.Color.FromArgb(0, 255, 255, 255), 0.0
                        )
                    )
                    fade_brush2.GradientStops.Add(
                        Media.GradientStop(
                            Media.Color.FromArgb(0, 255, 255, 255), fade_start2
                        )
                    )
                    fade_brush2.GradientStops.Add(
                        Media.GradientStop(
                            Media.Color.FromArgb(200, 245, 245, 245), 1.0
                        )
                    )
                    fade2.Background = fade_brush2

                    com_inner.Children.Add(fade2)
                except Exception:
                    pass

            Grid.SetColumn(com_border, 4)
            row_grid.Children.Add(com_border)

            grade_box = TextBox()
            grade_box.Text = row.grade
            grade_box.Margin = Thickness(0, 0, 4, 2)
            grade_box.Tag = (mix, row_index, 'grade')
            grade_box.TextChanged += self.on_cell_changed
            Grid.SetColumn(grade_box, 5)
            row_grid.Children.Add(grade_box)

            minus_btn = Button()
            minus_btn.Content = u'−'
            minus_btn.Width = 20
            minus_btn.Height = 20
            minus_btn.Margin = Thickness(4, 0, 0, 2)
            minus_btn.Foreground = Media.Brushes.White
            minus_btn.Background = Media.Brushes.IndianRed
            minus_btn.BorderBrush = Media.Brushes.Transparent
            minus_btn.Tag = (mix, row_index)
            minus_btn.Click += self.on_remove_row
            Grid.SetColumn(minus_btn, 6)
            row_grid.Children.Add(minus_btn)

            body.Children.Add(row_grid)

        if len(mix.rows) < MAX_SPECIES:
            # Container grid so both buttons split the full width
            add_row_panel = Grid()
            add_row_panel.Margin = Thickness(0, 2, 0, 0)

            col1 = ColumnDefinition()
            col1.Width = GridLength(1, GridUnitType.Star)
            col2 = ColumnDefinition()
            col2.Width = GridLength(1, GridUnitType.Star)  # 50/50 split

            add_row_panel.ColumnDefinitions.Add(col1)
            add_row_panel.ColumnDefinitions.Add(col2)

            # --- Manual add button ---
            manual_border = Border()
            manual_border.Height = 20
            manual_border.Margin = Thickness(0, 0, 3, 0)   # small gap to the right
            manual_border.Padding = Thickness(4, 2, 4, 2)
            manual_border.Background = Media.Brushes.WhiteSmoke
            manual_border.BorderThickness = Thickness(1)
            manual_border.BorderBrush = Media.Brushes.LightGray
            manual_border.CornerRadius = CornerRadius(4)
            manual_border.Cursor = Cursors.Hand
            manual_border.HorizontalAlignment = HorizontalAlignment.Stretch

            manual_stack = StackPanel()
            manual_stack.Orientation = Orientation.Horizontal
            manual_border.Child = manual_stack

            plus_manual = TextBlock()
            plus_manual.Text = u'+'
            plus_manual.FontSize = 14
            plus_manual.Foreground = Media.Brushes.ForestGreen
            plus_manual.Margin = Thickness(0, -2, 2, 0)
            manual_stack.Children.Add(plus_manual)

            label_manual = TextBlock()
            label_manual.Text = u'Add Plant (manual)'
            label_manual.FontSize = 12
            label_manual.Foreground = Media.Brushes.ForestGreen
            manual_stack.Children.Add(label_manual)

            manual_border.Tag = mix
            manual_border.MouseLeftButtonUp += self.on_add_row

            Grid.SetColumn(manual_border, 0)
            add_row_panel.Children.Add(manual_border)

            # --- Plant library add button ---
            lib_border = Border()
            lib_border.Height = 20
            lib_border.Margin = Thickness(3, 0, 0, 0)      # small gap to the left
            lib_border.Padding = Thickness(4, 2, 4, 2)
            lib_border.Background = Media.Brushes.WhiteSmoke
            lib_border.BorderThickness = Thickness(1)
            lib_border.BorderBrush = Media.Brushes.LightGray
            lib_border.CornerRadius = CornerRadius(4)
            lib_border.Cursor = Cursors.Hand
            lib_border.HorizontalAlignment = HorizontalAlignment.Stretch

            lib_stack = StackPanel()
            lib_stack.Orientation = Orientation.Horizontal
            lib_border.Child = lib_stack

            plus_lib = TextBlock()
            plus_lib.Text = u'+'
            plus_lib.FontSize = 14
            plus_lib.Foreground = Media.Brushes.SteelBlue
            plus_lib.Margin = Thickness(0, -2, 2, 0)
            lib_stack.Children.Add(plus_lib)

            label_lib = TextBlock()
            label_lib.Text = u'Add Plant (library)'
            label_lib.FontSize = 12
            label_lib.Foreground = Media.Brushes.SteelBlue
            lib_stack.Children.Add(label_lib)

            lib_border.Tag = mix
            lib_border.MouseLeftButtonUp += self.on_add_row_from_library

            Grid.SetColumn(lib_border, 1)
            add_row_panel.Children.Add(lib_border)

            body.Children.Add(add_row_panel)


        self._update_mix_percent_summary(mix)

    def _begin_rename_mix(self, mix):
        """Inline rename of a mix header in the same position as the title text."""
        tb = getattr(mix, 'title_block', None)
        grid = getattr(mix, 'header_grid', None)
        if tb is None or grid is None:
            return

        try:
            header_stack = tb.Parent
        except Exception:
            header_stack = None
        if header_stack is None:
            return

        for child in list(header_stack.Children):
            try:
                if isinstance(child, TextBox) and getattr(child, 'Tag', None) is mix:
                    child.Focus()
                    child.SelectAll()
                    return
            except Exception:
                pass

        try:
            idx = header_stack.Children.IndexOf(tb)
        except Exception:
            idx = -1

        try:
            header_stack.Children.Remove(tb)
        except Exception:
            pass

        edit_box = TextBox()
        edit_box.Text = mix.mix_name or u''
        edit_box.Margin = tb.Margin
        edit_box.VerticalAlignment = tb.VerticalAlignment
        edit_box.FontSize = tb.FontSize
        edit_box.Foreground = tb.Foreground
        edit_box.Tag = mix

        if idx >= 0:
            header_stack.Children.Insert(idx, edit_box)
        else:
            header_stack.Children.Add(edit_box)

        def finish_edit(sender, commit):
            try:
                text_val = sender.Text or u''
            except Exception:
                text_val = u''

            try:
                header_stack.Children.Remove(sender)
            except Exception:
                pass
            try:
                if idx >= 0:
                    header_stack.Children.Insert(idx, tb)
                else:
                    header_stack.Children.Add(tb)
            except Exception:
                pass

            if commit:
                old_name = mix.mix_name
                new_name = text_val or u'(unnamed mix)'

                if old_name != new_name:
                    self._record_mix_rename(old_name, new_name)

                mix.mix_name = new_name
                tb.Text = mix.mix_name
                mix.dirty = True

                was_expanded = getattr(mix, 'is_expanded', False)

                try:
                    self.mixes.sort(key=lambda m: (m.mix_name or u'').lower())
                    mix.is_expanded = was_expanded
                    self._refresh_stack_panel()
                    if was_expanded:
                        self.current_expanded = mix
                except Exception:
                    pass

        def on_key(sender, args):
            try:
                keyname = str(args.Key)
            except Exception:
                keyname = ''
            if keyname in ('Return', 'Enter'):
                finish_edit(sender, True)
            elif keyname == 'Escape':
                finish_edit(sender, False)

        def on_lost_focus(sender, args):
            finish_edit(sender, True)

        edit_box.KeyDown += on_key
        edit_box.LostFocus += on_lost_focus
        try:
            edit_box.Focus()
            edit_box.SelectAll()
        except Exception:
            pass

    # ---- Event handlers ----
    def on_header_mouse_left_button_up(self, sender, args):
        """Single-click expands/collapses; quick double-click enters inline rename mode."""
        mix = sender.Tag
        if mix is None:
            return

        now = DateTime.Now
        double_click = False
        if self._last_header_mix is mix and self._last_header_click_time is not None:
            try:
                delta = now - self._last_header_click_time
                if delta.TotalMilliseconds <= 400:
                    double_click = True
            except Exception:
                double_click = False

        self._last_header_mix = mix
        self._last_header_click_time = now

        if double_click:
            self._begin_rename_mix(mix)
            try:
                args.Handled = True
            except Exception:
                pass
            return

        if mix.is_expanded:
            mix.is_expanded = False
            if mix.body_panel is not None:
                mix.body_panel.Visibility = Visibility.Collapsed
            if mix.arrow is not None:
                mix.arrow.Text = u'▸'
            if self.current_expanded is mix:
                self.current_expanded = None
        else:
            if self.current_expanded is not None and self.current_expanded is not mix:
                other = self.current_expanded
                other.is_expanded = False
                if other.body_panel is not None:
                    other.body_panel.Visibility = Visibility.Collapsed
                if other.arrow is not None:
                    other.arrow.Text = u'▸'

            mix.is_expanded = True
            if mix.body_panel is not None:
                mix.body_panel.Visibility = Visibility.Visible
            if mix.arrow is not None:
                mix.arrow.Text = u'▾'
            self.current_expanded = mix

    def on_duplicate_icon_mouse_enter(self, sender, args):
        try:
            img = sender.Child
            if img is not None:
                img.Opacity = 1.0
        except Exception:
            pass

    def on_duplicate_icon_mouse_leave(self, sender, args):
        try:
            img = sender.Child
            if img is not None:
                img.Opacity = 0.5
        except Exception:
            pass

    def on_area_color_click(self, sender, args):
        """Open palette + colour picker to edit the Area colour for this mix."""
        mix = getattr(sender, 'Tag', None)
        if mix is None:
            return

        entry = getattr(mix, 'area_color_entry', None)
        if entry is None or self._color_scheme is None:
            forms.alert(
                'This mix is not linked to a "{0}" color scheme entry.\n'
                'Ensure the Area Color Fill Scheme has an entry whose value or caption matches the mix name.'
                .format(COLOR_SCHEME_NAME)
            )
            return

        curr_dbcol = getattr(mix, 'area_color_new_dbcolor', None)
        if curr_dbcol is None:
            curr_dbcol = getattr(mix, 'area_color_dbcolor', None)

        new_dbcol = pick_area_color_with_palette(curr_dbcol)
        if new_dbcol is None:
            return

        mix.area_color_new_dbcolor = new_dbcol
        mix.dirty = True

        try:
            sender.Background = _dbcolor_to_media_brush(new_dbcol)
            sender.Opacity = 1.0
        except Exception:
            pass

    def on_place_area_click(self, sender, args):
        """Run Boundary.py when 'Place Area' is clicked for a mix."""
        mix = getattr(sender, 'Tag', None)
        if mix is None:
            return

        doc = self.doc

        # Optional: get mix name for future use
        mix_name = getattr(mix, 'mix_name', None)
        if mix_name is not None:
            mix_name = _to_unicode(mix_name).strip()

        # Close current window before launching external tool
        try:
            if self.window is not None:
                self.window.Close()
        except Exception:
            pass

        try:
            # Ensure our script directory is on sys.path
            if SCRIPT_DIR not in sys.path:
                sys.path.append(SCRIPT_DIR)

            try:
                import Boundary as bs
            except Exception:
                bs = None

            if bs is None:
                forms.alert(
                    'Could not find "Boundary.py" next to this tool.\n'
                    'Make sure it is in the same folder as script.py.',
                    title='Place Area'
                )
            else:
                # Reload so edits to Boundary.py are picked up
                try:
                    reload(bs)
                except Exception:
                    pass

                try:
                    # Prefer Boundary.main(mix_name) if it accepts an argument,
                    # otherwise fall back to main().
                    if mix_name:
                        try:
                            bs.main(mix_name)
                        except TypeError:
                            # Older Boundary without parameter support
                            bs.main()
                    else:
                        bs.main()
                except Exception as ex:
                    forms.alert(
                        u'Boundary.py raised an error:\n{0}'.format(ex),
                        title='Place Area'
                    )
        finally:
            # Re-open the Mix Editor so the user can keep working
            controller = MixWindowController(doc)
            if controller.window is not None:
                controller.show()



    def on_cell_changed(self, sender, args):
        tag = sender.Tag
        if not tag:
            return
        mix, row_index, field = tag
        text = sender.Text or u''
        if 0 <= row_index < len(mix.rows):
            row = mix.rows[row_index]
            if field == 'code':
                row.code = text
            elif field == 'pct':
                row.pct = text
                self._update_mix_percent_summary(mix)
            elif field == 'spacing':
                row.spacing = text
            elif field == 'bot':
                row.bot = text
            elif field == 'com':
                row.com = text
            elif field == 'grade':
                row.grade = text
            mix.dirty = True

    def on_textbox_got_keyboard_focus(self, sender, args):
        try:
            sender.SelectAll()
        except Exception:
            pass

    def on_textbox_preview_mouse_left_button_down(self, sender, args):
        try:
            if not sender.IsKeyboardFocusWithin:
                args.Handled = True
                sender.Focus()
        except Exception:
            pass

    def on_pct_lost_focus(self, sender, args):
        tag = sender.Tag
        if not tag:
            return
        mix, row_index, field = tag
        if field != 'pct':
            return
        text = sender.Text or u''
        raw = pct_display_to_raw(text)
        if not raw:
            if 0 <= row_index < len(mix.rows):
                mix.rows[row_index].pct = u''
                self._update_mix_percent_summary(mix)
            return
        display = pct_raw_to_display(raw)
        if 0 <= row_index < len(mix.rows):
            mix.rows[row_index].pct = display
        if display != text:
            sender.Text = display
        self._update_mix_percent_summary(mix)

    def on_space_lost_focus(self, sender, args):
        tag = sender.Tag
        if not tag:
            return
        mix, row_index, field = tag
        if field != 'spacing':
            return
        text = sender.Text or u''
        raw = space_display_to_raw(text)
        if not raw:
            if 0 <= row_index < len(mix.rows):
                mix.rows[row_index].spacing = u''
            return
        display = space_raw_to_display(raw)
        if 0 <= row_index < len(mix.rows):
            mix.rows[row_index].spacing = display
        if display != text:
            sender.Text = display

    def on_remove_row(self, sender, args):
        tag = sender.Tag
        if not tag:
            return
        mix, row_index = tag
        mix.remove_row_at(row_index)
        self._render_mix_body(mix)

    def on_add_row(self, sender, args):
        mix = sender.Tag
        if mix is None:
            return
        if len(mix.rows) >= MAX_SPECIES:
            forms.alert('Maximum of {0} species per mix.'.format(MAX_SPECIES))
            return
        mix.add_row()
        self._render_mix_body(mix)

    def on_add_row_from_library(self, sender, args):
        """Add one or more plant rows using the external Create Plant tool."""
        mix = sender.Tag
        if mix is None:
            return

        # ------------------------------------------------------------
        # 0. Work out how many slots are left in this mix
        # ------------------------------------------------------------
        slots_left = MAX_SPECIES - len(mix.rows)
        if slots_left <= 0:
            forms.alert(
                u"Maximum of {0} species per mix reached.".format(MAX_SPECIES),
                title="Add Plant (library)"
            )
            return

        # ------------------------------------------------------------
        # 0a. Current total percent in this mix (0–100)
        # ------------------------------------------------------------
        current_total_percent = 0.0
        for row in mix.rows:
            raw = pct_display_to_raw(row.pct)
            if not raw:
                continue
            try:
                # pct_display_to_raw returns a decimal (e.g. '0.5' for 50%)
                val = float(raw) * 100.0
            except Exception:
                continue
            current_total_percent += val

        percent_remaining = max(0.0, 100.0 - current_total_percent)

        # ------------------------------------------------------------
        # 0b. Most common grade in existing rows, if any
        # ------------------------------------------------------------
        grades = []
        for row in mix.rows:
            g = (row.grade or u'').strip()
            if g:
                grades.append(g)

        most_common_grade = None
        if grades:
            counts = {}
            for g in grades:
                counts[g] = counts.get(g, 0) + 1
            most_common_grade = max(counts, key=counts.get)

        # ------------------------------------------------------------
        # 1. Load the Create Plant script as a module
        # ------------------------------------------------------------
        if not os.path.exists(CREATE_PLANT_SCRIPT_PATH):
            forms.alert(
                u"Could not find Create Plant script at:\n{0}".format(CREATE_PLANT_SCRIPT_PATH),
                title="Add Plant (library)"
            )
            return

        try:
            plant_library = imp.load_source("plant_library_for_mix", CREATE_PLANT_SCRIPT_PATH)
        except Exception as ex:
            forms.alert(
                u"Error loading Create Plant script:\n{0}".format(ex),
                title="Add Plant (library)"
            )
            return

        if not hasattr(plant_library, "open_plant_library_dialog_for_mix"):
            forms.alert(
                u"The Create Plant script has no 'open_plant_library_dialog_for_mix()' function.\n"
                u"Please add one to that script.",
                title="Add Plant (library)"
            )
            return

        # ------------------------------------------------------------
        # 2. Call the plant picker in 'mix mode', passing context
        # ------------------------------------------------------------
        try:
            selected = plant_library.open_plant_library_dialog_for_mix(
                max_slots=slots_left,
                percent_remaining=percent_remaining,
                current_total_percent=current_total_percent,
                most_common_grade=most_common_grade
            )
        except TypeError:
            # Fallback if the other script still has the old signature
            selected = plant_library.open_plant_library_dialog_for_mix()
        except Exception as ex:
            forms.alert(
                u"Error while running plant library dialog:\n{0}".format(ex),
                title="Add Plant (library)"
            )
            return

        # If user cancelled or nothing selected, do nothing
        if not selected:
            return

        # ------------------------------------------------------------
        # 3. For each returned plant, add and populate a SpeciesRow
        # ------------------------------------------------------------
        for row_data in selected:
            if len(mix.rows) >= MAX_SPECIES:
                forms.alert(
                    u"Maximum of {0} species per mix reached.".format(MAX_SPECIES),
                    title="Add Plant (library)"
                )
                break

            mix.add_row()
            row = mix.rows[-1]

            code      = _to_unicode(row_data.get('Code', u''))
            spread_mm = row_data.get('SpreadMM', None)
            botanical = _to_unicode(row_data.get('Botanical', u''))
            common    = _to_unicode(row_data.get('Common', u''))
            percent   = row_data.get('Percent', None)
            grade     = row_data.get('Grade', None)

            row.code = code
            row.bot  = botanical
            row.com  = common

            # Spacing from SpreadMM (mm) -> display string (e.g. '3m')
            if spread_mm not in (None, u''):
                try:
                    raw_mm_str = _to_unicode(spread_mm)
                    row.spacing = space_raw_to_display(raw_mm_str)
                except Exception:
                    row.spacing = u''
            else:
                row.spacing = u''

            # Percent – convert whatever we get ('50', 50, 0.5, '50%')
            if percent not in (None, u''):
                try:
                    row.pct = pct_raw_to_display(percent)
                except Exception:
                    row.pct = u''
            else:
                row.pct = u''

            # Grade – already a display string
            if grade not in (None, u''):
                row.grade = _to_unicode(grade)
            else:
                row.grade = u''

        # 4. Rebuild the UI and refresh the header percent summary
        self._render_mix_body(mix)
        self._update_mix_percent_summary(mix)


    def on_create_new_mix(self, sender, args):
        symbol = self._get_mix_symbol()
        if symbol is None:
            forms.alert('Could not find a FamilySymbol for family "{0}".'.format(FAMILY_NAME))
            return

        doc = self.doc
        new_inst = None

        t = DB.Transaction(doc, 'Create Mix Schedule')
        try:
            t.Start()

            if not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()

            base_x_mm = 1567.922
            base_y_mm = 2686.130
            offset_per_mix_mm = 325.0

            position_index = len(self.mixes)
            x_mm = base_x_mm + offset_per_mix_mm * position_index
            y_mm = base_y_mm

            x_ft = x_mm / 304.8
            y_ft = y_mm / 304.8

            pt = DB.XYZ(x_ft, y_ft, 0.0)

            view = self._get_target_view()
            if view is None:
                raise Exception('Could not determine a view to place the mix schedule.')

            new_inst = doc.Create.NewFamilyInstance(pt, symbol, view)

            set_param(new_inst, PARAM_NUM_SPECIES, 0)

            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert('Failed to create new mix:\n{0}'.format(ex))
            return

        if new_inst is None:
            return

        new_mix = MixModel(new_inst)
        new_mix.num_species = 0
        new_mix.rows = []
        new_mix.is_expanded = True

        try:
            self._attach_color_to_mix(new_mix)
        except Exception:
            pass

        self.mixes.append(new_mix)
        self._refresh_stack_panel()
        self.current_expanded = new_mix

        try:
            self._begin_rename_mix(new_mix)
        except Exception:
            pass

    def on_duplicate_mix(self, sender, args):
        mix = getattr(sender, 'Tag', None)
        if mix is None:
            return

        try:
            args.Handled = True
        except Exception:
            pass

        symbol = self._get_mix_symbol()
        if symbol is None:
            forms.alert('Could not find a FamilySymbol for family "{0}".'.format(FAMILY_NAME))
            return

        doc = self.doc
        new_inst = None

        t = DB.Transaction(doc, 'Duplicate Mix Schedule')
        try:
            t.Start()

            if not symbol.IsActive:
                symbol.Activate()
                doc.Regenerate()

            base_x_mm = 1567.922
            base_y_mm = 2686.130
            offset_per_mix_mm = 325.0

            position_index = len(self.mixes)
            x_mm = base_x_mm + offset_per_mix_mm * position_index
            y_mm = base_y_mm

            x_ft = x_mm / 304.8
            y_ft = y_mm / 304.8

            pt = DB.XYZ(x_ft, y_ft, 0.0)

            view = self._get_target_view()
            if view is None:
                raise Exception('Could not determine a view to place the mix schedule.')

            new_inst = doc.Create.NewFamilyInstance(pt, symbol, view)

            base_name = mix.mix_name or u'(unnamed mix)'
            copy_name = self._generate_copy_name(base_name)

            src_num = getattr(mix, 'num_species', 0)
            if src_num < 0:
                src_num = 0
            if src_num > MAX_SPECIES:
                src_num = MAX_SPECIES

            set_param(new_inst, PARAM_MIX_NAME, copy_name)
            set_param(new_inst, PARAM_NUM_SPECIES, src_num)

            for i in range(1, MAX_SPECIES + 1):
                code_name  = PARAM_SPECIES_CODE_TEMPLATE.format(i)
                pct_name   = PARAM_SPECIES_PCT_TEMPLATE.format(i)
                space_name = PARAM_SPECIES_SPACE_TEMPLATE.format(i)
                bot_name   = PARAM_SPECIES_BOT_TEMPLATE.format(i)
                com_name   = PARAM_SPECIES_COM_TEMPLATE.format(i)
                grade_name = PARAM_SPECIES_GRADE_TEMPLATE.format(i)

                for pname in (code_name, pct_name, space_name, bot_name, com_name, grade_name):
                    copy_param_between_elements(mix.element, new_inst, pname)

            t.Commit()
        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert('Failed to duplicate mix:\n{0}'.format(ex))
            return

        if new_inst is None:
            return

        new_mix = MixModel(new_inst)
        new_mix.is_expanded = True

        try:
            self._attach_color_to_mix(new_mix)
        except Exception:
            pass

        self.mixes.append(new_mix)
        self._refresh_stack_panel()
        self.current_expanded = new_mix

        try:
            self._begin_rename_mix(new_mix)
        except Exception:
            pass

    def _apply_color_scheme_color_updates(self):
        """Push any pending Area colour updates back into the Color Fill Scheme."""
        scheme = self._color_scheme
        if scheme is None:
            return

        mixes_with_new = []
        for mix in self.mixes:
            if getattr(mix, 'area_color_entry', None) is not None                     and getattr(mix, 'area_color_new_dbcolor', None) is not None:
                mixes_with_new.append(mix)

        if not mixes_with_new:
            return

        try:
            entries = list(scheme.GetEntries())
        except Exception:
            try:
                entries = list(scheme.Entries)
            except Exception:
                entries = []

        if not entries:
            return

        changed = False

        for mix in mixes_with_new:
            target_val, target_cap = _get_color_entry_keys(mix.area_color_entry)
            new_col = mix.area_color_new_dbcolor

            for e in entries:
                v_key, c_key = _get_color_entry_keys(e)

                if (target_val and v_key == target_val) or                    (not target_val and target_cap and c_key == target_cap) or                    (target_cap and c_key == target_cap):
                    try:
                        e.Color = new_col
                        changed = True
                        mix.area_color_dbcolor = new_col
                        mix.area_color_new_dbcolor = None
                    except Exception:
                        pass
                    break

        if changed:
            try:
                scheme.SetEntries(entries)
            except Exception:
                LOGGER.warning('Failed to SetEntries on ColorFillScheme for updated colours.')

    def _update_filled_region_strips(self):
        """Ensure each mix has a FilledRegionType + strip matching its Area colour."""
        doc = self.doc

        host_view = self._get_target_view()
        if host_view is None:
            try:
                host_view = doc.ActiveView
            except Exception:
                host_view = None

        for mix in self.mixes:
            col = getattr(mix, 'area_color_new_dbcolor', None)
            if col is None:
                col = getattr(mix, 'area_color_dbcolor', None)

            if col is None:
                fr_debug(
                    u'FR: Mix "{0}" has no colour (no scheme entry or unset); skipping.'
                    .format(mix.mix_name)
                )
                continue

            fr_debug(
                u'FR: Mix "{0}" using colour RGB({1},{2},{3}) for filled region.'
                .format(mix.mix_name, col.Red, col.Green, col.Blue)
            )

            try:
                _ensure_filled_region_for_mix(doc, mix.mix_name, col, host_view)
            except Exception as ex:
                fr_debug(
                    u'FR: Exception in _ensure_filled_region_for_mix for mix "{0}": {1}'
                    .format(mix.mix_name, ex)
                )
                continue

        # optional: force a regenerate after all changes
        try:
            doc.Regenerate()
        except Exception:
            pass


    def on_apply(self, sender, args):
        if not self.mixes:
            return

        t = DB.Transaction(self.doc, 'Update Mix Schedules')
        try:
            t.Start()
            for mix in self.mixes:
                num = len(mix.rows)
                if num > MAX_SPECIES:
                    num = MAX_SPECIES

                set_param(mix.element, PARAM_MIX_NAME, mix.mix_name or u'')
                set_param(mix.element, PARAM_NUM_SPECIES, num)

                for i in range(1, MAX_SPECIES + 1):
                    code_name  = PARAM_SPECIES_CODE_TEMPLATE.format(i)
                    pct_name   = PARAM_SPECIES_PCT_TEMPLATE.format(i)
                    space_name = PARAM_SPECIES_SPACE_TEMPLATE.format(i)
                    bot_name   = PARAM_SPECIES_BOT_TEMPLATE.format(i)
                    com_name   = PARAM_SPECIES_COM_TEMPLATE.format(i)
                    grade_name = PARAM_SPECIES_GRADE_TEMPLATE.format(i)

                    if i <= num:
                        row = mix.rows[i - 1]

                        code_val = row.code or u''

                        pct_raw = pct_display_to_raw(row.pct)
                        pct_val = pct_raw if pct_raw not in (None, u'') else u''

                        space_raw_mm = space_display_to_raw(row.spacing)
                        space_mm = space_raw_mm if space_raw_mm not in (None, u'') else u''

                        bot_val = row.bot or u''
                        com_val = row.com or u''
                        grade_val = row.grade or u''
                    else:
                        code_val = u''
                        pct_val = u''
                        space_mm = u''
                        bot_val = u''
                        com_val = u''
                        grade_val = u''

                    set_param(mix.element, code_name,  code_val)
                    set_param(mix.element, pct_name,   pct_val)
                    set_param(mix.element, bot_name,   bot_val)
                    set_param(mix.element, com_name,   com_val)
                    set_param(mix.element, grade_name, grade_val)

                    try:
                        p_space = mix.element.LookupParameter(space_name)
                    except Exception:
                        p_space = None
                    if p_space and p_space.StorageType == DB.StorageType.Double:
                        try:
                            if space_mm in (None, u'', u''):
                                p_space.Set(0.0)
                            else:
                                mm_val = float(space_mm)
                                ft_val = mm_val / 304.8
                                p_space.Set(ft_val)
                        except Exception:
                            pass
                    else:
                        set_param(mix.element, space_name, space_mm)

            if self._area_renames:
                for (old_name, new_name) in self._area_renames:
                    old_name = _to_unicode(old_name).strip()
                    new_name = _to_unicode(new_name).strip()
                    if not old_name or not new_name or old_name == new_name:
                        continue

                    areas = (DB.FilteredElementCollector(self.doc)
                             .OfCategory(DB.BuiltInCategory.OST_Areas)
                             .WhereElementIsNotElementType())

                    for area in areas:
                        nm = get_param(area, AREA_NAME_PARAM)
                        if nm and _to_unicode(nm).strip() == old_name:
                            set_param(area, AREA_NAME_PARAM, new_name)

                self._area_renames = []

            self._apply_color_scheme_color_updates()
            self._update_filled_region_strips()

            t.Commit()
            # After committing, show a debug popup with everything we logged
            show_fr_debug_popup()


        except Exception as ex:
            try:
                t.RollBack()
            except Exception:
                pass
            forms.alert('Failed to update mix schedules:\n{0}'.format(ex))
            return

        LOGGER.info(
            'Mix schedules, related Areas, Area names, Area colours, '
            'and FilledRegion strips updated.'
        )

    def on_close(self, sender, args):
        # Run the schedule header script when the user closes the window
        if (self._ghost_area_id is not None or
                self._ghost_boundary_ids or
                self._ghost_sketch_plane_id is not None):
            _schedule_ghost_cleanup(
                self.doc,
                self._ghost_area_id,
                self._ghost_boundary_ids,
                self._ghost_sketch_plane_id
            )
            self._ghost_area_id = None
            self._ghost_boundary_ids = []
            self._ghost_sketch_plane_id = None

        if self.window is not None:
            self.window.Close()

    def show(self):
        if self.window is not None:
            self.window.ShowDialog()


# -----------------------------
# Run
# -----------------------------
controller = MixWindowController(revit.doc)
if controller.window is not None:
    controller.show()
