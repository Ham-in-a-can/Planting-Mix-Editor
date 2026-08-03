# coding: utf-8
"""Pure compatibility helpers for the Mix Editor/Plant Library boundary.

This module deliberately has no pyRevit or Revit imports so its payload and
function-signature behaviour can be regression tested outside Revit.
"""


def _nonblank(value):
    return value is not None and unicode_text(value).strip() != u''


def unicode_text(value):
    if value is None:
        return u''
    try:
        return unicode(value)  # noqa: F821 - available under IronPython 2
    except NameError:
        return str(value)


def _classification(payload, default_groundcover):
    if _nonblank(payload.get('IsGroundcover', None)):
        return _yes_no_bool(payload.get('IsGroundcover'), default_groundcover)
    if _nonblank(payload.get('IsTree', None)):
        return not _yes_no_bool(payload.get('IsTree'), not default_groundcover)
    return default_groundcover


def _yes_no_bool(value, default):
    text = unicode_text(value).strip().lower()
    if text in (u'1', u'true', u'yes', u'y'):
        return True
    if text in (u'0', u'false', u'no', u'n'):
        return False
    return default


def apply_library_payload_to_row(row, payload, percent_converter,
                                 spread_converter, default_percent=u'10%'):
    """Validate and apply one authoritative Plant Library payload to ``row``.

    Returns ``(success, reason)``. The row is not changed on validation failure.
    Plant Library display strings are retained; only SpreadMM and Percent use
    the Mix Editor's established unit/percentage conversion functions.
    """
    if payload is None or not callable(getattr(payload, 'get', None)):
        return False, u'payload is not dictionary-like'

    botanical = unicode_text(payload.get('Botanical', u'')).strip()
    if not botanical:
        return False, u'botanical name is blank'

    spacing_supplied = _nonblank(payload.get('Spacing', None))
    spread = payload.get('SpreadMM', None)
    final_spacing = u''
    if spacing_supplied:
        final_spacing = unicode_text(payload.get('Spacing')).strip()
    elif _nonblank(spread):
        spread_text = unicode_text(spread).strip().replace(u',', u'.')
        try:
            float(spread_text)
        except (TypeError, ValueError):
            return False, u'SpreadMM is not numeric'
        final_spacing = spread_converter(spread_text)

    percent = payload.get('Percent', None)
    if _nonblank(percent):
        final_percent = percent_converter(percent)
    else:
        final_percent = default_percent

    # Validate everything before mutating the target row.
    values = {
        'code': unicode_text(payload.get('Code', u'')),
        'bot': botanical,
        'com': unicode_text(payload.get('Common', u'')),
        'grade': unicode_text(payload.get('Grade', u'')),
        'spacing': final_spacing,
        'pct': final_percent,
        'is_groundcover': _classification(
            payload, getattr(row, 'is_groundcover', True)
        ),
    }
    for name, value in values.items():
        setattr(row, name, value)
    return True, u''


def supported_keyword_arguments(function, context):
    """Return supported context kwargs, or ``None`` if introspection fails."""
    code = getattr(function, 'func_code', None)
    if code is None:
        code = getattr(function, '__code__', None)
    if code is not None:
        try:
            count = int(code.co_argcount)
            names = tuple(code.co_varnames[:count])
            # CO_VARKEYWORDS: function accepts **kwargs.
            if int(getattr(code, 'co_flags', 0)) & 0x08:
                return dict(context)
            return dict((key, value) for key, value in context.items()
                        if key in names)
        except (AttributeError, TypeError, ValueError):
            pass

    # IronPython implementations vary; getargspec is a useful second source.
    try:
        import inspect
        spec = inspect.getargspec(function)
        if spec.keywords:
            return dict(context)
        return dict((key, value) for key, value in context.items()
                    if key in spec.args)
    except (AttributeError, TypeError, ValueError):
        return None


def invoke_plant_library(function, context):
    """Invoke the library exactly once with only arguments it supports."""
    supported = supported_keyword_arguments(function, context)
    if supported is None:
        # Defensive last resort: make one modern call. Any TypeError is allowed
        # to propagate as the real error rather than risking a repeated dialog.
        return function(**context)
    if supported:
        return function(**supported)
    return function()
