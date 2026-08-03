# coding: utf-8
"""Pure, IronPython-compatible helpers for the Plant Library integration."""

import math


CONTEXT_ARGUMENTS = (
    'max_slots',
    'percent_remaining',
    'current_total_percent',
    'most_common_grade',
    'mix_context',
)


def _function_signature(function):
    """Return ``(argument names, accepts kwargs)``, or ``None``."""
    code = getattr(function, 'func_code', None) or getattr(function, '__code__', None)
    if code is None:
        return None
    return (tuple(code.co_varnames[:code.co_argcount]), bool(code.co_flags & 0x08))


def _is_signature_type_error(error):
    """Recognise only errors raised while Python is binding call arguments."""
    message = str(error).lower()
    markers = (
        'unexpected keyword argument',
        'takes no arguments',
        'takes exactly 0 arguments',
        'takes 0 positional arguments',
    )
    return any(marker in message for marker in markers)


def invoke_plant_library(function, context):
    """Invoke a Plant Library entry point once with only supported context.

    IronPython exposes ``func_code`` while CPython exposes ``__code__``.  If
    neither is available, a tightly-scoped compatibility retry is allowed only
    for a recognisable argument-binding error; exceptions from inside the
    library are propagated without a retry.
    """
    signature = _function_signature(function)
    if signature is not None:
        names, accepts_kwargs = signature
        supported = dict((key, context[key]) for key in CONTEXT_ARGUMENTS
                         if (accepts_kwargs or key in names) and key in context)
        return function(**supported)

    kwargs = dict((key, context[key]) for key in CONTEXT_ARGUMENTS if key in context)
    try:
        return function(**kwargs)
    except TypeError as error:
        if not _is_signature_type_error(error):
            raise
        return function()


def species_update_from_payload(payload, percent_converter, spread_converter,
                                default_is_groundcover=True):
    """Validate one resolved payload and return authoritative row values."""
    if not hasattr(payload, 'get'):
        raise ValueError('payload is not dictionary-like')

    botanical = payload.get('Botanical', u'')
    botanical = u'' if botanical is None else unicode_value(botanical).strip()
    if not botanical:
        raise ValueError('Botanical is required')

    try:
        keys = payload.keys()
    except Exception:
        raise ValueError('payload is not dictionary-like')

    if 'Spacing' in keys and payload.get('Spacing') not in (None, u'', ''):
        spacing = unicode_value(payload.get('Spacing')).strip()
    else:
        spread = payload.get('SpreadMM', None)
        if spread in (None, u'', ''):
            spacing = u''
        else:
            try:
                numeric_spread = float(spread)
            except (TypeError, ValueError):
                raise ValueError('SpreadMM must be numeric')
            if math.isnan(numeric_spread) or math.isinf(numeric_spread):
                raise ValueError('SpreadMM must be finite')
            spacing = spread_converter(unicode_value(spread))

    percent = payload.get('Percent', None)
    pct = u'' if percent in (None, u'', '') else percent_converter(percent)

    if 'IsGroundcover' in keys:
        is_groundcover = bool(payload.get('IsGroundcover'))
    elif 'IsTree' in keys:
        is_groundcover = not bool(payload.get('IsTree'))
    else:
        is_groundcover = bool(default_is_groundcover)

    return {
        'code': unicode_value(payload.get('Code', u'')),
        'bot': botanical,
        'com': unicode_value(payload.get('Common', u'')),
        'grade': unicode_value(payload.get('Grade', u'')),
        'pct': pct,
        'spacing': spacing,
        'is_groundcover': is_groundcover,
    }


def unicode_value(value):
    if value is None:
        return u''
    try:
        return unicode(value)
    except NameError:
        return str(value)
