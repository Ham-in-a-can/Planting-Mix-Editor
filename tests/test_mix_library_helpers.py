# coding: utf-8
import unittest
import os

from mix_library_helpers import (
    apply_library_payload_to_row,
    invoke_plant_library,
)


class Row(object):
    def __init__(self, groundcover=True):
        self.code = self.pct = self.spacing = u''
        self.bot = self.com = self.grade = u''
        self.is_groundcover = groundcover


def percent(value):
    text = str(value).replace('%', '')
    number = float(text)
    if '%' not in str(value) and abs(number) < 1.0 - 1e-9:
        number *= 100.0
    return ('%g' % number) + '%'


def spread(value):
    return ('%g' % (float(value) / 1000.0)) + 'm'


class PayloadTests(unittest.TestCase):
    def apply(self, payload, groundcover=True):
        row = Row(groundcover)
        result = apply_library_payload_to_row(row, payload, percent, spread)
        return row, result

    def test_resolved_model_values_are_authoritative(self):
        row, result = self.apply({
            'Code': 'MODEL01', 'Botanical': 'Acaena inermis',
            'Common': 'Model common', 'Grade': 'Model grade',
            'SpreadMM': 1200,
        })
        self.assertEqual((True, u''), result)
        self.assertEqual(
            ('MODEL01', 'Acaena inermis', 'Model common', 'Model grade', '1.2m'),
            (row.code, row.bot, row.com, row.grade, row.spacing)
        )

    def test_adopted_library_values_and_explicit_spacing(self):
        row, _ = self.apply({
            'Code': 'LIB01', 'Botanical': 'Poa cita', 'Common': 'Silver tussock',
            'Grade': '1.5L', 'SpreadMM': 1000,
        })
        self.assertEqual(('LIB01', '1m'), (row.code, row.spacing))
        row, _ = self.apply({
            'Botanical': 'Poa cita', 'Spacing': '1.35m', 'SpreadMM': 1000,
        })
        self.assertEqual('1.35m', row.spacing)

    def test_classification_and_branch_default(self):
        row, _ = self.apply({'Botanical': 'One', 'IsGroundcover': False})
        self.assertFalse(row.is_groundcover)
        row, _ = self.apply({'Botanical': 'Two', 'IsTree': True})
        self.assertFalse(row.is_groundcover)
        row, _ = self.apply({'Botanical': 'Three'}, groundcover=False)
        self.assertFalse(row.is_groundcover)

    def test_percentage_grade_and_existing_rows(self):
        existing = Row()
        existing.pct = '75%'
        row, _ = self.apply({'Botanical': 'One', 'Percent': 0.25, 'Grade': ''})
        self.assertEqual('25%', row.pct)
        self.assertEqual('', row.grade)
        self.assertEqual('75%', existing.pct)
        row, _ = self.apply({'Botanical': 'Two', 'Grade': 'Resolved 2L'})
        self.assertEqual('Resolved 2L', row.grade)

    def test_malformed_is_skipped_without_mutation(self):
        row, result = self.apply({'Botanical': '', 'Code': 'BAD'})
        self.assertFalse(result[0])
        self.assertEqual('', row.code)
        valid, result = self.apply({'Botanical': 'Valid', 'SpreadMM': 1000})
        self.assertTrue(result[0])
        self.assertEqual('Valid', valid.bot)
        row, result = self.apply({'Botanical': 'Bad spread', 'SpreadMM': 'wide'})
        self.assertFalse(result[0])


class InvocationTests(unittest.TestCase):
    def test_modern_and_old_signatures_receive_supported_context(self):
        calls = []

        def modern(max_slots=None, percent_remaining=None,
                   current_total_percent=None, most_common_grade=None,
                   mix_context=None):
            calls.append(mix_context)
            return 'modern'

        self.assertEqual('modern', invoke_plant_library(modern, {
            'max_slots': 2, 'percent_remaining': 50,
            'current_total_percent': 50, 'most_common_grade': '2L',
            'mix_context': {'mix_name': 'A'},
        }))
        self.assertEqual({'mix_name': 'A'}, calls[0])

        def old(max_slots=None, percent_remaining=None):
            return max_slots, percent_remaining

        self.assertEqual((2, 50), invoke_plant_library(old, {
            'max_slots': 2, 'percent_remaining': 50, 'mix_context': {},
        }))
        self.assertEqual('zero', invoke_plant_library(lambda: 'zero', {'max_slots': 2}))

    def test_internal_type_error_runs_once(self):
        calls = []

        def broken(max_slots=None, mix_context=None):
            calls.append(1)
            raise TypeError('internal failure')

        with self.assertRaisesRegex(TypeError, 'internal failure'):
            invoke_plant_library(broken, {'max_slots': 2, 'mix_context': {}})
        self.assertEqual(1, len(calls))


class EditorRefreshContractTests(unittest.TestCase):
    def test_successful_import_keeps_sort_and_refresh_pipeline(self):
        root = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        with open(os.path.join(root, 'script.py'), 'r') as source_file:
            source = source_file.read()
        start = source.index('    def on_add_row_from_library(')
        end = source.index('\n    def on_create_new_mix(', start)
        handler = source[start:end]
        calls = [
            '_sort_mix_rows_alphabetically(mix)',
            'self._render_mix_body(mix)',
            'self._update_mix_percent_summary(mix)',
            'self._update_approx_numbers_for_mix(mix)',
        ]
        positions = [handler.rindex(call) for call in calls]
        self.assertEqual(sorted(positions), positions)


if __name__ == '__main__':
    unittest.main()
