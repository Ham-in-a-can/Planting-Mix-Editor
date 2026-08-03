# coding: utf-8

import unittest

from library_integration import invoke_plant_library, species_update_from_payload


def percent_display(value):
    text = str(value).replace('%', '')
    number = float(text)
    if number <= 1.0:
        number *= 100.0
    return ('%g' % number) + '%'


def spread_display(value):
    return ('%g' % (float(value) / 1000.0)) + 'm'


class PayloadTests(unittest.TestCase):
    def convert(self, payload, default=True):
        return species_update_from_payload(
            payload, percent_display, spread_display, default
        )

    def test_resolved_model_values_are_authoritative(self):
        result = self.convert({
            'Code': 'MODEL01', 'Botanical': 'Acer rubrum',
            'Common': 'Model Common', 'Grade': 'Model Grade',
            'SpreadMM': 1200,
        })
        self.assertEqual('MODEL01', result['code'])
        self.assertEqual('Model Common', result['com'])
        self.assertEqual('Model Grade', result['grade'])
        self.assertEqual('1.2m', result['spacing'])

    def test_adopted_library_values_and_explicit_spacing(self):
        library = self.convert({
            'Code': 'LIB01', 'Botanical': 'Carex test', 'SpreadMM': 1000,
        })
        explicit = self.convert({
            'Botanical': 'Carex test', 'Spacing': '1.35m', 'SpreadMM': 1000,
        })
        self.assertEqual(('LIB01', '1m'), (library['code'], library['spacing']))
        self.assertEqual('1.35m', explicit['spacing'])

    def test_classification_and_branch_default(self):
        self.assertFalse(self.convert({'Botanical': 'Tree', 'IsGroundcover': False})['is_groundcover'])
        self.assertFalse(self.convert({'Botanical': 'Tree', 'IsTree': True})['is_groundcover'])
        self.assertFalse(self.convert({'Botanical': 'Unknown'}, False)['is_groundcover'])

    def test_percent_and_blank_grade(self):
        result = self.convert({'Botanical': 'Plant', 'Percent': '0.5', 'Grade': ''})
        self.assertEqual('50%', result['pct'])
        self.assertEqual('', result['grade'])

    def test_malformed_rows_can_be_skipped_independently(self):
        payloads = [None, {'Botanical': ''}, {'Botanical': 'Valid', 'SpreadMM': 1000}]
        valid = []
        for payload in payloads:
            try:
                valid.append(self.convert(payload))
            except (TypeError, ValueError):
                pass
        self.assertEqual(['Valid'], [item['bot'] for item in valid])


class InvocationTests(unittest.TestCase):
    def test_modern_signature_receives_mix_context(self):
        calls = []

        def modern(max_slots, percent_remaining, current_total_percent,
                   most_common_grade, mix_context):
            calls.append(mix_context)
            return 'ok'

        result = invoke_plant_library(modern, {
            'max_slots': 2, 'percent_remaining': 30,
            'current_total_percent': 70, 'most_common_grade': 'PB5',
            'mix_context': {'name': 'Mix'},
        })
        self.assertEqual('ok', result)
        self.assertEqual([{'name': 'Mix'}], calls)

    def test_kwargs_signature_receives_all_context(self):
        def modern(**kwargs):
            return kwargs

        context = {'max_slots': 2, 'mix_context': {'name': 'Mix'}}
        self.assertEqual(context, invoke_plant_library(modern, context))

    def test_old_and_zero_argument_signatures_receive_only_supported_values(self):
        def old(max_slots, percent_remaining):
            return max_slots, percent_remaining

        def ancient():
            return 'ancient'

        context = {'max_slots': 3, 'percent_remaining': 25, 'mix_context': {}}
        self.assertEqual((3, 25), invoke_plant_library(old, context))
        self.assertEqual('ancient', invoke_plant_library(ancient, context))

    def test_internal_type_error_is_not_retried(self):
        calls = []

        def broken(max_slots):
            calls.append(max_slots)
            raise TypeError('internal library failure')

        with self.assertRaises(TypeError):
            invoke_plant_library(broken, {'max_slots': 1})
        self.assertEqual([1], calls)


if __name__ == '__main__':
    unittest.main()
