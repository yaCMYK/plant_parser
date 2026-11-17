import functools
import http.server
import json
from pathlib import Path
import subprocess
import sys
import tempfile
import threading
import unittest
import xml.etree.ElementTree as ET
import zipfile

from bs4 import BeautifulSoup
import plant_parser_production as p


NS = {'s': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}


def workbook_rows(path, sheet):
    with zipfile.ZipFile(path) as archive:
        strings = [''.join(n.itertext()) for n in ET.fromstring(archive.read('xl/sharedStrings.xml'))]
        rows = []
        for row in ET.fromstring(archive.read(f'xl/worksheets/sheet{sheet}.xml')).findall('.//s:row', NS):
            values = {}
            for cell in row:
                value = cell.find('s:v', NS)
                if value is not None:
                    values[cell.attrib['r']] = strings[int(value.text)] if cell.get('t') == 's' else value.text
            rows.append(values)
        return rows


class ParserTests(unittest.TestCase):
    def test_length_ranges_and_units(self):
        for raw, lo, hi in [('150–200 cm', 1.5, 2), ('6 ft', 1.8288, 1.8288), ('20 mm', .02, .02)]:
            with self.subTest(raw=raw):
                v = p.parse_length(raw)[0]
                self.assertAlmostEqual(v.min_value, lo)
                self.assertAlmostEqual(v.max_value, hi)
                self.assertEqual(v.canonical_unit, 'm')

    def test_one_sided_lengths(self):
        for raw, lo, hi in [('до 2 m', None, 2), ('от 1.5 m', 1.5, None), ('max. 2 m', None, 2)]:
            with self.subTest(raw=raw):
                v = p.parse_length(raw)[0]
                self.assertEqual((v.min_value, v.max_value), (lo, hi))

    def test_ph_bounds(self):
        self.assertEqual(p.parse_ph('pH 5.5–7.0')[0].normalized_value, 'pH 5.5–7')
        self.assertEqual(p.parse_ph('14')[0].min_value, 14)
        for raw in ['15', '-1', '150', 'sandy soil']:
            with self.subTest(raw=raw):
                self.assertEqual(p.parse_ph(raw), [])

    def test_hardiness_systems(self):
        self.assertEqual(p.parse_hardiness('USDA zone 4–6')[0].normalized_value, 'USDA 4–6')
        self.assertFalse(any(v.canonical_unit == 'USDA zone' for v in p.parse_hardiness('hardy')))
        self.assertEqual(p.parse_temperature('-30 °C')[0].min_value, -30)
        self.assertFalse(any(v.normalized_value == 'зимостойкое' for v in p.parse_hardiness('not hardy')))

    def test_temperature_range(self):
        v = p.parse_temperature('-30 to -20 °C')[0]
        self.assertEqual((v.min_value, v.max_value), (-30, -20))
        self.assertAlmostEqual(p.parse_temperature('-4 °F')[0].min_value, -20)

    def test_explicit_tolerance(self):
        self.assertEqual(p.parse_explicit_tolerance('prefers wet soil', 'waterlogging'), [])
        self.assertEqual(p.parse_explicit_tolerance('waterlogging tolerant', 'waterlogging')[0].normalized_value, 'устойчиво / переносит')
        self.assertEqual(p.parse_explicit_tolerance('not drought tolerant', 'drought')[0].normalized_value, 'неустойчиво / не переносит')
        self.assertEqual(p.parse_rule_value(p.RULE_BY_KEY['waterlogging_tolerance'], 'High')[0].normalized_value, 'высокая')

    def test_container_and_suitability(self):
        values = p.parse_container_size('C5; 500 ml; 40×60 cm')
        self.assertTrue(any(v.normalized_value == '0.5 L' for v in values))
        self.assertTrue(any(v.normalized_value == 'C5' for v in values))
        self.assertTrue(any(v.normalized_value == '0.4×0.6 m' for v in values))
        self.assertFalse(p.context_requirements_ok(p.Candidate('Container size', 'C5', 'Container size C5', 'test'), p.RULE_BY_KEY['in_container']))
        self.assertEqual(p.parse_rule_value(p.RULE_BY_KEY['in_container'], 'Yes')[0].normalized_value, 'да')

    def test_growth_and_currency(self):
        v = p.parse_growth_rate('20–30 cm/year')[0]
        self.assertEqual((v.min_value, v.max_value, v.canonical_unit), (.2, .3, 'm/year'))
        self.assertEqual(p.parse_currency('€ 12.50')[0].normalized_value, '12.5 EUR')
        self.assertFalse(p.context_requirements_ok(p.Candidate('price', '12 EUR', 'planting price', 'test'), p.RULE_BY_KEY['cost_material']))

    def test_all_values_and_unknown_evidence_survive(self):
        candidates = [p.Candidate('Высота во взрослом виде', '150 cm; 2 m', 'Высота во взрослом виде', 'test'),
                      p.Candidate('Unknown custom attribute', 'original text', '', 'test')]
        accepted, rejected = p.map_candidates_to_observations('id', candidates, 'https://example.org/p', 'now')
        record = p.PlantRecord('id', 'https://example.org/p')
        for obs in accepted:
            record.add(obs)
        self.assertEqual(len(record.observations['mature_height']), 2)
        self.assertEqual(p.aggregate_cell(record, 'mature_height', 'cm'), '150 cm\n200 cm')
        self.assertTrue(any(o.raw_value == 'original text' for o in rejected))
        self.assertTrue(all(o.raw_value == '150 cm; 2 m' for o in accepted))

    def test_definition_list_context_is_local(self):
        soup = BeautifulSoup('<dl><dt>Высота во взрослом виде</dt><dd>2 m</dd><dt>Высота саженца</dt><dd>50 cm</dd></dl>', 'lxml')
        candidates = p.extract_label_value_candidates(soup)
        self.assertNotIn('саженца', candidates[0].context)

    def test_raw_variants_remain_distinct(self):
        values = p.parse_length('100 cm; 100 CM')
        self.assertEqual(len(values), 2)

    def test_container_dimensions_display(self):
        values = p.parse_container_size('40×60 cm')
        v = next(v for v in values if '×' in v.normalized_value)
        obs = p.Observation('id', 'container_size', '', '', '40×60 cm', v.normalized_value, v.min_value, v.max_value, v.source_unit, v.canonical_unit, 1, 'test', '', '', '')
        self.assertEqual(p.display_observation(obs, 'cm'), '40×60 cm')


class WorkflowTests(unittest.TestCase):
    def test_cli_crawl_and_xlsx(self):
        class Handler(http.server.SimpleHTTPRequestHandler):
            def log_message(self, *args):
                pass

        with tempfile.TemporaryDirectory() as temp:
            root = Path(temp)
            (root / 'robots.txt').write_text('User-agent: *\nDisallow: /blocked\n')
            (root / 'index.html').write_text('<html><h1>Каталог</h1><a href="/plant.html">Карточка</a><a href="/blocked">Blocked</a></html>')
            (root / 'plant.html').write_text('<html><meta charset="utf-8"><h1>Pinus sylvestris</h1><table><tr><td>Высота во взрослом виде</td><td>150–200 cm; 3 m</td></tr><tr><td>Unknown custom attribute</td><td>=1+1</td></tr></table></html>')
            server = http.server.ThreadingHTTPServer(('127.0.0.1', 0), functools.partial(Handler, directory=temp))
            thread = threading.Thread(target=server.serve_forever, daemon=True)
            thread.start()
            try:
                output = root / 'plants.xlsx'
                result = subprocess.run([sys.executable, str(Path(p.__file__)), '--url', f'http://127.0.0.1:{server.server_port}/index.html', '--browser', 'requests', '--delay', '0', '--output', str(output)], capture_output=True, text=True)
                self.assertEqual(result.returncode, 0, result.stderr)
                report = json.loads(result.stdout)
                self.assertEqual(report['unique_plants'], 1)
                self.assertEqual(report['errors'], 0)
                with zipfile.ZipFile(output) as archive:
                    names = [n.get('name') for n in ET.fromstring(archive.read('xl/workbook.xml')).find('s:sheets', NS)]
                    self.assertEqual(names, ['База растений', 'Raw evidence', 'Правила полей', 'Crawl log', 'Отчёт парсинга'])
                    self.assertNotIn(b'<f>', archive.read('xl/worksheets/sheet2.xml'))
                evidence = workbook_rows(output, 2)
                self.assertTrue(any('=1+1' in row.values() for row in evidence))
                self.assertTrue(any('150–200 cm; 3 m' in row.values() and '1.5' in row.values() and '2' in row.values() for row in evidence))
                self.assertTrue(any('blocked_by_robots' in row.values() for row in workbook_rows(output, 4)))
                self.assertTrue(any('1.5–2 m\n3 m' in row.values() for row in workbook_rows(output, 1)))
            finally:
                server.shutdown()
                server.server_close()
                thread.join()

    def test_schema_columns_match(self):
        path = Path(p.__file__).with_name('База_растений_финальная_схема.xlsx')
        with zipfile.ZipFile(path) as archive:
            row = ET.fromstring(archive.read('xl/worksheets/sheet1.xml')).find('.//s:row', NS)
            headers = [''.join(c.itertext()) for c in row]
        self.assertEqual([r.label for r in p.FIELD_RULES], headers)


if __name__ == '__main__':
    unittest.main()
