"""Regression cases from the supplied Greeninfo workbook; HTML below is synthetic."""
import io
import json
from pathlib import Path
import tempfile
import unittest
from unittest.mock import patch
from bs4 import BeautifulSoup
import plant_parser_production as p
from site_profiles import allowed_link
from crawl_history import CrawlHistory

ROOT = 'https://www.greeninfo.ru/decor_trees/cornus_alba.html'
NEXT = 'https://www.greeninfo.ru/decor_trees/cornus_kousa.html'

def html(title='Дерен белый', link=NEXT):
    return f'''<html><h1>{title}</h1><main><i>Cornus alba</i>
    <table><tr><td>Высота во взрослом виде</td><td>2–3 м</td></tr></table>
    <p>Полное описание растения с особенностями выращивания.</p>
    <p><b>Зимостойкость:</b> высокая зимостойкость</p>
    <a href="{link}">Дерен Коуза</a><a href="/img/work/catalog/a.JPG">Фото</a>
    <a href="{ROOT}/Forum">Форум</a><a href="/">Главная</a>
    <a href="/decor_trees/cornus.html/Social">Сообщество</a>
    </main><aside><p>Использование удобрений: помогая - не навредить</p></aside></html>'''

class GreeninfoTests(unittest.TestCase):
    def test_tracking_and_query_order_have_one_identity(self):
        self.assertEqual(p.normalize_url('https://WWW.greeninfo.ru:443/decor_trees/cornus_alba.html?b=2&a=1&utm_source=x#photo'), p.normalize_url(ROOT+'?a=1&b=2'))

    def test_routing_matches_actual_log(self):
        soup = BeautifulSoup(html(), 'lxml')
        self.assertTrue(p.is_probable_plant_page(soup, ROOT))
        for url in ['https://www.greeninfo.ru/', ROOT+'/Forum', ROOT+'/Article',
                    'https://www.greeninfo.ru/decor_trees/cornus.html/Social',
                    'https://www.greeninfo.ru/decor_trees/berberis_thunbergii/miniatyurnye-rasteniya-dlya-malenkogo-sada_art.html']:
            self.assertFalse(p.is_probable_plant_page(soup, url), url)
        self.assertEqual(p.extract_internal_links(soup, ROOT, ROOT), [NEXT])
        self.assertFalse(allowed_link('https://www.greeninfo.ru/img/photo.JPG'))
        self.assertFalse(p.is_probable_plant_page(BeautifulSoup('<h1>Contact information</h1>', 'lxml')))

    def test_card_fields_and_full_text(self):
        record, rejected = p.build_plant_record(ROOT, BeautifulSoup(html(), 'lxml'), 'now')
        self.assertEqual(p.aggregate_cell(record, 'ru_genus'), 'Дерен')
        self.assertEqual(p.aggregate_cell(record, 'ru_species'), 'белый')
        self.assertEqual(p.aggregate_cell(record, 'latin_species'), 'alba')
        self.assertEqual(p.aggregate_cell(record, 'mature_height'), '2–3 m')
        self.assertIn('Полное описание', record.page_text)
        self.assertNotIn('Использование удобрений', record.page_text)
        self.assertEqual(p.aggregate_cell(record, 'planting_type'), '')
        _, rejected = p.map_candidates_to_observations('x', [p.Candidate('Использование удобрений', 'помогая - не навредить', '', 'test')], ROOT, 'now')
        self.assertTrue(rejected)

    def test_resume_with_durable_history_and_retry_error(self):
        calls = []
        class Fetcher:
            def __init__(self, *args): self.session = None
            def close(self): pass
            def fetch(self, url):
                calls.append(url)
                if url == NEXT and calls.count(NEXT) == 1: raise RuntimeError('temporary failure')
                return html('Дерен белый' if url == ROOT else 'Дерен Коуза', NEXT if url == ROOT else ROOT).replace('<i>Cornus alba</i>', '<i>Cornus alba</i>' if url == ROOT else '<i>Cornus kousa</i>'), 200
        with tempfile.TemporaryDirectory() as directory, patch.object(p, 'Fetcher', Fetcher):
            args = dict(root_url=ROOT, max_pages=1, browser='requests', timeout=1, delay=0, retries=0, include_regex='', exclude_regex='', ignore_robots=True, history_path=str(Path(directory)/'history.sqlite3'))
            records, _, log = p.crawl_site(**args)
            self.assertEqual(len(records), 1)
            records, _, log = p.crawl_site(**args)
            self.assertEqual(sum(e.status=='error' for e in log), 1)
            records, _, log = p.crawl_site(**args)
            self.assertEqual(len(records), 2)
            records, _, log = p.crawl_site(**args)
            self.assertTrue(all(e.status=='cached' for e in log))
            self.assertEqual(calls.count(ROOT), 1)
            self.assertEqual(calls.count(NEXT), 2)
            self.assertEqual(len(records), 2)

    def test_http_status_is_preserved_and_mime_not_retried(self):
        import requests
        fetcher = p.Fetcher('requests', 1, 0, 0)
        response = requests.Response(); response.status_code=404; response.url=ROOT
        with patch.object(fetcher.session, 'get', return_value=response):
            with self.assertRaises(RuntimeError) as error: fetcher.fetch(ROOT)
            self.assertEqual(error.exception.http_status, 404)
        response.status_code=200; response.headers['content-type']='image/jpeg'
        fetcher.retries=2
        with patch.object(fetcher.session, 'get', return_value=response) as get:
            with self.assertRaises(ValueError): fetcher.fetch(ROOT)
            self.assertEqual(get.call_count, 1)
        fetcher.close()

    def test_source_export_is_not_truncated(self):
        from test_parser_regressions import workbook_rows
        record = p.PlantRecord('x', ROOT, page_text='text'*20000, page_html='<html>original</html>')
        with tempfile.TemporaryDirectory() as directory:
            output = Path(directory)/'test.xlsx'
            p.export_xlsx(str(output), [record], [], [], ROOT)
            rows = workbook_rows(output, 6)[1:]
            self.assertEqual(''.join(row['E'+str(i+2)] for i,row in enumerate(rows) if row.get('C'+str(i+2))=='text'), record.page_text)

if __name__ == '__main__': unittest.main()
