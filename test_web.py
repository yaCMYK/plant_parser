import base64
import io
import json
from pathlib import Path
import tempfile
import threading
import time
import unittest
import zipfile

from bs4 import BeautifulSoup
import plant_parser_production as parser
from web import create_app


class WebTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.addCleanup(self.temp.cleanup)
        self.headers = {'Authorization': 'Basic ' + base64.b64encode(b'admin:test-password').decode()}

    def client(self, crawl=None, password='test-password'):
        return create_app(self.temp.name, password=password, crawl=crawl).test_client()

    def wait_done(self, client, job_id):
        for _ in range(100):
            state = client.get('/api/jobs/' + job_id, headers=self.headers).get_json()
            if state['status'] != 'running':
                return state
            time.sleep(.01)
        self.fail('Job did not finish')

    def test_health_and_authentication(self):
        client = self.client()
        self.assertEqual(client.get('/health').status_code, 200)
        self.assertEqual(client.get('/').status_code, 401)
        self.assertEqual(client.get('/', headers=self.headers).status_code, 200)
        self.assertEqual(self.client(password='').get('/health').status_code, 503)
        self.assertEqual(self.client(password='').get('/').status_code, 503)

    def test_validation(self):
        client = self.client()
        for payload in [[], {}, {'url': 'file:///etc/passwd'}, {'url': 'https://example.org', 'max_pages': 501}, {'url': 'https://user:pass@example.org'}, {'url': 'https://example.org', 'display_unit': 'ft'}]:
            with self.subTest(payload=payload):
                self.assertEqual(client.post('/api/jobs', json=payload, headers=self.headers).status_code, 400)
        self.assertEqual(client.get('/api/jobs/invalid', headers=self.headers).status_code, 404)

    def test_job_download_and_persistence(self):
        def crawl(**kwargs):
            soup = BeautifulSoup('<h1>Pinus sylvestris</h1><p>Высота во взрослом виде: 150–200 cm</p>', 'lxml')
            record, rejected = parser.build_plant_record(kwargs['root_url'], soup, 'now')
            return [record], rejected, [parser.CrawlEntry(kwargs['root_url'], 'ok', 200)]
        client = self.client(crawl)
        response = client.post('/api/jobs', json={'url': 'https://example.org/plants', 'max_pages': 5}, headers=self.headers)
        self.assertEqual(response.status_code, 202)
        job_id = response.get_json()['id']
        state = self.wait_done(client, job_id)
        self.assertEqual(state['status'], 'completed')
        self.assertEqual(state['plants'], 1)
        response = client.get(state['download'], headers=self.headers)
        self.assertEqual(response.status_code, 200)
        with zipfile.ZipFile(io.BytesIO(response.data)) as archive:
            self.assertIsNone(archive.testzip())
        response.close()
        restarted = self.client(crawl)
        self.assertEqual(restarted.get('/api/jobs/' + job_id, headers=self.headers).get_json()['status'], 'completed')

    def test_only_one_job_runs_and_failure_is_visible(self):
        started, release = threading.Event(), threading.Event()
        def crawl(**kwargs):
            started.set()
            release.wait(3)
            return [], [], []
        client = self.client(crawl)
        response = client.post('/api/jobs', json={'url': 'https://example.org'}, headers=self.headers)
        job_id = response.get_json()['id']
        self.assertTrue(started.wait(1))
        try:
            self.assertEqual(client.post('/api/jobs', json={'url': 'https://example.org'}, headers=self.headers).status_code, 409)
            self.assertEqual(client.get(f'/api/jobs/{job_id}/download', headers=self.headers).status_code, 404)
        finally:
            release.set()
        self.assertEqual(self.wait_done(client, job_id)['status'], 'failed')
        state = client.get('/api/jobs/' + job_id, headers=self.headers).get_json()
        self.assertIn('download', state)
        response = client.get(state['download'], headers=self.headers)
        self.assertEqual(response.status_code, 200)
        response.close()

    def test_interrupted_job_is_marked_failed(self):
        folder = Path(self.temp.name) / 'jobs' / ('a' * 32)
        folder.mkdir(parents=True)
        (folder / 'status.json').write_text(json.dumps({'id': 'a' * 32, 'status': 'running'}))
        client = self.client()
        state = client.get('/api/jobs/' + 'a' * 32, headers=self.headers).get_json()
        self.assertEqual(state['status'], 'failed')


if __name__ == '__main__':
    unittest.main()
