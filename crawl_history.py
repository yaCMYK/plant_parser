"""Durable per-site crawl snapshots. No pickle or executable serialized data."""
import json
from pathlib import Path
import sqlite3
import hashlib

class CrawlHistory:
    def __init__(self, path, site):
        Path(path).parent.mkdir(parents=True, exist_ok=True)
        self.db = sqlite3.connect(path)
        self.site = site
        self.db.execute('CREATE TABLE IF NOT EXISTS pages (site TEXT, url TEXT, state TEXT, payload TEXT, content_hash TEXT, updated TEXT DEFAULT CURRENT_TIMESTAMP, PRIMARY KEY(site,url))')

    def get(self, url):
        row = self.db.execute("SELECT payload FROM pages WHERE site=? AND url=? AND state IN ('completed', 'excluded')", (self.site, url)).fetchone()
        return json.loads(row[0]) if row else None

    def save(self, url, state, payload, html=''):
        self.db.execute('INSERT OR REPLACE INTO pages(site,url,state,payload,content_hash) VALUES(?,?,?,?,?)', (self.site, url, state, json.dumps(payload, ensure_ascii=False), hashlib.sha256(html.encode()).hexdigest()))
        self.db.commit()

    def snapshots(self):
        return [json.loads(row[0]) for row in self.db.execute('SELECT payload FROM pages WHERE site=? AND state=? ORDER BY url', (self.site, 'completed'))]

    def close(self):
        self.db.close()


def restore_record(data):
    from plant_parser_production import PlantRecord, Observation
    record = PlantRecord(data['plant_id'], data['page_url'], data['page_title'], page_text=data.get('page_text', ''), page_html=data.get('page_html', ''))
    for key, observations in data['observations'].items():
        record.observations[key] = [Observation(**o) for o in observations]
    record.unclassified_images = data['unclassified_images']
    record.additional_sources = data.get('additional_sources', [])
    return record
