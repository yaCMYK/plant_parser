"""Authenticated Railway UI; run with exactly one Gunicorn worker."""
import hmac
import json
import logging
import os
from pathlib import Path
import re
import threading
from urllib.parse import urlparse
import uuid

from flask import Flask, jsonify, render_template, request, send_file

import plant_parser_production as parser


def create_app(data_dir=None, password=None, crawl=None):
    app = Flask(__name__)
    app.config['MAX_CONTENT_LENGTH'] = 8192
    storage = Path(data_dir or os.environ.get('DATA_DIR', '/data')) / 'jobs'
    storage.mkdir(parents=True, exist_ok=True)
    password = password if password is not None else os.environ.get('WEB_PASSWORD', '')
    crawl = crawl or parser.crawl_site
    busy = threading.Lock()
    logging.basicConfig(level=logging.INFO)

    def write_status(folder, payload):
        temporary = folder / 'status.tmp'
        temporary.write_text(json.dumps(payload, ensure_ascii=False), encoding='utf-8')
        temporary.replace(folder / 'status.json')

    # Processes do not survive deployments. Keep finished reports, mark interrupted jobs.
    for status_path in storage.glob('*/status.json'):
        state = json.loads(status_path.read_text(encoding='utf-8'))
        if state.get('status') == 'running':
            state.update(status='failed', error='Сервис перезапущен. Запустите задание повторно.')
            write_status(status_path.parent, state)

    @app.before_request
    def authenticate():
        if request.path == '/health':
            return None
        if not password:
            return jsonify(error='Задайте WEB_PASSWORD в настройках сервиса.'), 503
        auth = request.authorization
        if not auth or auth.type.lower() != 'basic' or not hmac.compare_digest(
            (auth.username or '').encode(), os.environ.get('WEB_USERNAME', 'admin').encode()
        ) or not hmac.compare_digest((auth.password or '').encode(), password.encode()):
            return ('Требуется вход', 401, {'WWW-Authenticate': 'Basic realm="Plant parser", charset="UTF-8"'})

    @app.get('/health')
    def health():
        return jsonify(status='ok' if password else 'missing_password'), 200 if password else 503

    @app.get('/')
    def index():
        return render_template('index.html')

    def run_job(job_id, folder, url, max_pages, display_unit):
        state = {'id': job_id, 'status': 'running', 'url': url}
        try:
            records, rejected, log = crawl(
                root_url=url, max_pages=max_pages,
                browser=os.environ.get('PARSER_BROWSER', 'requests'),
                timeout=25, delay=0.5, retries=2,
                include_regex='', exclude_regex=r'/(cart|checkout|login|account|privacy|terms)(/|$)',
                ignore_robots=False,
                history_path=str(storage.parent / "crawl.sqlite3"),
            )
            errors = sum(entry.status == 'error' for entry in log)
            parser.export_xlsx(str(folder / 'plants.tmp.xlsx'), records, rejected, log, url, display_unit)
            (folder / 'plants.tmp.xlsx').replace(folder / 'plants.xlsx')
            meaningful = sum(any(o.accepted_to_main and key not in {'source_page_url', 'source_image_url', 'scrape_date', 'comment'} for key, values in r.observations.items() for o in values) for r in records)
            state.update(status='completed' if meaningful else 'failed', plants=len(records), pages=len(log), errors=errors,
                         downloaded=sum(e.status == 'ok' for e in log), cached=sum(e.status == 'cached' for e in log),
                         warning='Частичный результат: проверьте Crawl log.' if errors else '',
                         error='' if meaningful else 'Содержательные данные не найдены. Скачайте диагностический XLSX.',
                         download=f'/api/jobs/{job_id}/download')
        except Exception:
            app.logger.exception('Job %s failed', job_id)
            state.update(status='failed', error='Не удалось собрать растения. Проверьте URL, доступность сайта и журнал сервиса.')
        finally:
            try:
                write_status(folder, state)
            finally:
                busy.release()

    @app.post('/api/jobs')
    def start_job():
        data = request.get_json(silent=True)
        if not isinstance(data, dict):
            return jsonify(error='Ожидается JSON-объект.'), 400
        url = data.get('url', '')
        if not isinstance(url, str):
            return jsonify(error='URL должен быть строкой.'), 400
        url = url.strip()
        try:
            parsed = urlparse(url)
            if parsed.scheme not in {'http', 'https'} or not parsed.hostname or parsed.username or parsed.password:
                raise ValueError()
            pages = int(data.get('max_pages', 50))
            if not 1 <= pages <= 500:
                raise ValueError()
            unit = data.get('display_unit', 'm')
            if unit not in {'m', 'cm', 'mm'}:
                raise ValueError()
        except (ValueError, TypeError):
            return jsonify(error='Укажите HTTP(S) URL без логина, от 1 до 500 страниц и единицы m/cm/mm.'), 400
        if not busy.acquire(blocking=False):
            return jsonify(error='Парсер уже выполняет задание. Дождитесь завершения.'), 409
        job_id = uuid.uuid4().hex
        folder = storage / job_id
        try:
            folder.mkdir()
            state = {'id': job_id, 'status': 'running', 'url': url}
            write_status(folder, state)
            threading.Thread(target=run_job, args=(job_id, folder, url, pages, unit), daemon=True).start()
        except Exception:
            busy.release()
            raise
        return jsonify(state), 202

    def job_folder(job_id):
        if not re.fullmatch(r'[0-9a-f]{32}', job_id):
            return None
        folder = storage / job_id
        return folder if (folder / 'status.json').is_file() else None

    @app.get('/api/jobs/<job_id>')
    def status(job_id):
        folder = job_folder(job_id)
        if folder is None:
            return jsonify(error='Задание не найдено.'), 404
        return jsonify(json.loads((folder / 'status.json').read_text(encoding='utf-8')))

    @app.get('/api/jobs/<job_id>/download')
    def download(job_id):
        folder = job_folder(job_id)
        if folder is None or not (folder / 'plants.xlsx').is_file():
            return jsonify(error='Файл ещё не готов или задание не найдено.'), 404
        return send_file(folder / 'plants.xlsx', as_attachment=True, download_name=f'plants-{job_id}.xlsx')

    return app

