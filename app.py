#!/usr/bin/env python3
"""COA Converter Web — local web app for COA PDF to template conversion."""

import json
import os
import logging
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs
from pathlib import Path

from job_manager import JobManager
from converter_service import run_conversion
from terminal_launcher import launch_verification, launch_verification_silent, focus_terminal

# --- Config ---
PORT = 5050
HOST = '0.0.0.0'
PROJECT_DIR = Path(__file__).parent
INPUT_DIR = PROJECT_DIR / 'input'
OUTPUT_DIR = PROJECT_DIR / 'output'
TEMPLATES_DIR = PROJECT_DIR / 'templates'
STATIC_DIR = PROJECT_DIR / 'static'

logging.basicConfig(level=logging.INFO,
                    format='[COA-Web] %(levelname)s: %(message)s')
logger = logging.getLogger(__name__)

# Ensure directories exist
INPUT_DIR.mkdir(parents=True, exist_ok=True)
OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

# Global job manager
jobs = JobManager()

# Batch processing lock
_batch_lock = threading.Lock()


def _json_response(handler, data, status=200):
    """Send a JSON response."""
    body = json.dumps(data, ensure_ascii=False).encode('utf-8')
    handler.send_response(status)
    handler.send_header('Content-Type', 'application/json; charset=utf-8')
    handler.send_header('Content-Length', len(body))
    handler.end_headers()
    handler.wfile.write(body)


def _read_body(handler) -> bytes:
    """Read request body."""
    length = int(handler.headers.get('Content-Length', 0))
    return handler.rfile.read(length) if length else b''


def _parse_multipart(handler) -> dict:
    """Parse multipart/form-data for file uploads. Returns {filename: bytes}."""
    content_type = handler.headers.get('Content-Type', '')
    if 'boundary=' not in content_type:
        return {}

    boundary = content_type.split('boundary=')[1].strip()
    body = _read_body(handler)

    files = {}
    fields = {}
    parts = body.split(f'--{boundary}'.encode())

    for part in parts:
        if not part or part == b'--\r\n' or part == b'--':
            continue

        if b'\r\n\r\n' not in part:
            continue

        header_data, file_data = part.split(b'\r\n\r\n', 1)
        header_str = header_data.decode('utf-8', errors='replace')

        # Remove trailing \r\n from file_data
        if file_data.endswith(b'\r\n'):
            file_data = file_data[:-2]

        if 'filename="' in header_str:
            # Extract filename
            for h in header_str.split('\r\n'):
                if 'filename="' in h:
                    fname = h.split('filename="')[1].split('"')[0]
                    if fname:
                        files[fname] = file_data
        elif 'name="' in header_str:
            # Extract form field
            for h in header_str.split('\r\n'):
                if 'name="' in h:
                    name = h.split('name="')[1].split('"')[0]
                    fields[name] = file_data.decode('utf-8', errors='replace')

    return {'files': files, 'fields': fields}


class COAHandler(SimpleHTTPRequestHandler):
    """HTTP handler with API routes and static file serving."""

    def log_message(self, format, *args):
        logger.info(f'{self.address_string()} - {format % args}')

    # --- Routing ---

    def do_GET(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == '/' or path == '/index.html':
            self._serve_file(STATIC_DIR / 'index.html', 'text/html')
        elif path == '/style.css':
            self._serve_file(STATIC_DIR / 'style.css', 'text/css')
        elif path == '/app.js':
            self._serve_file(STATIC_DIR / 'app.js', 'application/javascript')
        elif path == '/api/templates':
            self._handle_templates()
        elif path == '/api/jobs':
            self._handle_list_jobs()
        elif path.startswith('/api/jobs/'):
            job_id = path.split('/api/jobs/')[1]
            self._handle_get_job(job_id)
        elif path.startswith('/api/download/'):
            job_id = path.split('/api/download/')[1]
            self._handle_download(job_id)
        else:
            self.send_error(404)

    def do_POST(self):
        parsed = urlparse(self.path)
        path = parsed.path

        if path == '/api/upload':
            self._handle_upload()
        elif path.startswith('/api/convert/'):
            job_id = path.split('/api/convert/')[1]
            self._handle_convert(job_id)
        elif path == '/api/convert-all':
            self._handle_convert_all()
        elif path.startswith('/api/verify/'):
            job_id = path.split('/api/verify/')[1]
            self._handle_verify(job_id)
        elif path == '/api/focus-terminal':
            focus_terminal()
            _json_response(self, {'ok': True})
        else:
            self.send_error(404)

    def do_DELETE(self):
        parsed = urlparse(self.path)
        path = parsed.path
        if path.startswith('/api/jobs/'):
            job_id = path.split('/api/jobs/')[1]
            if jobs.delete_job(job_id):
                _json_response(self, {'ok': True})
            else:
                _json_response(self, {'error': 'Job not found'}, 404)
        else:
            self.send_error(404)

    # --- Handlers ---

    def _serve_file(self, filepath: Path, content_type: str):
        if not filepath.exists():
            self.send_error(404)
            return
        data = filepath.read_bytes()
        self.send_response(200)
        self.send_header('Content-Type', f'{content_type}; charset=utf-8')
        self.send_header('Content-Length', len(data))
        self.end_headers()
        self.wfile.write(data)

    def _handle_templates(self):
        templates = []
        if TEMPLATES_DIR.exists():
            for f in sorted(TEMPLATES_DIR.iterdir()):
                if f.suffix.lower() in ('.xlsx', '.docx'):
                    templates.append({
                        'name': f.name,
                        'path': str(f),
                        'format': f.suffix.lower()[1:],
                    })
        _json_response(self, templates)

    def _handle_list_jobs(self):
        _json_response(self, jobs.get_all_jobs())

    def _handle_get_job(self, job_id):
        job = jobs.get_job(job_id)
        if job:
            _json_response(self, job)
        else:
            _json_response(self, {'error': 'Job not found'}, 404)

    def _handle_upload(self):
        result = _parse_multipart(self)
        files = result.get('files', {})
        if not files:
            _json_response(self, {'error': 'No files uploaded'}, 400)
            return

        created = []
        for fname, data in files.items():
            if not fname.lower().endswith('.pdf'):
                continue
            # Save to input directory
            save_path = INPUT_DIR / fname
            # Avoid overwriting
            counter = 1
            while save_path.exists():
                stem = Path(fname).stem
                save_path = INPUT_DIR / f'{stem}_{counter}.pdf'
                counter += 1
            save_path.write_bytes(data)
            job = jobs.create_job(pdf_name=save_path.name, pdf_path=str(save_path))
            created.append(job)

        if not created:
            _json_response(self, {'error': 'No valid PDF files'}, 400)
        else:
            _json_response(self, created, 201)

    def _handle_convert(self, job_id):
        body = _read_body(self)
        try:
            params = json.loads(body) if body else {}
        except json.JSONDecodeError:
            params = {}

        template_path = params.get('template_path')
        force_verify = params.get('force_verify', False)

        if not template_path:
            _json_response(self, {'error': 'template_path required'}, 400)
            return

        job = jobs.get_job(job_id)
        if not job:
            _json_response(self, {'error': 'Job not found'}, 404)
            return
        if job['status'] != 'pending':
            _json_response(self, {'error': f'Job is {job["status"]}, not pending'}, 400)
            return

        template_name = os.path.basename(template_path)
        ext = os.path.splitext(template_path)[1]
        pdf_stem = Path(job['pdf_path']).stem
        output_path = str(OUTPUT_DIR / f'{pdf_stem}{ext}')

        # Avoid overwriting output
        counter = 1
        while os.path.exists(output_path):
            output_path = str(OUTPUT_DIR / f'{pdf_stem}_{counter}{ext}')
            counter += 1

        claude_mode = params.get('claude_mode', 'silent')

        jobs.update_job(job_id, template_name=template_name,
                        template_path=template_path,
                        force_verify=force_verify)

        def on_verify_needed(jid, pdf, tpl, out):
            if claude_mode == 'interactive':
                launch_verification(jobs, jid, pdf, tpl, out)
            else:
                launch_verification_silent(jobs, jid, pdf, tpl, out)

        run_conversion(jobs, job_id, job['pdf_path'], template_path,
                       output_path, on_complete=on_verify_needed)

        _json_response(self, {'ok': True, 'job_id': job_id})

    def _handle_convert_all(self):
        body = _read_body(self)
        try:
            params = json.loads(body) if body else {}
        except json.JSONDecodeError:
            params = {}

        template_path = params.get('template_path')
        force_verify = params.get('force_verify', False)
        claude_mode = params.get('claude_mode', 'silent')

        if not template_path:
            _json_response(self, {'error': 'template_path required'}, 400)
            return

        pending = jobs.get_pending_jobs()
        if not pending:
            _json_response(self, {'error': 'No pending jobs'}, 400)
            return

        def _batch():
            with _batch_lock:
                for job in pending:
                    jid = job['id']
                    template_name = os.path.basename(template_path)
                    ext = os.path.splitext(template_path)[1]
                    pdf_stem = Path(job['pdf_path']).stem
                    output_path = str(OUTPUT_DIR / f'{pdf_stem}{ext}')
                    counter = 1
                    while os.path.exists(output_path):
                        output_path = str(OUTPUT_DIR / f'{pdf_stem}_{counter}{ext}')
                        counter += 1

                    jobs.update_job(jid, template_name=template_name,
                                    template_path=template_path,
                                    force_verify=force_verify)

                    def on_verify(j=jid, p=job['pdf_path'],
                                  t=template_path, o=output_path):
                        def _cb(_jid, _pdf, _tpl, _out):
                            if claude_mode == 'interactive':
                                launch_verification(jobs, j, p, t, o)
                            else:
                                launch_verification_silent(jobs, j, p, t, o)
                        return _cb

                    t = run_conversion(jobs, jid, job['pdf_path'],
                                       template_path, output_path,
                                       on_complete=on_verify())
                    t.join()  # Sequential processing

        threading.Thread(target=_batch, daemon=True).start()
        _json_response(self, {'ok': True, 'count': len(pending)})

    def _handle_verify(self, job_id):
        body = _read_body(self)
        try:
            params = json.loads(body) if body else {}
        except json.JSONDecodeError:
            params = {}
        claude_mode = params.get('claude_mode', 'silent')

        job = jobs.get_job(job_id)
        if not job:
            _json_response(self, {'error': 'Job not found'}, 404)
            return
        if not job.get('output_path'):
            _json_response(self, {'error': 'No output file yet'}, 400)
            return

        jobs.update_job(job_id, status='verifying')
        if claude_mode == 'interactive':
            launch_verification(jobs, job_id, job['pdf_path'],
                                job['template_path'], job['output_path'])
        else:
            launch_verification_silent(jobs, job_id, job['pdf_path'],
                                       job['template_path'], job['output_path'])
        _json_response(self, {'ok': True})

    def _handle_download(self, job_id):
        job = jobs.get_job(job_id)
        if not job:
            self.send_error(404)
            return
        output_path = job.get('output_path')
        if not output_path or not os.path.exists(output_path):
            self.send_error(404)
            return

        filename = os.path.basename(output_path)
        data = open(output_path, 'rb').read()
        self.send_response(200)
        self.send_header('Content-Type', 'application/octet-stream')
        self.send_header('Content-Disposition',
                         f'attachment; filename="{filename}"')
        self.send_header('Content-Length', len(data))
        self.end_headers()
        self.wfile.write(data)


def main():
    server = HTTPServer((HOST, PORT), COAHandler)
    logger.info(f'COA Converter Web running at http://{HOST}:{PORT}')
    try:
        server.serve_forever()
    except KeyboardInterrupt:
        logger.info('Shutting down...')
        server.shutdown()


if __name__ == '__main__':
    main()
