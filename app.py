#!/usr/bin/env python3
"""COA Converter Web — local web app for COA PDF to template conversion."""

import json
import os
import logging
import sys
import threading
from http.server import HTTPServer, SimpleHTTPRequestHandler
from urllib.parse import urlparse, parse_qs, quote
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
SAMPLES_DIR = PROJECT_DIR / 'samples'

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


_LOCALHOST_IPS = {'127.0.0.1', '::1'}

# Interactive AI verification (visible Terminal window with live Claude output)
# currently relies on macOS Terminal.app + osascript. On other platforms we
# transparently force claude_mode='silent'. The silent path itself is fully
# cross-platform (just subprocess.run claude -p).
_INTERACTIVE_ALLOWED = sys.platform == 'darwin'


class COAHandler(SimpleHTTPRequestHandler):
    """HTTP handler with API routes and static file serving."""

    def _is_local(self) -> bool:
        """Return True if the request comes from localhost."""
        return self.client_address[0] in _LOCALHOST_IPS

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
        elif path == '/api/client-info':
            _json_response(self, {
                'is_local': self._is_local(),
                'platform': sys.platform,
                'interactive_allowed': _INTERACTIVE_ALLOWED,
            })
        elif path.startswith('/api/download/'):
            job_id = path.split('/api/download/')[1]
            self._handle_download(job_id)
        elif path.startswith('/api/sample/'):
            sample_name = path.split('/api/sample/')[1]
            self._handle_sample_download(sample_name)
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
        elif path.startswith('/api/report-error/'):
            job_id = path.split('/api/report-error/')[1]
            self._handle_report_error(job_id)
        elif path.startswith('/api/remove/'):
            job_id = path.split('/api/remove/')[1]
            self._handle_remove(job_id)
        elif path == '/api/focus-terminal':
            focus_terminal()
            _json_response(self, {'ok': True})
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
        if not self._is_local() or not _INTERACTIVE_ALLOWED:
            claude_mode = 'silent'

        jobs.update_job(job_id, template_name=template_name,
                        template_path=template_path)

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

        # Support both array (template_paths) and single (template_path)
        template_paths = params.get('template_paths', [])
        if not template_paths:
            tp = params.get('template_path')
            if tp:
                template_paths = [tp]
        if not template_paths:
            _json_response(self, {'error': 'template_paths required'}, 400)
            return

        claude_mode = params.get('claude_mode', 'silent')
        if not self._is_local() or not _INTERACTIVE_ALLOWED:
            claude_mode = 'silent'

        pending = jobs.get_pending_jobs()
        if not pending:
            _json_response(self, {'error': 'No pending jobs'}, 400)
            return

        # Build (job_id, pdf_path, template_path) work items
        work_items = []
        for job in pending:
            for i, tpl_path in enumerate(template_paths):
                if i == 0:
                    work_items.append((job['id'], job['pdf_path'], tpl_path))
                else:
                    new_job = jobs.create_job(pdf_name=job['pdf_name'],
                                              pdf_path=job['pdf_path'])
                    work_items.append((new_job['id'], job['pdf_path'], tpl_path))

        def _batch():
            with _batch_lock:
                threads = []
                for jid, pdf_path, tpl_path in work_items:
                    template_name = os.path.basename(tpl_path)
                    tpl_stem = Path(template_name).stem
                    ext = os.path.splitext(tpl_path)[1]
                    pdf_stem = Path(pdf_path).stem
                    output_path = str(OUTPUT_DIR / f'{pdf_stem}_{tpl_stem}{ext}')
                    counter = 1
                    while os.path.exists(output_path):
                        output_path = str(OUTPUT_DIR / f'{pdf_stem}_{tpl_stem}_{counter}{ext}')
                        counter += 1

                    jobs.update_job(jid, template_name=template_name,
                                    template_path=tpl_path)

                    def on_verify(j=jid, p=pdf_path, t=tpl_path, o=output_path):
                        def _cb(_jid, _pdf, _tpl, _out):
                            if claude_mode == 'interactive':
                                launch_verification(jobs, j, p, t, o)
                            else:
                                launch_verification_silent(jobs, j, p, t, o)
                        return _cb

                    t = run_conversion(jobs, jid, pdf_path, tpl_path,
                                       output_path, on_complete=on_verify())
                    threads.append(t)

                # Wait for all conversions (parallel, not sequential)
                for t in threads:
                    t.join()

        threading.Thread(target=_batch, daemon=True).start()
        _json_response(self, {'ok': True, 'count': len(work_items)})

    def _handle_verify(self, job_id):
        body = _read_body(self)
        try:
            params = json.loads(body) if body else {}
        except json.JSONDecodeError:
            params = {}
        claude_mode = params.get('claude_mode', 'silent')
        if not self._is_local() or not _INTERACTIVE_ALLOWED:
            claude_mode = 'silent'

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

    def _handle_report_error(self, job_id):
        body = _read_body(self)
        try:
            params = json.loads(body) if body else {}
        except json.JSONDecodeError:
            params = {}

        message = params.get('message', '').strip()
        if not message:
            _json_response(self, {'error': 'Error message required'}, 400)
            return

        job = jobs.get_job(job_id)
        if not job:
            _json_response(self, {'error': 'Job not found'}, 404)
            return
        if not job.get('output_path'):
            _json_response(self, {'error': 'No output file yet'}, 400)
            return

        claude_mode = params.get('claude_mode', 'silent')
        if not self._is_local() or not _INTERACTIVE_ALLOWED:
            claude_mode = 'silent'

        jobs.update_job(job_id, status='verifying')

        from terminal_launcher import launch_error_fix_silent, launch_error_fix
        if claude_mode == 'interactive':
            launch_error_fix(jobs, job_id, job['pdf_path'],
                             job['template_path'], job['output_path'], message)
        else:
            launch_error_fix_silent(jobs, job_id, job['pdf_path'],
                                    job['template_path'], job['output_path'],
                                    message)
        _json_response(self, {'ok': True})

    def _handle_remove(self, job_id):
        """Delete job and clean up files on disk."""
        job = jobs.get_job(job_id)
        if not job:
            _json_response(self, {'error': 'Job not found'}, 404)
            return
        # Remove input PDF
        pdf_path = job.get('pdf_path')
        if pdf_path and os.path.exists(pdf_path):
            try:
                os.remove(pdf_path)
            except OSError:
                pass
        # Remove output file
        output_path = job.get('output_path')
        if output_path and os.path.exists(output_path):
            try:
                os.remove(output_path)
            except OSError:
                pass
        jobs.delete_job(job_id)
        _json_response(self, {'ok': True})

    def _handle_sample_download(self, sample_name):
        """Diagnostic endpoint: download a pre-baked sample file directly.

        Bypasses the upload + convert + AI-verify pipeline so we can isolate
        whether the HTTP download layer itself is healthy on a given client
        (e.g. Windows browser over LAN). Uses the same byte-writing code path
        as the real /api/download/<job_id> handler, including the same
        diagnostic logs prefixed [下载-sample].

        Allowed names are restricted to files under samples/ to prevent any
        kind of path traversal.
        """
        # Normalize and validate — only allow plain filenames inside SAMPLES_DIR
        safe_name = os.path.basename(sample_name)
        if not safe_name or safe_name != sample_name:
            logger.warning(f'[下载-sample] 非法 sample 名称: {sample_name!r}')
            self.send_error(400)
            return

        sample_path = SAMPLES_DIR / safe_name
        if not sample_path.exists() or not sample_path.is_file():
            logger.warning(f'[下载-sample] 文件不存在: {sample_path}')
            self.send_error(404)
            return

        try:
            file_size = sample_path.stat().st_size
        except OSError as e:
            logger.error(f'[下载-sample] stat 失败 {sample_path}: {e}')
            self.send_error(500)
            return

        logger.info(f'[下载-sample] name={safe_name} path={sample_path} size={file_size}B')

        try:
            with open(sample_path, 'rb') as f:
                data = f.read()
        except OSError as e:
            logger.error(f'[下载-sample] 读取失败 {sample_path}: {e}')
            self.send_error(500)
            return

        actual_bytes = len(data)
        self.send_response(200)
        self.send_header('Content-Type', 'application/octet-stream')
        ascii_name = safe_name.encode('ascii', 'replace').decode('ascii')
        self.send_header('Content-Disposition',
                         f"attachment; filename=\"{ascii_name}\"; "
                         f"filename*=UTF-8''{quote(safe_name)}")
        self.send_header('Content-Length', str(actual_bytes))
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Connection', 'close')
        self.end_headers()
        try:
            self.wfile.write(data)
            self.wfile.flush()
            logger.info(f'[下载-sample] {safe_name} 写入 {actual_bytes}B 完成')
        except (BrokenPipeError, ConnectionResetError) as e:
            logger.warning(f'[下载-sample] 客户端断开连接 {safe_name}: {e}')

    def _handle_download(self, job_id):
        job = jobs.get_job(job_id)
        if not job:
            logger.warning(f'[下载] job {job_id} 不存在')
            self.send_error(404)
            return
        output_path = job.get('output_path')
        if not output_path or not os.path.exists(output_path):
            logger.warning(f'[下载] job {job_id} 文件不存在: {output_path}')
            self.send_error(404)
            return

        filename = os.path.basename(output_path)
        try:
            file_size = os.path.getsize(output_path)
        except OSError as e:
            logger.error(f'[下载] 获取文件大小失败 {output_path}: {e}')
            self.send_error(500)
            return

        logger.info(f'[下载] job={job_id} path={output_path} size={file_size}B')

        if file_size == 0:
            # Empty file on disk — log loudly so Windows-debug users can see it
            logger.error(f'[下载] 警告：磁盘上的文件大小为 0！job={job_id} path={output_path}')

        try:
            with open(output_path, 'rb') as f:
                data = f.read()
        except OSError as e:
            # On Windows, the file may be locked by Excel/Word if user opened it
            logger.error(f'[下载] 读取文件失败 {output_path}: {e}')
            self.send_error(500)
            return

        actual_bytes = len(data)
        if actual_bytes != file_size:
            logger.warning(
                f'[下载] 读取字节数 ({actual_bytes}) 与文件大小 ({file_size}) 不一致 '
                f'(平台={sys.platform})')

        self.send_response(200)
        self.send_header('Content-Type', 'application/octet-stream')
        # RFC 6266: ASCII fallback + UTF-8 encoded filename for special chars
        ascii_name = filename.encode('ascii', 'replace').decode('ascii')
        self.send_header('Content-Disposition',
                         f"attachment; filename=\"{ascii_name}\"; "
                         f"filename*=UTF-8''{quote(filename)}")
        self.send_header('Content-Length', str(actual_bytes))
        self.send_header('Cache-Control', 'no-cache, no-store, must-revalidate')
        self.send_header('Connection', 'close')
        self.end_headers()
        try:
            self.wfile.write(data)
            self.wfile.flush()
            logger.info(f'[下载] job={job_id} 写入 {actual_bytes}B 完成')
        except (BrokenPipeError, ConnectionResetError) as e:
            logger.warning(f'[下载] 客户端断开连接 job={job_id}: {e}')


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
