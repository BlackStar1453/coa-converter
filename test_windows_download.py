"""Integration test for the /api/download HTTP endpoint.

Reproduces the "Windows blank download" bug end-to-end: spin up the real HTTP
server on a free port, register a job whose output_path points at a known good
file, fetch the URL, and assert the bytes received match the file on disk.

Runs on every platform — but its primary value is catching Windows-specific
regressions where socket buffering / Content-Length / file locking would
truncate the response. If this test passes on the windows-latest GitHub runner,
the server-side download path is provably correct on Windows.
"""

import hashlib
import http.client
import shutil
import sys
import threading
import time
import unittest
from http.server import HTTPServer
from pathlib import Path

PROJECT_DIR = Path(__file__).parent
sys.path.insert(0, str(PROJECT_DIR))

import app  # noqa: E402

TEMPLATES_DIR = PROJECT_DIR / 'templates'
OUTPUT_DIR = PROJECT_DIR / 'output'


def _md5(data: bytes) -> str:
    return hashlib.md5(data).hexdigest()


class DownloadEndpointTest(unittest.TestCase):
    """Direct HTTP test of /api/download — the actual code path users hit."""

    @classmethod
    def setUpClass(cls):
        # Bind to an ephemeral port so multiple test runs / dev servers don't clash
        cls.server = HTTPServer(('127.0.0.1', 0), app.COAHandler)
        cls.port = cls.server.server_address[1]
        cls.thread = threading.Thread(target=cls.server.serve_forever, daemon=True)
        cls.thread.start()
        # Give the server a moment to start accepting connections
        time.sleep(0.2)

    @classmethod
    def tearDownClass(cls):
        cls.server.shutdown()
        cls.server.server_close()

    def _make_job_with_output(self, source_file: Path, dest_name: str) -> dict:
        """Copy source_file into output/ and register a job referencing it."""
        OUTPUT_DIR.mkdir(exist_ok=True)
        dst = OUTPUT_DIR / dest_name
        shutil.copy2(source_file, dst)

        job = app.jobs.create_job(pdf_name='fake.pdf',
                                  pdf_path=str(PROJECT_DIR / 'fake.pdf'))
        app.jobs.update_job(job['id'],
                            output_path=str(dst),
                            status='done',
                            ai_verified=False)
        return job

    def _fetch(self, path: str):
        conn = http.client.HTTPConnection('127.0.0.1', self.port, timeout=10)
        try:
            conn.request('GET', path)
            resp = conn.getresponse()
            body = resp.read()
            headers = dict(resp.getheaders())
            return resp.status, headers, body
        finally:
            conn.close()

    def test_xlsx_download_byte_exact(self):
        """An XLSX file must download with identical bytes (the blank-on-Windows reproducer)."""
        src = TEMPLATES_DIR / 'Allergen -.xlsx'
        if not src.exists():
            self.skipTest(f'template not found: {src}')

        job = self._make_job_with_output(src, 'test_xlsx_dl.xlsx')
        expected_bytes = (OUTPUT_DIR / 'test_xlsx_dl.xlsx').read_bytes()
        self.assertGreater(len(expected_bytes), 0,
                           'source file must be non-empty for this test to be meaningful')

        status, headers, body = self._fetch(f'/api/download/{job["id"]}')
        self.assertEqual(status, 200)
        self.assertEqual(
            len(body), len(expected_bytes),
            f'downloaded {len(body)}B != on-disk {len(expected_bytes)}B '
            f'(platform={sys.platform})')
        self.assertEqual(_md5(body), _md5(expected_bytes),
                         'downloaded MD5 mismatches on-disk MD5')
        # Header sanity
        self.assertEqual(int(headers.get('Content-Length', '0')), len(expected_bytes))
        self.assertIn('attachment', headers.get('Content-Disposition', ''))

    def test_docx_download_byte_exact(self):
        """Same check for DOCX templates."""
        src = TEMPLATES_DIR / 'Nutrition info -.docx'
        if not src.exists():
            self.skipTest(f'template not found: {src}')

        job = self._make_job_with_output(src, 'test_docx_dl.docx')
        expected_bytes = (OUTPUT_DIR / 'test_docx_dl.docx').read_bytes()

        status, _, body = self._fetch(f'/api/download/{job["id"]}')
        self.assertEqual(status, 200)
        self.assertEqual(_md5(body), _md5(expected_bytes))

    def test_download_404_for_unknown_job(self):
        status, _, _ = self._fetch('/api/download/nonexistent-job-id')
        self.assertEqual(status, 404)

    def test_download_404_when_file_missing(self):
        """If output_path is set but the file was deleted, return 404 not 200-empty."""
        job = app.jobs.create_job(pdf_name='fake.pdf', pdf_path='/tmp/fake.pdf')
        app.jobs.update_job(job['id'],
                            output_path=str(PROJECT_DIR / 'definitely-does-not-exist.xlsx'),
                            status='done')
        status, _, body = self._fetch(f'/api/download/{job["id"]}')
        self.assertEqual(status, 404)
        # Critically: we should never return 200 with a blank body
        self.assertNotEqual((status, len(body)), (200, 0),
                            'should not return 200 with empty body (blank download bug)')


if __name__ == '__main__':
    unittest.main()
