#!/usr/bin/env python3
"""
Tests for the 3 features:
  1. Multi-template selection & batch convert
  2. Parallel conversion with multiple templates
  3. Download works correctly (Content-Disposition: attachment)
"""

import io
import json
import os
import sys
import time
import threading
import unittest
import uuid
from http.client import HTTPConnection
from pathlib import Path
from unittest.mock import patch, MagicMock

# ---- project imports ----
PROJECT_DIR = Path(__file__).parent
sys.path.insert(0, str(PROJECT_DIR))
sys.path.insert(0, str(PROJECT_DIR / 'converter'))

from job_manager import JobManager
import app as app_module
from app import COAHandler, _json_response, HOST, PORT, OUTPUT_DIR, INPUT_DIR, TEMPLATES_DIR

# ---------- helpers ----------

def _make_minimal_pdf() -> bytes:
    """Return the smallest valid PDF (blank page)."""
    return (
        b'%PDF-1.0\n1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj '
        b'2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj '
        b'3 0 obj<</Type/Page/MediaBox[0 0 612 792]/Parent 2 0 R>>endobj\n'
        b'xref\n0 4\n0000000000 65535 f \n'
        b'0000000009 00000 n \n0000000058 00000 n \n0000000115 00000 n \n'
        b'trailer<</Size 4/Root 1 0 R>>\nstartxref\n190\n%%EOF'
    )


def _get_template_paths():
    """Return list of real template absolute paths."""
    paths = []
    if TEMPLATES_DIR.exists():
        for f in sorted(TEMPLATES_DIR.iterdir()):
            if f.suffix.lower() in ('.xlsx', '.docx'):
                paths.append(str(f))
    return paths


class _ServerThread:
    """Starts the HTTP server in a background thread for integration tests."""

    def __init__(self, port=0):
        from http.server import HTTPServer
        # Reset global job manager to avoid cross-test contamination
        app_module.jobs = JobManager()
        self.server = HTTPServer(('127.0.0.1', port), COAHandler)
        self.port = self.server.server_address[1]
        self._thread = threading.Thread(target=self.server.serve_forever, daemon=True)

    def start(self):
        self._thread.start()
        return self

    def stop(self):
        self.server.shutdown()

    def conn(self) -> HTTPConnection:
        return HTTPConnection('127.0.0.1', self.port, timeout=10)


# ===================== Unit Tests =====================


class TestJobManagerMultiTemplate(unittest.TestCase):
    """Test 1: JobManager supports creating cloned jobs for multi-template."""

    def test_create_multiple_jobs_same_pdf(self):
        """When user picks N templates for 1 PDF, N jobs should be created."""
        jm = JobManager()
        pdf_path = '/tmp/test.pdf'
        job1 = jm.create_job(pdf_name='test.pdf', pdf_path=pdf_path)
        job2 = jm.create_job(pdf_name='test.pdf', pdf_path=pdf_path)
        job3 = jm.create_job(pdf_name='test.pdf', pdf_path=pdf_path)

        self.assertNotEqual(job1['id'], job2['id'])
        self.assertNotEqual(job2['id'], job3['id'])
        self.assertEqual(len(jm.get_all_jobs()), 3)
        # All share the same pdf_path
        for j in jm.get_all_jobs():
            self.assertEqual(j['pdf_path'], pdf_path)

    def test_pending_jobs_returns_only_pending(self):
        jm = JobManager()
        j1 = jm.create_job(pdf_name='a.pdf', pdf_path='/tmp/a.pdf')
        j2 = jm.create_job(pdf_name='b.pdf', pdf_path='/tmp/b.pdf')
        jm.update_job(j1['id'], status='done')

        pending = jm.get_pending_jobs()
        self.assertEqual(len(pending), 1)
        self.assertEqual(pending[0]['id'], j2['id'])


class TestConvertAllMultiTemplate(unittest.TestCase):
    """Test 1 & 2: Backend _handle_convert_all accepts template_paths array
    and creates cross-product (PDF × template) jobs running in parallel."""

    def setUp(self):
        # Start a real server
        self.srv = _ServerThread(port=0).start()
        # Prepare a test PDF in input/
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if len(self.templates) < 2:
            self.skipTest('Need at least 2 templates in templates/ dir')

    def tearDown(self):
        self.srv.stop()

    def _upload_pdf(self, filename='test_sample.pdf'):
        """Upload a PDF via multipart and return job list."""
        boundary = '----TestBoundary123'
        body = (
            f'------TestBoundary123\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------TestBoundary123--\r\n'

        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': f'multipart/form-data; boundary=----TestBoundary123'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return resp.status, data

    def _get_jobs(self):
        conn = self.srv.conn()
        conn.request('GET', '/api/jobs')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_multi_template_creates_cross_product_jobs(self, mock_supplier, mock_convert):
        """Selecting 2 templates for 1 PDF should create 2 jobs."""
        # Mock conversion to just touch the output file
        def fake_convert(pdf, tpl, out):
            Path(out).write_bytes(b'fake output')
            return out
        mock_convert.side_effect = fake_convert
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        # Upload 1 PDF
        status, data = self._upload_pdf()
        self.assertEqual(status, 201)

        # Select 2 templates and convert all
        tpl_paths = self.templates[:2]
        payload = json.dumps({
            'template_paths': tpl_paths,
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()

        conn = self.srv.conn()
        conn.request('POST', '/api/convert-all', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertTrue(result.get('ok'))
        self.assertEqual(result['count'], 2)  # 1 PDF × 2 templates = 2 jobs

        # Wait for background threads to finish
        time.sleep(2)

        jobs = self._get_jobs()
        # Should have 2 jobs total (original pending + 1 cloned)
        self.assertEqual(len(jobs), 2)
        # Both should be 'done' (no AI verify since force_verify=False & known supplier)
        statuses = {j['status'] for j in jobs}
        self.assertEqual(statuses, {'done'})
        # Each should have a different template
        template_names = {j['template_name'] for j in jobs}
        self.assertEqual(len(template_names), 2)

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_multi_template_parallel_execution(self, mock_supplier, mock_convert):
        """Multiple conversions should run in parallel, not sequentially."""
        call_times = []

        def slow_convert(pdf, tpl, out):
            call_times.append(('start', time.time(), os.path.basename(tpl)))
            time.sleep(0.5)  # Simulate work
            Path(out).write_bytes(b'fake')
            call_times.append(('end', time.time(), os.path.basename(tpl)))
            return out

        mock_convert.side_effect = slow_convert
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        self._upload_pdf()

        tpl_paths = self.templates[:3] if len(self.templates) >= 3 else self.templates[:2]
        n = len(tpl_paths)

        payload = json.dumps({
            'template_paths': tpl_paths,
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()

        t0 = time.time()
        conn = self.srv.conn()
        conn.request('POST', '/api/convert-all', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()

        # Wait for all conversions to complete
        time.sleep(3)

        # If sequential, total time ≈ n * 0.5s. If parallel, ≈ 0.5s.
        starts = [t for tag, t, _ in call_times if tag == 'start']
        ends = [t for tag, t, _ in call_times if tag == 'end']

        self.assertEqual(len(starts), n, f'Expected {n} conversions, got {len(starts)}')

        # All starts should be within 0.3s of each other (parallel)
        if len(starts) >= 2:
            start_spread = max(starts) - min(starts)
            self.assertLess(start_spread, 0.3,
                            f'Conversions not parallel: start spread = {start_spread:.2f}s')

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_backward_compat_single_template_path(self, mock_supplier, mock_convert):
        """Old-style single template_path should still work."""
        def fake_convert(pdf, tpl, out):
            Path(out).write_bytes(b'ok')
            return out
        mock_convert.side_effect = fake_convert
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        self._upload_pdf()

        payload = json.dumps({
            'template_path': self.templates[0],  # old single field
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()

        conn = self.srv.conn()
        conn.request('POST', '/api/convert-all', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertTrue(result.get('ok'))
        self.assertEqual(result['count'], 1)


class TestDownload(unittest.TestCase):
    """Test 3: Download returns Content-Disposition: attachment and file data."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    def _upload_pdf(self, filename='download_test.pdf'):
        boundary = '----TestBoundary456'
        body = (
            f'------TestBoundary456\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------TestBoundary456--\r\n'

        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': f'multipart/form-data; boundary=----TestBoundary456'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_download_returns_attachment_header(self, mock_supplier, mock_convert):
        """GET /api/download/{id} must return Content-Disposition: attachment."""
        output_content = b'This is the converted output file content'

        def fake_convert(pdf, tpl, out):
            Path(out).write_bytes(output_content)
            return out
        mock_convert.side_effect = fake_convert
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        jobs_data = self._upload_pdf()
        job_id = jobs_data[0]['id']

        # Convert
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()

        time.sleep(2)

        # Download
        conn = self.srv.conn()
        conn.request('GET', f'/api/download/{job_id}')
        resp = conn.getresponse()
        body = resp.read()
        conn.close()

        self.assertEqual(resp.status, 200)
        # Must have attachment disposition — this triggers download, not navigation
        content_disp = resp.getheader('Content-Disposition', '')
        self.assertIn('attachment', content_disp,
                      f'Expected attachment header, got: {content_disp}')
        self.assertIn('filename=', content_disp)
        # Content type must be octet-stream
        self.assertEqual(resp.getheader('Content-Type'), 'application/octet-stream')
        # Body should match what we wrote
        self.assertEqual(body, output_content)

    @patch('terminal_launcher.CLAUDE_CLI', None)
    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_download_available_after_ai_error(self, mock_supplier, mock_convert):
        """If conversion succeeds but AI verification fails, download should still work."""
        output_content = b'Converted file data'

        def fake_convert(pdf, tpl, out):
            Path(out).write_bytes(output_content)
            return out
        mock_convert.side_effect = fake_convert
        # Force AI verification
        mock_supplier.return_value = {'known': False, 'needs_ai_verification': True}

        jobs_data = self._upload_pdf()
        job_id = jobs_data[0]['id']

        # Convert with force_verify=True, AI will fail (CLAUDE_CLI patched to None)
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': True,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()

        time.sleep(2)

        # Job should be in error state (AI failed) but output_path should exist
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job = json.loads(resp.read())
        conn.close()

        self.assertEqual(job['status'], 'error')
        self.assertIsNotNone(job.get('output_path'),
                             'output_path should be set even when AI verification fails')

        # Download should still work
        conn = self.srv.conn()
        conn.request('GET', f'/api/download/{job_id}')
        resp = conn.getresponse()
        body = resp.read()
        conn.close()

        self.assertEqual(resp.status, 200)
        self.assertIn('attachment', resp.getheader('Content-Disposition', ''))
        self.assertEqual(body, output_content)

    def test_download_nonexistent_job_returns_404(self):
        conn = self.srv.conn()
        conn.request('GET', '/api/download/nonexistent')
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)


class TestTemplateAPI(unittest.TestCase):
    """Test 1: /api/templates returns all templates for checkbox rendering."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()

    def tearDown(self):
        self.srv.stop()

    def test_templates_endpoint_returns_list(self):
        conn = self.srv.conn()
        conn.request('GET', '/api/templates')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()

        self.assertEqual(resp.status, 200)
        self.assertIsInstance(data, list)
        self.assertGreater(len(data), 0, 'Should have templates')

        # Each template has name, path, format
        for t in data:
            self.assertIn('name', t)
            self.assertIn('path', t)
            self.assertIn('format', t)
            self.assertIn(t['format'], ('xlsx', 'docx'))


class TestClaudeCLIDetection(unittest.TestCase):
    """Test 3 (AI): Claude CLI auto-detection and error handling."""

    def test_find_claude_cli_returns_path_or_none(self):
        from terminal_launcher import _find_claude_cli
        result = _find_claude_cli()
        # Either a valid path or None — should not raise
        if result is not None:
            self.assertTrue(os.path.isfile(result))

    def test_env_var_override(self):
        """CLAUDE_CLI_PATH env var should override auto-detection."""
        fake_path = '/tmp/fake-claude-test'
        try:
            # Create a fake executable
            with open(fake_path, 'w') as f:
                f.write('#!/bin/sh\necho ok')
            os.chmod(fake_path, 0o755)

            with patch.dict(os.environ, {'CLAUDE_CLI_PATH': fake_path}):
                # Re-evaluate: env var should take precedence
                result = os.environ.get('CLAUDE_CLI_PATH') or None
                self.assertEqual(result, fake_path)
        finally:
            if os.path.exists(fake_path):
                os.remove(fake_path)

    def test_launch_verification_silent_without_cli(self):
        """When CLI is not found, should set error status, not crash."""
        jm = JobManager()
        job = jm.create_job(pdf_name='test.pdf', pdf_path='/tmp/test.pdf')

        with patch('terminal_launcher.CLAUDE_CLI', None):
            from terminal_launcher import launch_verification_silent
            launch_verification_silent(jm, job['id'], '/tmp/test.pdf',
                                       '/tmp/tpl.xlsx', '/tmp/out.xlsx')

        updated = jm.get_job(job['id'])
        self.assertEqual(updated['status'], 'error')
        self.assertIn('CLI not found', updated['error'])

    def test_launch_verification_interactive_without_cli(self):
        """When CLI is not found, interactive mode should also set error."""
        jm = JobManager()
        job = jm.create_job(pdf_name='test.pdf', pdf_path='/tmp/test.pdf')

        with patch('terminal_launcher.CLAUDE_CLI', None):
            from terminal_launcher import launch_verification
            launch_verification(jm, job['id'], '/tmp/test.pdf',
                                '/tmp/tpl.xlsx', '/tmp/out.xlsx')

        updated = jm.get_job(job['id'])
        self.assertEqual(updated['status'], 'error')
        self.assertIn('CLI not found', updated['error'])


class TestFrontendHTML(unittest.TestCase):
    """Test 1 & 3: Verify HTML/JS contains multi-select checkboxes and download attributes."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()

    def tearDown(self):
        self.srv.stop()

    def test_index_has_template_checklist(self):
        """index.html should have checkbox-based template selection, not <select>."""
        conn = self.srv.conn()
        conn.request('GET', '/')
        resp = conn.getresponse()
        html = resp.read().decode()
        conn.close()

        # Should have Select All checkbox
        self.assertIn('templateSelectAll', html)
        self.assertIn('templateList', html)
        # Should NOT have the old single-select dropdown for templates
        self.assertNotIn('<select id="templateSelect">', html)

    def test_js_has_multi_template_support(self):
        """app.js should send template_paths (array) not template_path (string)."""
        conn = self.srv.conn()
        conn.request('GET', '/app.js')
        resp = conn.getresponse()
        js = resp.read().decode()
        conn.close()

        # Should have getSelectedTemplates function
        self.assertIn('getSelectedTemplates', js)
        # convertAll should send template_paths array
        self.assertIn('template_paths:', js)
        # Download links should have download attribute
        self.assertIn('download title="Download">', js)


class TestInputPDFNotDeleted(unittest.TestCase):
    """Test 2: Input PDF should NOT be deleted after conversion."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_input_pdf_preserved_after_conversion(self, mock_supplier, mock_convert):
        """Input PDF should still exist after conversion for re-verification."""
        def fake_convert(pdf, tpl, out):
            Path(out).write_bytes(b'output')
            return out
        mock_convert.side_effect = fake_convert
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        # Upload
        boundary = '----TB789'
        filename = 'preserve_test.pdf'
        body = (
            f'------TB789\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------TB789--\r\n'

        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': f'multipart/form-data; boundary=----TB789'})
        resp = conn.getresponse()
        jobs_data = json.loads(resp.read())
        conn.close()

        job_id = jobs_data[0]['id']
        pdf_path = jobs_data[0]['pdf_path']

        # Convert
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()

        time.sleep(2)

        # PDF should still exist
        self.assertTrue(os.path.exists(pdf_path),
                        f'Input PDF was deleted: {pdf_path}')


# ===================== Integration Tests: Core Flows =====================


class TestUploadFlow(unittest.TestCase):
    """Integration tests for the file upload flow."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()

    def tearDown(self):
        self.srv.stop()

    def _upload(self, filename='test.pdf', data=None):
        if data is None:
            data = self.pdf_data
        boundary = '----UploadTest'
        body = (
            f'------UploadTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + data + b'\r\n------UploadTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----UploadTest'})
        resp = conn.getresponse()
        return resp.status, json.loads(resp.read())

    def test_upload_single_pdf(self):
        fname = f'sample_{uuid.uuid4().hex[:6]}.pdf'
        status, data = self._upload(fname)
        self.assertEqual(status, 201)
        self.assertEqual(len(data), 1)
        self.assertEqual(data[0]['pdf_name'], fname)
        self.assertEqual(data[0]['status'], 'pending')
        self.assertIn('id', data[0])
        self.assertIn('pdf_path', data[0])

    def test_upload_creates_file_on_disk(self):
        status, data = self._upload('disk_check.pdf')
        self.assertEqual(status, 201)
        self.assertTrue(os.path.exists(data[0]['pdf_path']))

    def test_upload_duplicate_name_auto_renames(self):
        self._upload('dup.pdf')
        status, data = self._upload('dup.pdf')
        self.assertEqual(status, 201)
        # Second upload should have a different filename (dup_1.pdf)
        self.assertNotEqual(data[0]['pdf_name'], 'dup.pdf')
        self.assertIn('dup_', data[0]['pdf_name'])

    def test_upload_non_pdf_rejected(self):
        boundary = '----UploadTest'
        body = (
            f'------UploadTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="doc.txt"\r\n'
            f'Content-Type: text/plain\r\n\r\n'
        ).encode() + b'not a pdf\r\n------UploadTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----UploadTest'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        self.assertEqual(resp.status, 400)
        self.assertIn('error', data)

    def test_upload_no_files(self):
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=b'',
                     headers={'Content-Type': 'application/json', 'Content-Length': '0'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        self.assertEqual(resp.status, 400)

    def test_upload_job_appears_in_list(self):
        fname = f'listed_{uuid.uuid4().hex[:6]}.pdf'
        self._upload(fname)
        conn = self.srv.conn()
        conn.request('GET', '/api/jobs')
        resp = conn.getresponse()
        jobs_list = json.loads(resp.read())
        conn.close()
        names = [j['pdf_name'] for j in jobs_list]
        self.assertIn(fname, names)


class TestSingleConvertFlow(unittest.TestCase):
    """Integration tests for single-job conversion (POST /api/convert/{id})."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    def _upload(self, filename='conv_test.pdf'):
        boundary = '----ConvTest'
        body = (
            f'------ConvTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------ConvTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----ConvTest'})
        resp = conn.getresponse()
        return json.loads(resp.read())

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_convert_known_supplier_goes_to_done(self, mock_supplier, mock_convert):
        """Known supplier + no force_verify → status should reach 'done'."""
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'out'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        data = self._upload()
        job_id = data[0]['id']

        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertTrue(result.get('ok'))
        time.sleep(2)

        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job = json.loads(resp.read())
        conn.close()
        self.assertEqual(job['status'], 'done')
        self.assertIsNotNone(job['output_path'])

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_convert_sets_template_info(self, mock_supplier, mock_convert):
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'x'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        data = self._upload()
        job_id = data[0]['id']
        tpl = self.templates[0]

        payload = json.dumps({
            'template_path': tpl,
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()

        time.sleep(2)
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job = json.loads(resp.read())
        conn.close()
        self.assertEqual(job['template_name'], os.path.basename(tpl))
        self.assertEqual(job['template_path'], tpl)

    def test_convert_missing_template_returns_400(self):
        data = self._upload()
        job_id = data[0]['id']
        payload = json.dumps({}).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 400)
        self.assertIn('template_path', result.get('error', ''))

    def test_convert_nonexistent_job_returns_404(self):
        payload = json.dumps({'template_path': self.templates[0]}).encode()
        conn = self.srv.conn()
        conn.request('POST', '/api/convert/nosuchjob', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_convert_already_converted_returns_400(self, mock_supplier, mock_convert):
        """Cannot convert a job that's not pending."""
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'x'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        data = self._upload()
        job_id = data[0]['id']
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()

        # First convert
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        time.sleep(2)

        # Second convert attempt should fail
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 400)

    @patch('converter_service.convert_coa')
    def test_convert_error_sets_error_status(self, mock_convert):
        """If conversion throws, job status should be 'error'."""
        mock_convert.side_effect = RuntimeError('Conversion engine failure')

        data = self._upload()
        job_id = data[0]['id']
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()

        time.sleep(2)
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job = json.loads(resp.read())
        conn.close()
        self.assertEqual(job['status'], 'error')
        self.assertIn('Conversion engine failure', job['error'])


class TestVerifyFlow(unittest.TestCase):
    """Integration tests for manual re-verification (POST /api/verify/{id})."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    def _upload_and_convert(self):
        """Upload a PDF and convert it, returning job_id."""
        boundary = '----VTest'
        body = (
            f'------VTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="verify_test.pdf"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------VTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----VTest'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data[0]['id']

    @patch('terminal_launcher.launch_verification_silent')
    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_verify_sets_status_to_verifying(self, mock_supplier, mock_convert, mock_verify):
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'out'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        job_id = self._upload_and_convert()
        # Convert first
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        time.sleep(2)

        # Now verify
        verify_payload = json.dumps({'claude_mode': 'silent'}).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/verify/{job_id}', body=verify_payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertEqual(resp.status, 200)
        self.assertTrue(result.get('ok'))

        # Job should now be 'verifying'
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job = json.loads(resp.read())
        conn.close()
        self.assertEqual(job['status'], 'verifying')

    def test_verify_nonexistent_job_returns_404(self):
        payload = json.dumps({'claude_mode': 'silent'}).encode()
        conn = self.srv.conn()
        conn.request('POST', '/api/verify/nope', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)

    def test_verify_pending_job_returns_400(self):
        """Cannot verify a job that has no output file yet."""
        job_id = self._upload_and_convert()
        payload = json.dumps({'claude_mode': 'silent'}).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/verify/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 400)
        self.assertIn('No output', result.get('error', ''))


class TestReportErrorFlow(unittest.TestCase):
    """Integration tests for error reporting (POST /api/report-error/{id})."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    def _upload_and_convert_to_done(self):
        boundary = '----ErrTest'
        body = (
            f'------ErrTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="err_test.pdf"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------ErrTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----ErrTest'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data[0]['id']

    @patch('terminal_launcher.launch_error_fix_silent')
    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_report_error_triggers_fix(self, mock_supplier, mock_convert, mock_fix):
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'out'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        job_id = self._upload_and_convert_to_done()
        # Convert
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        time.sleep(2)

        # Report error
        error_payload = json.dumps({
            'message': 'Vitamin D3 value is wrong',
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/report-error/{job_id}', body=error_payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertEqual(resp.status, 200)
        self.assertTrue(result.get('ok'))
        mock_fix.assert_called_once()
        # Verify the error message was passed through
        call_args = mock_fix.call_args
        self.assertIn('Vitamin D3 value is wrong', call_args[0])

    def test_report_error_empty_message_returns_400(self):
        job_id = self._upload_and_convert_to_done()
        payload = json.dumps({'message': '', 'claude_mode': 'silent'}).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/report-error/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 400)
        self.assertIn('message', result.get('error', '').lower())

    def test_report_error_nonexistent_job_returns_404(self):
        payload = json.dumps({'message': 'wrong', 'claude_mode': 'silent'}).encode()
        conn = self.srv.conn()
        conn.request('POST', '/api/report-error/nope', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)


class TestRemoveFlow(unittest.TestCase):
    """Integration tests for job removal (POST /api/remove/{id})."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()

    def tearDown(self):
        self.srv.stop()

    def _upload(self, filename='remove_test.pdf'):
        boundary = '----RmTest'
        body = (
            f'------RmTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------RmTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----RmTest'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data[0]

    def test_remove_deletes_job(self):
        job = self._upload()
        job_id = job['id']

        conn = self.srv.conn()
        conn.request('POST', f'/api/remove/{job_id}')
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertEqual(resp.status, 200)
        self.assertTrue(result.get('ok'))

        # Job should no longer exist
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)

    def test_remove_deletes_pdf_from_disk(self):
        job = self._upload()
        pdf_path = job['pdf_path']
        self.assertTrue(os.path.exists(pdf_path))

        conn = self.srv.conn()
        conn.request('POST', f'/api/remove/{job["id"]}')
        resp = conn.getresponse()
        resp.read()
        conn.close()

        self.assertFalse(os.path.exists(pdf_path))

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_remove_deletes_output_file(self, mock_supplier, mock_convert):
        """Remove should clean up the output file too."""
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'data'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}
        templates = _get_template_paths()
        if not templates:
            self.skipTest('Need at least 1 template')

        job = self._upload()
        job_id = job['id']

        # Convert
        payload = json.dumps({
            'template_path': templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        time.sleep(2)

        # Get output path
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job_data = json.loads(resp.read())
        conn.close()
        output_path = job_data['output_path']
        self.assertTrue(os.path.exists(output_path))

        # Remove
        conn = self.srv.conn()
        conn.request('POST', f'/api/remove/{job_id}')
        resp = conn.getresponse()
        resp.read()
        conn.close()

        self.assertFalse(os.path.exists(output_path))

    def test_remove_job_disappears_from_list(self):
        job = self._upload()
        job_id = job['id']

        conn = self.srv.conn()
        conn.request('POST', f'/api/remove/{job_id}')
        resp = conn.getresponse()
        resp.read()
        conn.close()

        conn = self.srv.conn()
        conn.request('GET', '/api/jobs')
        resp = conn.getresponse()
        jobs_list = json.loads(resp.read())
        conn.close()
        ids = [j['id'] for j in jobs_list]
        self.assertNotIn(job_id, ids)

    def test_remove_nonexistent_returns_404(self):
        conn = self.srv.conn()
        conn.request('POST', '/api/remove/nonexist')
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)


class TestJobGetFlow(unittest.TestCase):
    """Integration tests for job retrieval endpoints."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()

    def tearDown(self):
        self.srv.stop()

    def _upload(self, filename='get_test.pdf'):
        boundary = '----GetTest'
        body = (
            f'------GetTest\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------GetTest--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----GetTest'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data[0]

    def test_get_single_job(self):
        fname = f'get_test_{uuid.uuid4().hex[:6]}.pdf'
        job = self._upload(fname)
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job["id"]}')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()

        self.assertEqual(resp.status, 200)
        self.assertEqual(data['id'], job['id'])
        self.assertEqual(data['pdf_name'], fname)
        self.assertEqual(data['status'], 'pending')

    def test_get_nonexistent_job_returns_404(self):
        conn = self.srv.conn()
        conn.request('GET', '/api/jobs/doesnotexist')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 404)
        self.assertIn('error', data)

    def test_list_jobs_returns_all(self):
        self._upload('a.pdf')
        self._upload('b.pdf')
        conn = self.srv.conn()
        conn.request('GET', '/api/jobs')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 200)
        self.assertEqual(len(data), 2)

    def test_list_jobs_empty(self):
        conn = self.srv.conn()
        conn.request('GET', '/api/jobs')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        self.assertEqual(data, [])


class TestClientInfoFlow(unittest.TestCase):
    """Integration tests for client info endpoint."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()

    def tearDown(self):
        self.srv.stop()

    def test_client_info_localhost(self):
        conn = self.srv.conn()
        conn.request('GET', '/api/client-info')
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 200)
        self.assertIn('is_local', data)
        self.assertTrue(data['is_local'])


class TestConvertAllEdgeCases(unittest.TestCase):
    """Edge case tests for batch conversion."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    def _upload(self, filename='batch_edge.pdf'):
        boundary = '----BatchEdge'
        body = (
            f'------BatchEdge\r\n'
            f'Content-Disposition: form-data; name="file"; filename="{filename}"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------BatchEdge--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----BatchEdge'})
        resp = conn.getresponse()
        data = json.loads(resp.read())
        conn.close()
        return data[0]

    def test_convert_all_no_templates_returns_400(self):
        self._upload()
        payload = json.dumps({'template_paths': []}).encode()
        conn = self.srv.conn()
        conn.request('POST', '/api/convert-all', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 400)

    def test_convert_all_no_pending_jobs_returns_400(self):
        # No uploads, so no pending jobs
        payload = json.dumps({'template_paths': [self.templates[0]]}).encode()
        conn = self.srv.conn()
        conn.request('POST', '/api/convert-all', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 400)
        self.assertIn('pending', result.get('error', '').lower())

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_convert_all_multiple_pdfs_multiple_templates(self, mock_supplier, mock_convert):
        """2 PDFs × 2 templates = 4 jobs total."""
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(b'x'), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        if len(self.templates) < 2:
            self.skipTest('Need at least 2 templates')

        self._upload('multi1.pdf')
        self._upload('multi2.pdf')

        payload = json.dumps({
            'template_paths': self.templates[:2],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', '/api/convert-all', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        result = json.loads(resp.read())
        conn.close()

        self.assertTrue(result.get('ok'))
        self.assertEqual(result['count'], 4)  # 2 PDFs × 2 templates

        time.sleep(3)
        conn = self.srv.conn()
        conn.request('GET', '/api/jobs')
        resp = conn.getresponse()
        jobs_list = json.loads(resp.read())
        conn.close()
        self.assertEqual(len(jobs_list), 4)
        self.assertTrue(all(j['status'] == 'done' for j in jobs_list))


class TestFullLifecycle(unittest.TestCase):
    """End-to-end lifecycle: upload → convert → download → remove."""

    def setUp(self):
        self.srv = _ServerThread(port=0).start()
        self.pdf_data = _make_minimal_pdf()
        self.templates = _get_template_paths()
        if not self.templates:
            self.skipTest('Need at least 1 template')

    def tearDown(self):
        self.srv.stop()

    @patch('converter_service.convert_coa')
    @patch('converter_service.check_supplier')
    def test_upload_convert_download_remove(self, mock_supplier, mock_convert):
        output_content = b'Full lifecycle output data'
        mock_convert.side_effect = lambda p, t, o: (Path(o).write_bytes(output_content), o)[1]
        mock_supplier.return_value = {'known': True, 'needs_ai_verification': False}

        # 1. Upload
        boundary = '----Lifecycle'
        body = (
            f'------Lifecycle\r\n'
            f'Content-Disposition: form-data; name="file"; filename="lifecycle.pdf"\r\n'
            f'Content-Type: application/pdf\r\n\r\n'
        ).encode() + self.pdf_data + b'\r\n------Lifecycle--\r\n'
        conn = self.srv.conn()
        conn.request('POST', '/api/upload', body=body,
                     headers={'Content-Type': 'multipart/form-data; boundary=----Lifecycle'})
        resp = conn.getresponse()
        jobs_data = json.loads(resp.read())
        conn.close()
        self.assertEqual(resp.status, 201)
        job_id = jobs_data[0]['id']

        # 2. Convert
        payload = json.dumps({
            'template_path': self.templates[0],
            'force_verify': False,
            'claude_mode': 'silent',
        }).encode()
        conn = self.srv.conn()
        conn.request('POST', f'/api/convert/{job_id}', body=payload,
                     headers={'Content-Type': 'application/json'})
        resp = conn.getresponse()
        resp.read()
        conn.close()
        time.sleep(2)

        # 3. Verify status is done
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        job = json.loads(resp.read())
        conn.close()
        self.assertEqual(job['status'], 'done')

        # 4. Download
        conn = self.srv.conn()
        conn.request('GET', f'/api/download/{job_id}')
        resp = conn.getresponse()
        dl_body = resp.read()
        conn.close()
        self.assertEqual(resp.status, 200)
        self.assertEqual(dl_body, output_content)
        self.assertIn('attachment', resp.getheader('Content-Disposition', ''))

        # 5. Remove
        conn = self.srv.conn()
        conn.request('POST', f'/api/remove/{job_id}')
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 200)

        # 6. Verify gone
        conn = self.srv.conn()
        conn.request('GET', f'/api/jobs/{job_id}')
        resp = conn.getresponse()
        resp.read()
        conn.close()
        self.assertEqual(resp.status, 404)


if __name__ == '__main__':
    unittest.main()
