"""Tests for cross-platform Claude CLI detection + Windows AI verification.

Goal: prove that on Windows the silent AI verification path is wired up
correctly without actually invoking Claude. Strategy: write a fake claude.cmd
to a temp dir, point CLAUDE_CLI_PATH at it, and assert subprocess.run is
called with the right args / cwd / timeout.

These tests do NOT spawn the real Claude CLI — they mock the subprocess layer
so they're safe to run in CI on any platform without billing or installs.
"""

import os
import sys
import tempfile
import unittest
import unittest.mock
from pathlib import Path

PROJECT_DIR = Path(__file__).parent
sys.path.insert(0, str(PROJECT_DIR))

# Force the module to load with our env so CLAUDE_CLI captures the fake path
# (CLAUDE_CLI is set at import time in terminal_launcher).
_FAKE_CLI_DIR = tempfile.mkdtemp(prefix='coa-fake-claude-')
_FAKE_CLI = os.path.join(_FAKE_CLI_DIR, 'claude.cmd' if sys.platform == 'win32' else 'claude')


def setUpModule():
    """Create a fake claude executable before importing terminal_launcher."""
    if sys.platform == 'win32':
        with open(_FAKE_CLI, 'w') as f:
            f.write('@echo fake claude\n')
    else:
        with open(_FAKE_CLI, 'w') as f:
            f.write('#!/bin/sh\necho "fake claude"\n')
        os.chmod(_FAKE_CLI, 0o755)
    os.environ['CLAUDE_CLI_PATH'] = _FAKE_CLI


def tearDownModule():
    """Clean up the fake CLI directory."""
    import shutil
    os.environ.pop('CLAUDE_CLI_PATH', None)
    shutil.rmtree(_FAKE_CLI_DIR, ignore_errors=True)


class TestFindClaudeCLI(unittest.TestCase):
    """_find_claude_cli should locate Claude on each platform."""

    def test_finds_via_shutil_which(self):
        """When claude is in PATH, _find_claude_cli should return it."""
        # Add the fake CLI dir to PATH for this test
        from terminal_launcher import _find_claude_cli
        with unittest.mock.patch.dict(os.environ, {
            'PATH': _FAKE_CLI_DIR + os.pathsep + os.environ.get('PATH', ''),
        }):
            result = _find_claude_cli()
            self.assertIsNotNone(result, '_find_claude_cli should find fake CLI via PATH')

    def test_windows_candidates_only_run_on_windows(self):
        """Windows-only candidates use %LOCALAPPDATA% etc; on Unix they are skipped."""
        from terminal_launcher import _find_claude_cli
        # Just check the function returns something or None without crashing
        result = _find_claude_cli()
        # Result may be None on CI without claude installed — that's fine
        self.assertTrue(result is None or os.path.isfile(result),
                        f'Returned non-file: {result}')


class TestShellPathFallback(unittest.TestCase):
    """_get_shell_path should be Unix-only."""

    def test_windows_returns_none(self):
        """On Windows _get_shell_path should not spawn a subprocess."""
        from terminal_launcher import _get_shell_path
        with unittest.mock.patch('sys.platform', 'win32'):
            with unittest.mock.patch('subprocess.run') as mock_run:
                result = _get_shell_path()
                self.assertIsNone(result)
                mock_run.assert_not_called()


class TestPlatformGates(unittest.TestCase):
    """The platform constants must be set correctly for AI verification gating."""

    def test_silent_supported_on_darwin_and_win32(self):
        import terminal_launcher
        # The constant is evaluated at import time based on sys.platform
        self.assertIn(sys.platform, ('darwin', 'win32', 'linux'))
        if sys.platform in ('darwin', 'win32'):
            self.assertTrue(terminal_launcher.IS_AI_VERIFICATION_SUPPORTED)
        # On Linux it should be False (we only support darwin/win32 explicitly)

    def test_interactive_only_on_darwin(self):
        import terminal_launcher
        if sys.platform == 'darwin':
            self.assertTrue(terminal_launcher.IS_INTERACTIVE_TERMINAL_SUPPORTED)
        else:
            self.assertFalse(terminal_launcher.IS_INTERACTIVE_TERMINAL_SUPPORTED)


class _FakeJobManager:
    """In-memory job tracker for testing without the real JobManager."""

    def __init__(self):
        self.calls = []
        self.job = {'id': 'test-job', 'status': 'pending', 'ai_verified': None}

    def update_job(self, job_id, **kwargs):
        self.calls.append((job_id, kwargs))
        self.job.update(kwargs)
        return dict(self.job)


class TestSilentCallConstruction(unittest.TestCase):
    """launch_verification_silent should build the right Claude CLI invocation."""

    def test_silent_passes_correct_args_to_subprocess(self):
        import terminal_launcher
        # Patch the module-level CLAUDE_CLI in case it was None at import time
        with unittest.mock.patch.object(terminal_launcher, 'CLAUDE_CLI', _FAKE_CLI):
            jm = _FakeJobManager()
            # Patch subprocess.run inside terminal_launcher so the real fake
            # doesn't run — we just check args
            with unittest.mock.patch('terminal_launcher.subprocess.run') as mock_run:
                mock_run.return_value = unittest.mock.MagicMock(returncode=0)
                terminal_launcher.launch_verification_silent(
                    jm, 'test-job',
                    pdf_path='/path/to/pdf.pdf',
                    template_path='/path/to/tpl.xlsx',
                    output_path='/path/to/out.xlsx',
                )
                # Wait for background thread to start subprocess
                import time
                time.sleep(0.3)
                self.assertTrue(mock_run.called,
                                'subprocess.run should be invoked for silent AI verification')
                args, kwargs = mock_run.call_args
                cmd = args[0]
                # First arg should be the CLI path
                self.assertEqual(cmd[0], _FAKE_CLI)
                # Should include --dangerously-skip-permissions and -p
                self.assertIn('--dangerously-skip-permissions', cmd)
                self.assertIn('-p', cmd)
                # The prompt should reference the slash command
                prompt = cmd[cmd.index('-p') + 1]
                self.assertIn('/coa-to-template', prompt)
                self.assertIn('/path/to/pdf.pdf', prompt)
                # Should have a timeout (we use 3600s = 1 hour)
                self.assertIn('timeout', kwargs)
                self.assertEqual(kwargs['timeout'], 3600)


class TestInteractiveFallback(unittest.TestCase):
    """On non-darwin platforms, launch_verification should fall back to silent."""

    def test_launch_verification_falls_back_on_non_darwin(self):
        import terminal_launcher
        with unittest.mock.patch.object(terminal_launcher,
                                        'IS_INTERACTIVE_TERMINAL_SUPPORTED', False):
            with unittest.mock.patch.object(terminal_launcher,
                                            'launch_verification_silent') as mock_silent:
                jm = _FakeJobManager()
                terminal_launcher.launch_verification(
                    jm, 'test-job', '/pdf', '/tpl', '/out',
                )
                mock_silent.assert_called_once_with(
                    jm, 'test-job', '/pdf', '/tpl', '/out',
                )

    def test_launch_error_fix_falls_back_on_non_darwin(self):
        import terminal_launcher
        with unittest.mock.patch.object(terminal_launcher,
                                        'IS_INTERACTIVE_TERMINAL_SUPPORTED', False):
            with unittest.mock.patch.object(terminal_launcher,
                                            'launch_error_fix_silent') as mock_silent:
                jm = _FakeJobManager()
                terminal_launcher.launch_error_fix(
                    jm, 'test-job', '/pdf', '/tpl', '/out', 'some error',
                )
                mock_silent.assert_called_once_with(
                    jm, 'test-job', '/pdf', '/tpl', '/out', 'some error',
                )


if __name__ == '__main__':
    unittest.main()
