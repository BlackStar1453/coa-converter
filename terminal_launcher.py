"""Launch Terminal.app with Claude Code for AI verification."""

import os
import shlex
import shutil
import subprocess
import threading
import time
import logging

logger = logging.getLogger(__name__)

MARKER_DIR = '/tmp'


def _find_claude_cli():
    """Find Claude CLI binary from known paths or PATH."""
    candidates = [
        os.path.expanduser('~/.local/bin/claude'),
        os.path.expanduser('~/.claude/local/claude'),
        '/usr/local/bin/claude',
        os.path.expanduser('~/.npm-global/bin/claude'),
    ]
    for path in candidates:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path
    found = shutil.which('claude')
    if found:
        return found
    return None


CLAUDE_CLI = os.environ.get('CLAUDE_CLI_PATH') or _find_claude_cli()


def _escape_for_applescript(s: str) -> str:
    """Escape a string for embedding in AppleScript double quotes."""
    return s.replace('\\', '\\\\').replace('"', '\\"')


def launch_verification(job_manager, job_id: str, pdf_path: str,
                        template_path: str, output_path: str):
    """Open Terminal.app running Claude Code for COA verification."""
    if not CLAUDE_CLI:
        logger.error('Claude CLI not found')
        job_manager.update_job(job_id, status='error',
                               error='Claude CLI not found. Install Claude Code CLI or disable AI verification.')
        return

    marker_file = os.path.join(MARKER_DIR, f'coa-verify-{job_id}.done')

    # Remove stale marker
    if os.path.exists(marker_file):
        os.remove(marker_file)

    # Build the command to run in Terminal
    # Claude Code will execute the /coa-to-template skill with --dangerously-skip-permissions
    pdf_q = shlex.quote(pdf_path)
    tpl_q = shlex.quote(template_path)
    out_q = shlex.quote(output_path)
    marker_q = shlex.quote(marker_file)

    # Write a temp script to avoid AppleScript escaping nightmares
    script_path = f'/tmp/coa-verify-{job_id}.sh'
    script_content = f"""#!/bin/bash
trap 'echo "done" > {marker_q}' EXIT
echo "=== COA AI Verification (Job: {job_id}) ==="
echo "PDF: {pdf_path}"
echo "Template: {os.path.basename(template_path)}"
echo "Output: {output_path}"
echo "---"
{CLAUDE_CLI} --dangerously-skip-permissions "/coa-to-template {pdf_q} {tpl_q} {out_q}"
echo "=== Verification complete. You can close this window. ==="
"""
    with open(script_path, 'w') as f:
        f.write(script_content)
    os.chmod(script_path, 0o755)

    terminal_cmd = script_path

    applescript = f'tell application "Terminal" to do script "{_escape_for_applescript(terminal_cmd)}"'

    try:
        subprocess.run(['osascript', '-e', applescript], check=True,
                       capture_output=True, timeout=10)
        logger.info(f'Terminal launched for job {job_id}')
        # Bring Terminal to front
        subprocess.run(['osascript', '-e',
                        'tell application "Terminal" to activate'],
                       capture_output=True, timeout=5)
    except Exception as e:
        logger.error(f'Failed to launch Terminal for job {job_id}: {e}')
        job_manager.update_job(job_id, status='error',
                               error=f'Terminal launch failed: {e}')
        return

    # Start polling for completion marker
    _start_marker_poll(job_manager, job_id, marker_file)


def _start_marker_poll(job_manager, job_id: str, marker_file: str):
    """Poll for the completion marker file in a background thread."""

    def _poll():
        script_path = f'/tmp/coa-verify-{job_id}.sh'
        timeout = 3600  # 1 hour max
        elapsed = 0
        while elapsed < timeout:
            time.sleep(3)
            elapsed += 3
            if os.path.exists(marker_file):
                # Clean up marker and script
                for f in (marker_file, script_path):
                    try:
                        os.remove(f)
                    except OSError:
                        pass
                job_manager.update_job(job_id, status='done')
                logger.info(f'Verification complete for job {job_id}')
                return
        # Timeout - clean up script
        try:
            os.remove(script_path)
        except OSError:
            pass
        logger.warning(f'Verification timed out for job {job_id}')
        job_manager.update_job(job_id, status='error',
                               error='Verification timed out (1h)')

    t = threading.Thread(target=_poll, daemon=True)
    t.start()


def launch_verification_silent(job_manager, job_id: str, pdf_path: str,
                               template_path: str, output_path: str):
    """Run Claude Code silently with -p flag and capture output."""
    if not CLAUDE_CLI:
        logger.error('Claude CLI not found')
        job_manager.update_job(job_id, status='error',
                               error='Claude CLI not found. Install Claude Code CLI or disable AI verification.')
        return

    # No shlex.quote() here — subprocess.run() passes args directly, not via shell
    prompt = f"/coa-to-template {pdf_path} {template_path} {output_path}"

    def _run():
        try:
            result = subprocess.run(
                [CLAUDE_CLI, '--dangerously-skip-permissions', '-p', prompt],
                cwd=os.path.dirname(output_path),
                capture_output=True, text=True, timeout=300,
            )
            output = result.stdout or ''
            if result.stderr:
                output += '\n--- stderr ---\n' + result.stderr

            if result.returncode == 0:
                job_manager.update_job(job_id, status='done',
                                       ai_output=output)
                logger.info(f'Silent verification complete for job {job_id}')
            else:
                job_manager.update_job(job_id, status='error',
                                       error='Claude returned non-zero exit',
                                       ai_output=output)
        except subprocess.TimeoutExpired:
            job_manager.update_job(job_id, status='error',
                                   error='Claude timed out (5 min)')
        except Exception as e:
            logger.error(f'Silent verification failed for job {job_id}: {e}')
            job_manager.update_job(job_id, status='error',
                                   error=f'Silent run failed: {e}')

    t = threading.Thread(target=_run, daemon=True)
    t.start()


def focus_terminal():
    """Bring Terminal.app to front."""
    try:
        subprocess.run(['osascript', '-e',
                        'tell application "Terminal" to activate'],
                       capture_output=True, timeout=5)
    except Exception:
        pass
