"""Launch Terminal.app with Claude Code for AI verification."""

import os
import shlex
import shutil
import subprocess
import sys
import threading
import time
import logging

logger = logging.getLogger(__name__)

# Silent AI verification (subprocess.run claude -p) is cross-platform — both
# macOS and Windows can use it as long as the Claude CLI is installed.
# The interactive mode (visible Terminal window with live Claude output) still
# requires macOS Terminal.app + osascript; on Windows we transparently fall
# back to silent mode (see launch_verification / launch_error_fix below).
IS_AI_VERIFICATION_SUPPORTED = sys.platform in ('darwin', 'win32')
IS_INTERACTIVE_TERMINAL_SUPPORTED = sys.platform == 'darwin'

MARKER_DIR = '/tmp'


def _skip_ai_verification(job_manager, job_id: str, reason: str):
    """Mark a job as done without AI verification (used on non-macOS platforms)."""
    logger.info(f'[AI验证] 跳过 job {job_id}: {reason}')
    job_manager.update_job(job_id, status='done', ai_verified=False)


def _find_claude_cli():
    """Find Claude CLI binary from known paths or PATH.

    Checks common installation locations for various package managers
    (Homebrew, npm global, nvm, volta, bun, pnpm; Windows native installer,
    npm global, .local\\bin) before falling back to PATH search. Also loads
    shell profile PATH on Unix to handle cases where the web server process
    has a different PATH than the user's shell.
    """
    if sys.platform == 'win32':
        local_appdata = os.environ.get('LOCALAPPDATA', '')
        appdata = os.environ.get('APPDATA', '')
        user_profile = os.environ.get('USERPROFILE', '')
        candidates = [
            # Native installer (Claude Code for Windows)
            os.path.join(local_appdata, 'Programs', 'claude', 'claude.exe'),
            # npm global install
            os.path.join(appdata, 'npm', 'claude.cmd'),
            os.path.join(appdata, 'npm', 'claude.exe'),
            # Direct user install
            os.path.join(user_profile, '.local', 'bin', 'claude.exe'),
            os.path.join(user_profile, '.claude', 'local', 'claude.exe'),
        ]
        # Filter empty strings (env var missing) and check existence — on
        # Windows os.access(..., X_OK) reports True for any readable file,
        # so we only check isfile.
        for path in candidates:
            if path and os.path.isfile(path):
                return path
        # PATH fallback: try common executable extensions
        for name in ('claude', 'claude.exe', 'claude.cmd', 'claude.bat'):
            found = shutil.which(name)
            if found:
                return found
        return None

    # Unix (macOS, Linux)
    candidates = [
        # Homebrew (Apple Silicon + Intel)
        '/opt/homebrew/bin/claude',
        '/usr/local/bin/claude',
        # Claude Code direct installs
        os.path.expanduser('~/.local/bin/claude'),
        os.path.expanduser('~/.claude/local/claude'),
        # npm global installs
        os.path.expanduser('~/.npm-global/bin/claude'),
        # nvm-managed Node.js
        *_glob_expand('~/.nvm/versions/node/*/bin/claude'),
        # volta
        os.path.expanduser('~/.volta/bin/claude'),
        # bun
        os.path.expanduser('~/.bun/bin/claude'),
        # pnpm global
        os.path.expanduser('~/.local/share/pnpm/claude'),
        # system
        '/usr/bin/claude',
    ]
    for path in candidates:
        if os.path.isfile(path) and os.access(path, os.X_OK):
            return path

    # Fallback: search PATH
    found = shutil.which('claude')
    if found:
        return found

    # Last resort: try loading the user's shell PATH
    shell_path = _get_shell_path()
    if shell_path:
        found = shutil.which('claude', path=shell_path)
        if found:
            return found

    return None


def _glob_expand(pattern: str) -> list:
    """Expand a glob pattern with ~ into a list of matching paths."""
    import glob
    return glob.glob(os.path.expanduser(pattern))


def _get_shell_path() -> str | None:
    """Load PATH from the user's login shell to catch paths not in server env.

    Unix only: spawns a login shell and echoes $PATH. On Windows the concept
    doesn't apply (PATH is already inherited from the process env via the
    standard Windows API), so we return None to skip this fallback.
    """
    if sys.platform == 'win32':
        return None
    try:
        shell = os.environ.get('SHELL', '/bin/zsh')
        result = subprocess.run(
            [shell, '-l', '-c', 'echo $PATH'],
            capture_output=True, text=True, timeout=5,
        )
        if result.returncode == 0 and result.stdout.strip():
            return result.stdout.strip()
    except Exception:
        pass
    return None


CLAUDE_CLI = os.environ.get('CLAUDE_CLI_PATH') or _find_claude_cli()


def _escape_for_applescript(s: str) -> str:
    """Escape a string for embedding in AppleScript double quotes."""
    return s.replace('\\', '\\\\').replace('"', '\\"')


def launch_verification(job_manager, job_id: str, pdf_path: str,
                        template_path: str, output_path: str):
    """Open Terminal.app running Claude Code for COA verification.

    On Windows the visible-terminal interactive flow is not implemented yet
    (PowerShell + .ps1 + new-console). Auto-fallback to the silent path so
    the AI verification still happens — the user just doesn't see live output.
    """
    if not IS_INTERACTIVE_TERMINAL_SUPPORTED:
        logger.info(f'[AI验证] 平台 {sys.platform} 不支持交互式终端，'
                    f'自动降级到 silent 模式 (job {job_id})')
        launch_verification_silent(job_manager, job_id, pdf_path,
                                   template_path, output_path)
        return
    if not CLAUDE_CLI:
        _skip_ai_verification(job_manager, job_id, 'Claude CLI 未安装')
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


def _run_claude_silent(job_manager, job_id: str, prompt: str,
                       cwd: str, label: str = 'Silent run'):
    """Run Claude CLI in a subprocess and update job status on completion."""
    logger.info(f'[{label}] job {job_id} 启动 claude: cli={CLAUDE_CLI} cwd={cwd}')
    import time
    start = time.monotonic()
    try:
        result = subprocess.run(
            [CLAUDE_CLI, '--dangerously-skip-permissions', '-p', prompt],
            cwd=cwd, stdin=subprocess.DEVNULL,
            capture_output=True, text=True, timeout=3600,
        )
        elapsed = time.monotonic() - start
        stdout_head = (result.stdout or '')[:300].replace('\n', ' | ')
        stderr_head = (result.stderr or '')[:300].replace('\n', ' | ')
        logger.info(f'[{label}] job {job_id} claude 退出: rc={result.returncode} '
                    f'耗时={elapsed:.1f}s stdout_len={len(result.stdout or "")} '
                    f'stderr_len={len(result.stderr or "")}')
        if stdout_head:
            logger.info(f'[{label}] stdout 前300字: {stdout_head}')
        if stderr_head:
            logger.info(f'[{label}] stderr 前300字: {stderr_head}')
        if result.returncode == 0:
            job_manager.update_job(job_id, status='done')
            logger.info(f'{label} complete for job {job_id}')
        else:
            job_manager.update_job(job_id, status='error',
                                   error=f'Claude returned non-zero exit (rc={result.returncode}): '
                                         f'{stderr_head or stdout_head or "no output"}')
    except subprocess.TimeoutExpired:
        job_manager.update_job(job_id, status='error',
                               error='Claude timed out (1 hour)')
    except Exception as e:
        logger.error(f'{label} failed for job {job_id}: {e}')
        job_manager.update_job(job_id, status='error',
                               error=f'{label} failed: {e}')


def launch_verification_silent(job_manager, job_id: str, pdf_path: str,
                               template_path: str, output_path: str):
    """Run Claude Code silently with -p flag and stream output.

    Cross-platform: only requires Claude CLI to be installed. If not found,
    we mark the job as done without AI verification so the rest of the
    pipeline (download, etc.) remains usable.
    """
    if not IS_AI_VERIFICATION_SUPPORTED:
        _skip_ai_verification(job_manager, job_id,
                              f'平台 {sys.platform} 暂不支持 AI 验证')
        return
    if not CLAUDE_CLI:
        _skip_ai_verification(job_manager, job_id, 'Claude CLI 未安装')
        return

    prompt = f"/coa-to-template {pdf_path} {template_path} {output_path}"

    t = threading.Thread(target=_run_claude_silent,
                         args=(job_manager, job_id, prompt,
                               os.path.dirname(output_path),
                               'Silent verification'),
                         daemon=True)
    t.start()


def launch_error_fix(job_manager, job_id: str, pdf_path: str,
                     template_path: str, output_path: str, error_msg: str):
    """Open Terminal.app running Claude Code to fix reported errors.

    On Windows we fall back to the silent error-fix path (no visible terminal).
    """
    if not IS_INTERACTIVE_TERMINAL_SUPPORTED:
        logger.info(f'[AI错误修复] 平台 {sys.platform} 不支持交互式终端，'
                    f'自动降级到 silent 模式 (job {job_id})')
        launch_error_fix_silent(job_manager, job_id, pdf_path,
                                template_path, output_path, error_msg)
        return
    if not CLAUDE_CLI:
        _skip_ai_verification(job_manager, job_id, 'Claude CLI 未安装')
        return

    marker_file = os.path.join(MARKER_DIR, f'coa-verify-{job_id}.done')
    if os.path.exists(marker_file):
        os.remove(marker_file)

    pdf_q = shlex.quote(pdf_path)
    tpl_q = shlex.quote(template_path)
    out_q = shlex.quote(output_path)
    marker_q = shlex.quote(marker_file)
    # Escape error message for shell
    err_q = shlex.quote(error_msg)

    script_path = f'/tmp/coa-fix-{job_id}.sh'
    script_content = f"""#!/bin/bash
trap 'echo "done" > {marker_q}' EXIT
echo "=== COA Error Fix (Job: {job_id}) ==="
echo "Error reported: {error_msg}"
echo "---"
{CLAUDE_CLI} --dangerously-skip-permissions "/coa-fix-output {pdf_q} {tpl_q} {out_q} User reported error: {err_q}"
echo "=== Fix complete. You can close this window. ==="
"""
    with open(script_path, 'w') as f:
        f.write(script_content)
    os.chmod(script_path, 0o755)

    applescript = f'tell application "Terminal" to do script "{_escape_for_applescript(script_path)}"'
    try:
        subprocess.run(['osascript', '-e', applescript], check=True,
                       capture_output=True, timeout=10)
        subprocess.run(['osascript', '-e',
                        'tell application "Terminal" to activate'],
                       capture_output=True, timeout=5)
    except Exception as e:
        logger.error(f'Failed to launch Terminal for error fix {job_id}: {e}')
        job_manager.update_job(job_id, status='error',
                               error=f'Terminal launch failed: {e}')
        return

    _start_marker_poll(job_manager, job_id, marker_file)


def launch_error_fix_silent(job_manager, job_id: str, pdf_path: str,
                            template_path: str, output_path: str,
                            error_msg: str):
    """Run Claude Code silently to fix reported errors, streaming output."""
    if not IS_AI_VERIFICATION_SUPPORTED:
        _skip_ai_verification(job_manager, job_id,
                              f'平台 {sys.platform} 暂不支持 AI 错误修复')
        return
    if not CLAUDE_CLI:
        _skip_ai_verification(job_manager, job_id, 'Claude CLI 未安装')
        return

    prompt = (f"/coa-fix-output {pdf_path} {template_path} {output_path} "
              f"User reported error: {error_msg}")

    t = threading.Thread(target=_run_claude_silent,
                         args=(job_manager, job_id, prompt,
                               os.path.dirname(output_path),
                               'Error fix'),
                         daemon=True)
    t.start()


def focus_terminal():
    """Bring Terminal.app to front."""
    try:
        subprocess.run(['osascript', '-e',
                        'tell application "Terminal" to activate'],
                       capture_output=True, timeout=5)
    except Exception:
        pass
