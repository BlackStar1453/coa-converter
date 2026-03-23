// COA Converter Web — Frontend Logic

const API = '';
let templates = [];
let pollTimer = null;

// --- Init ---
document.addEventListener('DOMContentLoaded', () => {
  loadTemplates();
  setupUpload();
  setupButtons();
  startPolling();
});

// --- Templates ---
async function loadTemplates() {
  try {
    const res = await fetch(`${API}/api/templates`);
    templates = await res.json();
    const sel = document.getElementById('templateSelect');
    sel.innerHTML = '<option value="">-- Select Template --</option>';
    templates.forEach(t => {
      const opt = document.createElement('option');
      opt.value = t.path;
      opt.textContent = `${t.name} (${t.format})`;
      sel.appendChild(opt);
    });
  } catch (e) {
    console.error('Failed to load templates:', e);
  }
}

// --- Upload ---
function setupUpload() {
  const zone = document.getElementById('dropZone');
  const input = document.getElementById('fileInput');

  zone.addEventListener('click', () => input.click());

  zone.addEventListener('dragover', e => {
    e.preventDefault();
    zone.classList.add('dragover');
  });

  zone.addEventListener('dragleave', () => {
    zone.classList.remove('dragover');
  });

  zone.addEventListener('drop', e => {
    e.preventDefault();
    zone.classList.remove('dragover');
    if (e.dataTransfer.files.length) {
      uploadFiles(e.dataTransfer.files);
    }
  });

  input.addEventListener('change', () => {
    if (input.files.length) {
      uploadFiles(input.files);
      input.value = '';
    }
  });
}

async function uploadFiles(fileList) {
  const form = new FormData();
  let count = 0;
  for (const f of fileList) {
    if (f.name.toLowerCase().endsWith('.pdf')) {
      form.append('file', f, f.name);
      count++;
    }
  }
  if (!count) return;

  try {
    const res = await fetch(`${API}/api/upload`, { method: 'POST', body: form });
    const data = await res.json();
    if (res.ok) {
      refreshJobs();
      updateConvertButton();
    } else {
      alert(data.error || 'Upload failed');
    }
  } catch (e) {
    alert('Upload failed: ' + e.message);
  }
}

// --- Buttons ---
function setupButtons() {
  document.getElementById('btnConvertAll').addEventListener('click', convertAll);
  document.getElementById('templateSelect').addEventListener('change', updateConvertButton);
}

function updateConvertButton() {
  const tpl = document.getElementById('templateSelect').value;
  document.getElementById('btnConvertAll').disabled = !tpl;
}

function getClaudeMode() {
  return document.getElementById('claudeMode').value;
}

async function convertAll() {
  const tpl = document.getElementById('templateSelect').value;
  if (!tpl) return;

  const force = document.getElementById('forceVerify').checked;

  try {
    const res = await fetch(`${API}/api/convert-all`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template_path: tpl, force_verify: force, claude_mode: getClaudeMode() }),
    });
    const data = await res.json();
    if (!res.ok) alert(data.error || 'Convert failed');
    refreshJobs();
  } catch (e) {
    alert('Convert failed: ' + e.message);
  }
}

async function convertOne(jobId) {
  const tpl = document.getElementById('templateSelect').value;
  if (!tpl) { alert('Please select a template first'); return; }

  const force = document.getElementById('forceVerify').checked;

  try {
    const res = await fetch(`${API}/api/convert/${jobId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template_path: tpl, force_verify: force, claude_mode: getClaudeMode() }),
    });
    const data = await res.json();
    if (!res.ok) alert(data.error || 'Convert failed');
    refreshJobs();
  } catch (e) {
    alert('Convert failed: ' + e.message);
  }
}

async function verifyJob(jobId) {
  try {
    const res = await fetch(`${API}/api/verify/${jobId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ claude_mode: getClaudeMode() }),
    });
    const data = await res.json();
    if (!res.ok) alert(data.error || 'Verify failed');
    refreshJobs();
  } catch (e) {
    alert('Verify failed: ' + e.message);
  }
}

async function focusTerminal() {
  try {
    await fetch(`${API}/api/focus-terminal`, { method: 'POST' });
  } catch (e) {
    // ignore
  }
}

async function deleteJob(jobId) {
  try {
    await fetch(`${API}/api/jobs/${jobId}`, { method: 'DELETE' });
    refreshJobs();
  } catch (e) {
    // ignore
  }
}

// --- Job Display ---
async function refreshJobs() {
  try {
    const res = await fetch(`${API}/api/jobs`);
    const jobList = await res.json();
    renderJobs(jobList);
  } catch (e) {
    // ignore
  }
}

function renderJobs(jobList) {
  const container = document.getElementById('jobContainer');

  if (!jobList.length) {
    container.innerHTML = '<div class="empty-state">No jobs yet. Upload PDF files to start.</div>';
    return;
  }

  // Sort: newest first
  jobList.sort((a, b) => b.created_at.localeCompare(a.created_at));

  let html = `
    <table class="job-table">
      <thead>
        <tr>
          <th>File</th>
          <th>Template</th>
          <th>Status</th>
          <th>Actions</th>
        </tr>
      </thead>
      <tbody>`;

  for (const job of jobList) {
    const status = job.status;
    const badgeClass = `badge badge-${status}`;
    const statusLabel = {
      pending: 'Pending',
      converting: 'Converting...',
      verifying: 'AI Verifying...',
      done: 'Done',
      error: 'Error',
    }[status] || status;

    let actions = '';
    if (status === 'pending') {
      actions = `<button class="btn btn-primary btn-sm" onclick="convertOne('${job.id}')">Convert</button>`;
      actions += ` <button class="btn btn-secondary btn-sm" onclick="deleteJob('${job.id}')">&times;</button>`;
    } else if (status === 'converting') {
      actions = '<span style="color:#86868b;font-size:12px">Processing...</span>';
    } else if (status === 'verifying') {
      if (getClaudeMode() === 'interactive') {
        actions = `<button class="btn btn-secondary btn-sm" onclick="focusTerminal()">View Terminal</button>`;
      } else {
        actions = '<span style="color:#86868b;font-size:12px">AI verifying...</span>';
      }
    } else if (status === 'done') {
      actions = `<a class="btn btn-primary btn-sm" href="/api/download/${job.id}">Download</a>`;
      actions += ` <button class="btn btn-secondary btn-sm" onclick="verifyJob('${job.id}')">Re-verify</button>`;
      actions += ` <button class="btn btn-secondary btn-sm" onclick="deleteJob('${job.id}')">&times;</button>`;
    } else if (status === 'error') {
      actions = `<span class="error-text" title="${(job.error || '').replace(/"/g, '&quot;')}">${job.error || 'Unknown error'}</span>`;
      actions += ` <button class="btn btn-secondary btn-sm" onclick="deleteJob('${job.id}')">&times;</button>`;
    }

    let aiRow = '';
    if (job.ai_output) {
      aiRow = `<tr><td colspan="4"><details><summary style="cursor:pointer;color:#86868b;font-size:12px">AI Output</summary><pre style="white-space:pre-wrap;font-size:12px;max-height:300px;overflow:auto;background:#f5f5f7;padding:8px;border-radius:4px;margin-top:4px">${escHtml(job.ai_output)}</pre></details></td></tr>`;
    }

    html += `
        <tr>
          <td>${escHtml(job.pdf_name)}</td>
          <td>${escHtml(job.template_name || '-')}</td>
          <td><span class="${badgeClass}">${statusLabel}</span></td>
          <td><div class="actions">${actions}</div></td>
        </tr>${aiRow}`;
  }

  html += '</tbody></table>';
  container.innerHTML = html;
}

function escHtml(s) {
  if (!s) return '';
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// --- Polling ---
function startPolling() {
  refreshJobs();
  pollTimer = setInterval(refreshJobs, 2000);
}
