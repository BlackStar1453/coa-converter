// COA Converter Web — Frontend Logic

const API = '';
let templates = [];
let pollTimer = null;
let isLocal = true;

// --- Init ---
document.addEventListener('DOMContentLoaded', () => {
  loadTemplates();
  setupUpload();
  setupButtons();
  startPolling();
  checkClientInfo();
});

// --- Templates ---
async function loadTemplates() {
  try {
    const res = await fetch(`${API}/api/templates`);
    templates = await res.json();
    const list = document.getElementById('templateList');
    list.innerHTML = '';
    templates.forEach(t => {
      const lbl = document.createElement('label');
      const cb = document.createElement('input');
      cb.type = 'checkbox';
      cb.value = t.path;
      cb.className = 'template-cb';
      cb.addEventListener('change', () => { updateConvertButton(); updateTemplateSummary(); });
      lbl.appendChild(cb);
      lbl.appendChild(document.createTextNode(` ${t.name} (${t.format})`));
      list.appendChild(lbl);
    });
    updateTemplateSummary();
    setupTemplateDropdown();
  } catch (e) {
    console.error('Failed to load templates:', e);
  }
}

function getSelectedTemplates() {
  return Array.from(document.querySelectorAll('.template-cb:checked')).map(cb => cb.value);
}

function updateTemplateSummary() {
  const sel = getSelectedTemplates();
  const summary = document.getElementById('templateSummary');
  if (!templates.length) {
    summary.textContent = 'No templates available';
  } else if (sel.length === 0) {
    summary.textContent = '-- Select Template --';
    summary.style.color = '#86868b';
  } else {
    summary.textContent = `${sel.length} template${sel.length > 1 ? 's' : ''} selected`;
    summary.style.color = '#1d1d1f';
  }
}

function setupTemplateDropdown() {
  const dropdown = document.getElementById('templateDropdown');
  const toggle = document.getElementById('templateToggle');

  toggle.addEventListener('click', () => {
    dropdown.classList.toggle('open');
  });

  // Close when clicking outside
  document.addEventListener('click', (e) => {
    if (!dropdown.contains(e.target)) {
      dropdown.classList.remove('open');
    }
  });
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
  document.getElementById('templateSelectAll').addEventListener('change', (e) => {
    document.querySelectorAll('.template-cb').forEach(cb => cb.checked = e.target.checked);
    updateConvertButton();
    updateTemplateSummary();
  });
}

function updateConvertButton() {
  document.getElementById('btnConvertAll').disabled = getSelectedTemplates().length === 0;
}

function getClaudeMode() {
  return document.getElementById('claudeMode').value;
}

async function checkClientInfo() {
  try {
    const res = await fetch(`${API}/api/client-info`);
    const info = await res.json();
    isLocal = info.is_local;
    if (!isLocal) {
      const sel = document.getElementById('claudeMode');
      sel.value = 'silent';
      sel.disabled = true;
      sel.title = 'Interactive mode is only available on the local machine';
    }
  } catch (e) {
    // ignore
  }
}

async function convertAll() {
  const tpls = getSelectedTemplates();
  if (!tpls.length) return;

  const force = document.getElementById('forceVerify').checked;

  try {
    const res = await fetch(`${API}/api/convert-all`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template_paths: tpls, force_verify: force, claude_mode: getClaudeMode() }),
    });
    const data = await res.json();
    if (!res.ok) alert(data.error || 'Convert failed');
    refreshJobs();
  } catch (e) {
    alert('Convert failed: ' + e.message);
  }
}

async function convertOne(jobId) {
  const tpls = getSelectedTemplates();
  if (!tpls.length) { alert('Please select at least one template'); return; }

  const force = document.getElementById('forceVerify').checked;

  try {
    const res = await fetch(`${API}/api/convert/${jobId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ template_path: tpls[0], force_verify: force, claude_mode: getClaudeMode() }),
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

async function cancelJob(jobId) {
  try {
    await fetch(`${API}/api/cancel/${jobId}`, { method: 'PUT' });
    refreshJobs();
  } catch (e) {
    // ignore
  }
}

async function removeJob(jobId) {
  if (!confirm('Remove this job and delete associated files?')) return;
  try {
    await fetch(`${API}/api/remove/${jobId}`, { method: 'PUT' });
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
      converted: 'Unverified',
      verifying: 'AI Verifying...',
      done: 'Verified',
      error: 'Error',
    }[status] || status;

    let actions = '';
    const cancelBtn = ` <button class="btn btn-secondary btn-sm" onclick="cancelJob('${job.id}')" title="Cancel">Cancel</button>`;
    const removeBtn = ` <button class="btn btn-danger btn-sm" onclick="removeJob('${job.id}')" title="Remove job and files">Remove</button>`;

    if (status === 'pending') {
      actions = `<button class="btn btn-primary btn-sm" onclick="convertOne('${job.id}')">Convert</button>`;
      actions += removeBtn;
    } else if (status === 'converting') {
      actions = '<span style="color:#86868b;font-size:12px">Processing...</span>';
      actions += cancelBtn;
    } else if (status === 'converted') {
      actions = `<a class="btn btn-primary btn-sm" href="/api/download/${job.id}" download>Download</a>`;
      actions += ` <button class="btn btn-secondary btn-sm" onclick="verifyJob('${job.id}')">Verify</button>`;
      actions += cancelBtn + removeBtn;
    } else if (status === 'verifying') {
      actions = `<a class="btn btn-primary btn-sm" href="/api/download/${job.id}" download>Download</a>`;
      if (isLocal && getClaudeMode() === 'interactive') {
        actions += ` <button class="btn btn-secondary btn-sm" onclick="focusTerminal()">View Terminal</button>`;
      } else {
        actions += ' <span style="color:#86868b;font-size:12px">AI verifying...</span>';
      }
      actions += cancelBtn;
    } else if (status === 'done') {
      actions = `<a class="btn btn-primary btn-sm" href="/api/download/${job.id}" download>Download</a>`;
      actions += ` <button class="btn btn-warning btn-sm" onclick="reportError('${job.id}')">Report Error</button>`;
      actions += ` <button class="btn btn-secondary btn-sm" onclick="verifyJob('${job.id}')">Re-verify</button>`;
      actions += cancelBtn + removeBtn;
    } else if (status === 'error') {
      actions = `<span class="error-text" title="${(job.error || '').replace(/"/g, '&quot;')}">${job.error || 'Unknown error'}</span>`;
      if (job.output_path) {
        actions += ` <a class="btn btn-primary btn-sm" href="/api/download/${job.id}" download>Download</a>`;
        actions += ` <button class="btn btn-warning btn-sm" onclick="reportError('${job.id}')">Report Error</button>`;
      }
      actions += cancelBtn + removeBtn;
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

// --- Report Error ---
async function reportError(jobId) {
  const msg = prompt('Please describe the error you found in the output:');
  if (!msg || !msg.trim()) return;

  try {
    const res = await fetch(`${API}/api/report-error/${jobId}`, {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ message: msg.trim(), claude_mode: getClaudeMode() }),
    });
    const data = await res.json();
    if (!res.ok) alert(data.error || 'Report failed');
    refreshJobs();
  } catch (e) {
    alert('Report failed: ' + e.message);
  }
}

// --- Polling ---
function startPolling() {
  refreshJobs();
  pollTimer = setInterval(refreshJobs, 2000);
}
