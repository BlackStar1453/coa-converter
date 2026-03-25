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

async function removeJob(jobId) {
  try {
    await fetch(`${API}/api/remove/${jobId}`, { method: 'POST' });
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
    const ico = (svg, cls='') => `<svg class="icon${cls ? ' '+cls : ''}" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">${svg}</svg>`;
    const icoConvert = ico('<polygon points="6 3 20 12 6 21 6 3"/>');
    const icoDownload = ico('<path d="M21 15v4a2 2 0 0 1-2 2H5a2 2 0 0 1-2-2v-4"/><polyline points="7 10 12 15 17 10"/><line x1="12" y1="15" x2="12" y2="3"/>');
    const icoVerify = ico('<polyline points="20 6 9 17 4 12"/>');
    const icoTerminal = ico('<polyline points="4 17 10 11 4 5"/><line x1="12" y1="19" x2="20" y2="19"/>');
    const icoReport = ico('<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>');
    const icoClose = ico('<line x1="18" y1="6" x2="6" y2="18"/><line x1="6" y1="6" x2="18" y2="18"/>');
    const closeBtn = ` <button class="btn-icon" onclick="removeJob('${job.id}')" title="Delete job and files">${icoClose}</button>`;

    if (status === 'pending') {
      actions = `<button class="btn-icon btn-icon-primary" onclick="convertOne('${job.id}')" title="Convert">${icoConvert}</button>`;
      actions += closeBtn;
    } else if (status === 'converting') {
      actions = '<span style="color:#86868b;font-size:12px">Processing...</span>';
      actions += closeBtn;
    } else if (status === 'converted') {
      actions = `<a class="btn-icon btn-icon-primary" href="/api/download/${job.id}" download title="Download">${icoDownload}</a>`;
      actions += ` <button class="btn-icon" onclick="verifyJob('${job.id}')" title="Verify">${icoVerify}</button>`;
      actions += closeBtn;
    } else if (status === 'verifying') {
      actions = `<a class="btn-icon btn-icon-primary" href="/api/download/${job.id}" download title="Download">${icoDownload}</a>`;
      if (isLocal && getClaudeMode() === 'interactive') {
        actions += ` <button class="btn-icon" onclick="focusTerminal()" title="View Terminal">${icoTerminal}</button>`;
      } else {
        actions += ' <span style="color:#86868b;font-size:12px">AI verifying...</span>';
      }
      actions += closeBtn;
    } else if (status === 'done') {
      actions = `<a class="btn-icon btn-icon-primary" href="/api/download/${job.id}" download title="Download">${icoDownload}</a>`;
      actions += ` <button class="btn-icon btn-icon-warning" onclick="reportError('${job.id}')" title="Report Error">${icoReport}</button>`;
      actions += closeBtn;
    } else if (status === 'error') {
      actions = `<span class="error-text" title="${(job.error || '').replace(/"/g, '&quot;')}">${job.error || 'Unknown error'}</span>`;
      if (job.output_path) {
        actions += ` <a class="btn-icon btn-icon-primary" href="/api/download/${job.id}" download title="Download">${icoDownload}</a>`;
        actions += ` <button class="btn-icon btn-icon-warning" onclick="reportError('${job.id}')" title="Report Error">${icoReport}</button>`;
      }
      actions += closeBtn;
    }

    html += `
        <tr>
          <td>${escHtml(job.pdf_name)}</td>
          <td>${escHtml(job.template_name || '-')}</td>
          <td><span class="${badgeClass}">${statusLabel}</span></td>
          <td><div class="actions">${actions}</div></td>
        </tr>`;
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
