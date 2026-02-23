/* ============================================================
   Converto – PDF to EML  (runs entirely in the browser)
   ============================================================ */

// ── PDF.js setup ──────────────────────────────────────────────
import * as pdfjsLib from 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.min.mjs';
pdfjsLib.GlobalWorkerOptions.workerSrc =
  'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/4.4.168/pdf.worker.min.mjs';

// ── DOM refs ──────────────────────────────────────────────────
const pdfInput       = document.getElementById('pdfFile');
const sigSelect      = document.getElementById('signatureSelect');
const generateBtn    = document.getElementById('generateBtn');
const downloadAllBtn = document.getElementById('downloadAllBtn');
const statusEl       = document.getElementById('status');
const resultsEl      = document.getElementById('results');
const downloadHint   = document.getElementById('downloadHint');

// ── State ─────────────────────────────────────────────────────
const MAX_FILES      = 20;
const DEFAULT_SUBJECT = 'MODELO 347';
const MIN_DOWNLOAD_DELAY_MS = 400;
const MAX_DOWNLOAD_DELAY_MS = 800;
const RECIPIENT_SEARCH_HEIGHT_RATIO = 0.35;
const NO_DETECTED_EMAIL_PREFIX = 'NDE';
const SUBJECT_ANCHOR_REGEX = /Dto\.?\s*de\s*Contabilidad/i;
let selectedFiles    = [];
let results          = [];
let signaturesCache  = {};          // id → html string

// ── Init: load signature manifest ─────────────────────────────
(async () => {
  try {
    const res  = await fetch('signatures/index.json');
    const list = await res.json();
    for (const s of list) {
      const opt   = document.createElement('option');
      opt.value   = s.id;
      opt.textContent = s.label;
      sigSelect.appendChild(opt);
      // pre‑fetch each signature
      const r = await fetch(s.path);
      let html = await r.text();
      // if it's a full HTML document, extract <body> content
      const bodyMatch = html.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
      if (bodyMatch) html = bodyMatch[1];
      signaturesCache[s.id] = html.trim();
    }
  } catch (e) {
    console.warn('Could not load signatures:', e);
  }
})();

// ── Events ────────────────────────────────────────────────────
pdfInput.addEventListener('change', () => {
  selectedFiles = Array.from(pdfInput.files || []);
  downloadAllBtn.disabled = true;
  downloadHint.style.display = 'none';
  results = [];
  renderResults();
  if (selectedFiles.length > MAX_FILES) {
    setStatus(`Error: Maximum ${MAX_FILES} PDFs allowed.`, true);
    generateBtn.disabled = true;
    return;
  }
  if (!selectedFiles.length) {
    setStatus('');
    generateBtn.disabled = true;
    return;
  }
  setStatus(`${selectedFiles.length} PDF(s) selected.`);
  generateBtn.disabled = false;
});

generateBtn.addEventListener('click', async () => {
  if (!selectedFiles.length) return;
  if (selectedFiles.length > MAX_FILES) {
    setStatus(`Error: Maximum ${MAX_FILES} PDFs allowed.`, true);
    return;
  }

  setUiDisabled(true);
  results = selectedFiles.map(file => ({
    file,
    recipient: '',
    toEmail: '',
    subject: '',
    baseFilename: '',
    eml: '',
    status: 'Pending',
    error: ''
  }));
  renderResults();

  let completed = 0;
  const total = results.length;
  setStatus(`Generating EML… (0/${total})`);

  for (const result of results) {
    result.status = 'Processing';
    renderResults();
    try {
      const processed = await processPdfFile(result.file);
      Object.assign(result, processed, { status: 'Done', error: '' });
    } catch (err) {
      console.error(err);
      result.status = 'Error';
      result.error = err.message || 'Unknown error';
    }
    completed += 1;
    setStatus(`Generating EML… (${completed}/${total})`);
    renderResults();
  }

  const successCount = results.filter(r => r.status === 'Done').length;
  setStatus(`Generation complete. ${successCount}/${total} file(s) ready.`);
  downloadAllBtn.disabled = successCount === 0;
  downloadHint.style.display = successCount > 0 ? 'block' : 'none';
  setUiDisabled(false);
});

downloadAllBtn.addEventListener('click', async () => {
  if (downloadAllBtn.disabled) return;
  downloadAllBtn.disabled = true;
  await downloadAllResults();
  downloadAllBtn.disabled = false;
});

// ── Helpers ───────────────────────────────────────────────────
function setStatus(msg, isError) {
  statusEl.textContent = msg;
  statusEl.className   = isError ? 'error' : '';
}

function setUiDisabled(isDisabled) {
  pdfInput.disabled = isDisabled;
  sigSelect.disabled = isDisabled;
  generateBtn.disabled = isDisabled || !selectedFiles.length;
  if (isDisabled) {
    downloadAllBtn.disabled = true;
  }
}

async function processPdfFile(file) {
  const arrayBuf = await file.arrayBuffer();
  const pdf      = await pdfjsLib.getDocument({ data: arrayBuf }).promise;

  const pages = [];
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i);
    pages.push(await extractPage(page));
  }

  const firstPage = pages[0];
  const recipient = extractRecipient(firstPage);
  const toEmail = extractToEmail(firstPage);
  const subject = extractSubject(firstPage);

  let bodyHtml = '';
  for (const pg of pages) {
    bodyHtml += buildPageHtml(pg);
  }

  const sigId = sigSelect.value;
  if (sigId && signaturesCache[sigId]) {
    bodyHtml += '<br>' + signaturesCache[sigId];
  }

  const eml = buildEml(bodyHtml, { toEmail, subject });
  const baseFilename = buildOutputBaseFilename({
    recipient,
    toEmail,
    originalName: file.name
  });

  return { recipient, toEmail, subject, baseFilename, eml };
}

async function downloadAllResults() {
  const readyResults = results.filter(result => result.status === 'Done');
  if (!readyResults.length) return;

  const resolvedNames = resolveFileNameCollisions(
    readyResults.map(result => result.baseFilename)
  );

  for (let i = 0; i < readyResults.length; i++) {
    const filename = resolvedNames[i];
    downloadBlob(readyResults[i].eml, filename, 'message/rfc822');
    const delay = MIN_DOWNLOAD_DELAY_MS + Math.random() * (MAX_DOWNLOAD_DELAY_MS - MIN_DOWNLOAD_DELAY_MS);
    await new Promise(resolve => setTimeout(resolve, delay));
  }
}

function renderResults() {
  if (!results.length) {
    resultsEl.innerHTML = '';
    return;
  }

  const doneEntries = results
    .map((result, index) => ({ result, index }))
    .filter(entry => entry.result.status === 'Done');
  const resolvedNames = resolveFileNameCollisions(doneEntries.map(entry => entry.result.baseFilename));
  const finalNamesByIndex = new Map();
  doneEntries.forEach((entry, idx) => finalNamesByIndex.set(entry.index, resolvedNames[idx]));

  const rows = results.map((result, index) => {
    const recipient = result.recipient || 'UNKNOWN';
    const toValue = result.toEmail || NO_DETECTED_EMAIL_PREFIX;
    const subject = result.subject || '';
    const fileName = finalNamesByIndex.get(index) || (result.baseFilename ? `${result.baseFilename}.eml` : '');
    const statusClass = result.status ? result.status.toLowerCase() : 'pending';
    const error = result.error ? ` (${result.error})` : '';

    return `<tr>
      <td>${escapeHtml(result.file.name)}</td>
      <td>${escapeHtml(recipient)}</td>
      <td>${escapeHtml(toValue)}</td>
      <td>${escapeHtml(subject)}</td>
      <td>${escapeHtml(fileName)}</td>
      <td><span class="status-pill status-${statusClass}">${escapeHtml(result.status)}${escapeHtml(error)}</span></td>
    </tr>`;
  }).join('');

  resultsEl.innerHTML = `
    <table id="resultsTable">
      <thead>
        <tr>
          <th>PDF</th>
          <th>Recipient</th>
          <th>To</th>
          <th>Subject</th>
          <th>EML file</th>
          <th>Status</th>
        </tr>
      </thead>
      <tbody>${rows}</tbody>
    </table>
  `;
}

function buildOutputBaseFilename({ recipient, toEmail, originalName }) {
  const originalBase = originalName.replace(/\.[^.]+$/, '');
  if (!recipient) {
    return sanitizeFileName(`${NO_DETECTED_EMAIL_PREFIX}_UNKNOWN_${originalBase}`);
  }
  const prefix = toEmail ? '' : `${NO_DETECTED_EMAIL_PREFIX}_`;
  return sanitizeFileName(`${prefix}${recipient}`);
}

function resolveFileNameCollisions(names) {
  const counts = new Map();
  return names.map(name => {
    const safe = name || 'output';
    const current = (counts.get(safe) || 0) + 1;
    counts.set(safe, current);
    return current === 1 ? `${safe}.eml` : `${safe} (${current}).eml`;
  });
}

// ══════════════════════════════════════════════════════════════
//  PDF page extraction
// ══════════════════════════════════════════════════════════════

async function extractPage(page) {
  const viewport = page.getViewport({ scale: 1 });
  const height   = viewport.height;

  // ── text items ──────────────────────────────────────────────
  const tc    = await page.getTextContent();
  const textItems = tc.items.filter(it => it.str != null).map(it => {
    const tx = it.transform;          // [scaleX, skewY, skewX, scaleY, x, y]
    const x  = tx[4];
    const y  = height - tx[5];        // flip Y (PDF 0,0 = bottom-left)
    const w  = it.width  || 0;
    const h  = Math.abs(tx[3]) || it.height || 10;
    return { str: it.str, x, y: y - h, w, h, fontName: it.fontName };
  });

  // ── lines from operatorList ─────────────────────────────────
  const ops      = await page.getOperatorList();
  const segments = extractLineSegments(ops, height);

  return { textItems, segments, width: viewport.width, height };
}

// ── Extract line segments from operator list ──────────────────
function extractLineSegments(ops, pageH) {
  const segs  = [];
  let   cx    = 0, cy = 0;       // current point
  let   mx    = 0, my = 0;       // move-to point
  let   ctm   = [1, 0, 0, 1, 0, 0]; // current transform matrix
  const ctmStack = [];

  const OPS = pdfjsLib.OPS;

  const transformPoint = (x, y) => [
    ctm[0] * x + ctm[2] * y + ctm[4],
    pageH - (ctm[1] * x + ctm[3] * y + ctm[5])
  ];

  for (let i = 0; i < ops.fnArray.length; i++) {
    const fn   = ops.fnArray[i];
    const args = ops.argsArray[i];

    switch (fn) {
      case OPS.save:
        ctmStack.push(ctm.slice());
        break;
      case OPS.restore:
        if (ctmStack.length) ctm = ctmStack.pop();
        break;
      case OPS.transform: {
        const [a, b, c, d, e, f] = args;
        const n = [
          ctm[0]*a + ctm[2]*b,
          ctm[1]*a + ctm[3]*b,
          ctm[0]*c + ctm[2]*d,
          ctm[1]*c + ctm[3]*d,
          ctm[0]*e + ctm[2]*f + ctm[4],
          ctm[1]*e + ctm[3]*f + ctm[5],
        ];
        ctm = n;
        break;
      }
      case OPS.moveTo:
        [mx, my] = [args[0], args[1]];
        [cx, cy] = [mx, my];
        break;
      case OPS.lineTo: {
        const [x0, y0] = transformPoint(cx, cy);
        const [x1, y1] = transformPoint(args[0], args[1]);
        segs.push({ x0, y0, x1, y1 });
        [cx, cy] = [args[0], args[1]];
        break;
      }
      case OPS.rectangle: {
        const [rx, ry, rw, rh] = args;
        const [ax, ay] = transformPoint(rx, ry);
        const [bx, by] = transformPoint(rx + rw, ry);
        const [dx, dy] = transformPoint(rx + rw, ry + rh);
        const [ex, ey] = transformPoint(rx, ry + rh);
        segs.push({ x0: ax, y0: ay, x1: bx, y1: by });
        segs.push({ x0: bx, y0: by, x1: dx, y1: dy });
        segs.push({ x0: dx, y0: dy, x1: ex, y1: ey });
        segs.push({ x0: ex, y0: ey, x1: ax, y1: ay });
        break;
      }
      case OPS.closePath: {
        const [x0, y0] = transformPoint(cx, cy);
        const [x1, y1] = transformPoint(mx, my);
        if (Math.abs(x0 - x1) > 0.5 || Math.abs(y0 - y1) > 0.5)
          segs.push({ x0, y0, x1, y1 });
        [cx, cy] = [mx, my];
        break;
      }
    }
  }
  return segs;
}

// ══════════════════════════════════════════════════════════════
//  Table detection & HTML build
// ══════════════════════════════════════════════════════════════

// Tolerance in PDF points (~1mm) for coordinate clustering and snapping
const EPS = 3;

function near(a, b) { return Math.abs(a - b) <= EPS; }

function clusterValues(vals, eps) {
  if (!vals.length) return [];
  vals = [...vals].sort((a, b) => a - b);
  const clusters = [vals[0]];
  for (let i = 1; i < vals.length; i++) {
    if (vals[i] - clusters[clusters.length - 1] > eps)
      clusters.push(vals[i]);
    else
      clusters[clusters.length - 1] =
        (clusters[clusters.length - 1] + vals[i]) / 2;   // average
  }
  return clusters;
}

function snapTo(val, grid) {
  let best = grid[0], bestD = Math.abs(val - grid[0]);
  for (let i = 1; i < grid.length; i++) {
    const d = Math.abs(val - grid[i]);
    if (d < bestD) { best = grid[i]; bestD = d; }
  }
  return best;
}

/** Detect tables from line segments, return { tables, usedTextIndices } */
function detectTables(segments, textItems) {
  const hSegs = [];
  const vSegs = [];

  for (const s of segments) {
    if (near(s.y0, s.y1)) {                               // horizontal
      hSegs.push({ a: Math.min(s.x0, s.x1), b: Math.max(s.x0, s.x1), v: (s.y0 + s.y1) / 2 });
    } else if (near(s.x0, s.x1)) {                        // vertical
      vSegs.push({ a: Math.min(s.y0, s.y1), b: Math.max(s.y0, s.y1), v: (s.x0 + s.x1) / 2 });
    }
  }

  if (hSegs.length < 2 || vSegs.length < 2) return { tables: [], usedTextIndices: new Set() };

  // Build grids
  const xVals = [];
  for (const v of vSegs) xVals.push(v.v);
  const yVals = [];
  for (const h of hSegs) yVals.push(h.v);

  const xGrid = clusterValues(xVals, EPS);
  const yGrid = clusterValues(yVals, EPS);

  if (xGrid.length < 2 || yGrid.length < 2) return { tables: [], usedTextIndices: new Set() };

  // For each potential cell, check that bounding edges exist
  // We consider a "table region" as a connected rectangular block of cells
  // that have at least their four edges present in the segments.

  // Build edge look-ups
  function hasHEdge(x0, x1, y) {
    return hSegs.some(s => near(s.v, y) && s.a <= x0 + EPS && s.b >= x1 - EPS);
  }
  function hasVEdge(y0, y1, x) {
    return vSegs.some(s => near(s.v, x) && s.a <= y0 + EPS && s.b >= y1 - EPS);
  }

  // Build cell presence matrix
  const rows = yGrid.length - 1;
  const cols = xGrid.length - 1;
  const cellPresent = Array.from({ length: rows }, () => Array(cols).fill(false));

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      const x0 = xGrid[c], x1 = xGrid[c + 1];
      const y0 = yGrid[r], y1 = yGrid[r + 1];
      if (hasHEdge(x0, x1, y0) && hasHEdge(x0, x1, y1) &&
          hasVEdge(y0, y1, x0) && hasVEdge(y0, y1, x1)) {
        cellPresent[r][c] = true;
      }
    }
  }

  // Find connected table regions via flood fill
  const visited = Array.from({ length: rows }, () => Array(cols).fill(false));
  const tableRegions = [];

  for (let r = 0; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      if (cellPresent[r][c] && !visited[r][c]) {
        const cells = [];
        const queue = [[r, c]];
        visited[r][c] = true;
        while (queue.length) {
          const [cr, cc] = queue.shift();
          cells.push([cr, cc]);
          for (const [dr, dc] of [[0,1],[0,-1],[1,0],[-1,0]]) {
            const nr = cr + dr, nc = cc + dc;
            if (nr >= 0 && nr < rows && nc >= 0 && nc < cols &&
                cellPresent[nr][nc] && !visited[nr][nc]) {
              visited[nr][nc] = true;
              queue.push([nr, nc]);
            }
          }
        }
        if (cells.length >= 2) tableRegions.push(cells);     // at least 2 cells = table
      }
    }
  }

  // Build table objects
  const usedTextIndices = new Set();
  const tables = [];

  for (const cells of tableRegions) {
    const rMin = Math.min(...cells.map(c => c[0]));
    const rMax = Math.max(...cells.map(c => c[0]));
    const cMin = Math.min(...cells.map(c => c[1]));
    const cMax = Math.max(...cells.map(c => c[1]));

    const tRows = rMax - rMin + 1;
    const tCols = cMax - cMin + 1;
    const grid  = Array.from({ length: tRows }, () =>
      Array.from({ length: tCols }, () => ({ texts: [] }))
    );

    const tTop    = yGrid[rMin];
    const tBottom = yGrid[rMax + 1];
    const tLeft   = xGrid[cMin];
    const tRight  = xGrid[cMax + 1];

    // Assign text items to cells
    textItems.forEach((ti, idx) => {
      if (usedTextIndices.has(idx)) return;
      const midX = ti.x + ti.w / 2;
      const midY = ti.y + ti.h / 2;
      if (midX < tLeft - EPS || midX > tRight + EPS ||
          midY < tTop  - EPS || midY > tBottom + EPS) return;

      // Find which cell
      for (let r = rMin; r <= rMax; r++) {
        for (let c = cMin; c <= cMax; c++) {
          const cx0 = xGrid[c], cx1 = xGrid[c + 1];
          const cy0 = yGrid[r], cy1 = yGrid[r + 1];
          if (midX >= cx0 - EPS && midX <= cx1 + EPS &&
              midY >= cy0 - EPS && midY <= cy1 + EPS) {
            grid[r - rMin][c - cMin].texts.push(ti);
            usedTextIndices.add(idx);
            return;
          }
        }
      }
    });

    tables.push({ top: tTop, left: tLeft, bottom: tBottom, right: tRight, grid, tRows, tCols });
  }

  return { tables, usedTextIndices };
}

// ── Number format detection (3.426,64 style) ──────────────────
const EUROPEAN_NUMBER_REGEX = /^-?\d{1,3}(?:\.\d{3})*(?:,\d+)?$/;
function isEuropeanNumber(s) { return EUROPEAN_NUMBER_REGEX.test(s.trim()); }

// ── Build HTML for one page ──────────────────────────────────
function buildPageHtml(page) {
  const { textItems, segments } = page;
  const { tables, usedTextIndices } = detectTables(segments, textItems);

  // Sort tables top‑to‑bottom
  tables.sort((a, b) => a.top - b.top);

  // Collect "free" text items (not in any table)
  const freeItems = textItems.filter((_, i) => !usedTextIndices.has(i));

  // Build a combined list of blocks (tables + text paragraphs) sorted by Y
  const blocks = [];

  for (const tbl of tables) {
    blocks.push({ type: 'table', y: tbl.top, data: tbl });
  }

  // Group free text into lines by similar Y
  const lines = groupIntoLines(freeItems);
  for (const line of lines) {
    blocks.push({ type: 'text', y: line.y, data: line });
  }

  blocks.sort((a, b) => a.y - b.y);

  // Render
  let html = '';
  for (const block of blocks) {
    if (block.type === 'table') {
      html += renderTable(block.data);
    } else {
      html += renderTextLine(block.data);
    }
  }

  return html;
}

function groupIntoLines(items) {
  if (!items.length) return [];
  const sorted = [...items].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines  = [];
  let   curLine = { y: sorted[0].y, items: [sorted[0]] };

  for (let i = 1; i < sorted.length; i++) {
    if (Math.abs(sorted[i].y - curLine.y) <= EPS) {
      curLine.items.push(sorted[i]);
    } else {
      lines.push(curLine);
      curLine = { y: sorted[i].y, items: [sorted[i]] };
    }
  }
  lines.push(curLine);
  return lines;
}

const LINE_GAP = 2;

function buildLineText(items) {
  if (!items.length) return '';
  const sorted = [...items].sort((a, b) => a.x - b.x);
  let text = '';
  let prev = null;

  for (const item of sorted) {
    const value = normalizePdfText(item.str);
    if (!value) {
      continue;
    }
    if (text) {
      const gap = prev ? item.x - (prev.x + prev.w) : LINE_GAP + 1;
      if (gap > LINE_GAP) {
        text += ' ';
      }
    }
    text += value;
    prev = item;
  }
  return text.replace(/\s+/g, ' ').trim();
}

function normalizePdfText(str) {
  return String(str || '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function renderTextLine(line) {
  const text = buildLineText(line.items);
  if (!text) return '';
  return `<p style="font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:2px 0;">${escapeHtml(text)}</p>\n`;
}

function renderTable(tbl) {
  let html = '<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:8px 0;">\n';
  for (let r = 0; r < tbl.tRows; r++) {
    html += '  <tr>\n';
    for (let c = 0; c < tbl.tCols; c++) {
      const cell = tbl.grid[r][c];
      const text = buildLineText(cell.texts);
      const escapedText = escapeHtml(text);

      let align = 'left';
      let noWrap = '';
      if (isEuropeanNumber(text)) {
        align  = 'right';
        noWrap = 'white-space:nowrap;';
      }

      html += `    <td style="border:1px solid #000;padding:4px 6px;text-align:${align};${noWrap}">${escapedText}</td>\n`;
    }
    html += '  </tr>\n';
  }
  html += '</table>\n';
  return html;
}

// ══════════════════════════════════════════════════════════════
//  Recipient extraction
// ══════════════════════════════════════════════════════════════

const NIF_CODE_REGEX = /\b([A-Z]\d{7,8})\b/;

function extractRecipient(page) {
  if (!page) return '';
  const { textItems, height } = page;
  const sorted = [...textItems].sort((a, b) => a.y - b.y || a.x - b.x);
  const lines = groupIntoLines(sorted);
  const upperLimit = height * RECIPIENT_SEARCH_HEIGHT_RATIO;

  for (const line of lines) {
    if (line.y > upperLimit) continue;
    const lineStr = buildLineText(line.items);
    const m = lineStr.match(NIF_CODE_REGEX);
    if (m) {
      const afterNif = lineStr.substring(m.index + m[1].length).trim();
      if (afterNif) return afterNif;
      return '';
    }
  }
  return '';
}

const EMAIL_REGEX = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
const EMAIL_BLACKLIST = new Set(['conta@nmeconomista.com']);

function extractToEmail(page) {
  if (!page) return '';
  const lines = groupIntoLines(page.textItems).sort((a, b) => a.y - b.y || a.x - b.x);

  for (const line of lines) {
    const text = buildLineText(line.items);
    const anchorMatch = text.match(/^A\s+(.+)/i);
    if (anchorMatch) {
      const compact = anchorMatch[1].replace(/\s+/g, '');
      const anchorEmail = findFirstEmail(compact);
      if (anchorEmail) return anchorEmail;
    }
  }

  const candidates = [];
  for (const line of lines) {
    const text = buildLineText(line.items);
    if (!text) continue;
    collectEmails(text, line.y, candidates);
    const compact = text.replace(/\s+/g, '');
    if (compact !== text) {
      collectEmails(compact, line.y, candidates);
    }
  }

  if (!candidates.length) return '';
  candidates.sort((a, b) => a.y - b.y);
  const topY = candidates[0].y;
  const topCandidates = candidates.filter(c => Math.abs(c.y - topY) <= EPS);
  const uniqueEmails = Array.from(new Set(topCandidates.map(c => c.email)));
  return uniqueEmails.length === 1 ? uniqueEmails[0] : '';
}

function extractSubject(page) {
  if (!page) return DEFAULT_SUBJECT;
  const lines = groupIntoLines(page.textItems).sort((a, b) => a.y - b.y || a.x - b.x);
  for (let i = 0; i < lines.length; i++) {
    const text = buildLineText(lines[i].items);
    if (SUBJECT_ANCHOR_REGEX.test(text)) {
      for (let j = i - 1; j >= 0; j--) {
        const signer = buildLineText(lines[j].items);
        if (signer) {
          return `${DEFAULT_SUBJECT} ${signer}`;
        }
      }
      break;
    }
  }
  return DEFAULT_SUBJECT;
}

function collectEmails(text, y, list) {
  const matches = text.matchAll(new RegExp(EMAIL_REGEX.source, 'ig'));
  for (const match of matches) {
    const raw = match[0];
    const email = raw.toLowerCase();
    if (EMAIL_BLACKLIST.has(email)) continue;
    list.push({ email, y });
  }
}

function findFirstEmail(text) {
  const match = text.match(EMAIL_REGEX);
  if (!match) return '';
  const email = match[0].toLowerCase();
  return EMAIL_BLACKLIST.has(email) ? '' : email;
}

// ══════════════════════════════════════════════════════════════
//  Filename sanitisation
// ══════════════════════════════════════════════════════════════

function sanitizeFileName(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim()
    .substring(0, 120);
}

// ══════════════════════════════════════════════════════════════
//  EML generation (RFC 2822 + HTML body)
// ══════════════════════════════════════════════════════════════

function buildEml(bodyHtml, { toEmail, subject }) {
  const date     = new Date().toUTCString();

  const fullHtml = `<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">
${bodyHtml}
</body></html>`;

  const lines = [
    'MIME-Version: 1.0',
    `Date: ${date}`,
    toEmail ? `To: ${toEmail}` : null,
    `Subject: ${rfc2047EncodeHeaderValue(subject)}`,
    'Content-Type: text/html; charset=UTF-8',
    'Content-Transfer-Encoding: base64',
    '',
    emlBodyBase64EncodeWithWrap(fullHtml),
    ''
  ].filter(Boolean);

  return lines.join('\r\n');
}

function rfc2047EncodeHeaderValue(value) {
  const bytes = new TextEncoder().encode(String(value || ''));
  const base64 = base64FromBytes(bytes);
  return `=?UTF-8?B?${base64}?=`;
}

function emlBodyBase64EncodeWithWrap(html) {
  const bytes = new TextEncoder().encode(String(html || ''));
  const base64 = base64FromBytes(bytes);
  return base64.match(/.{1,76}/g)?.join('\r\n') || '';
}

function base64FromBytes(bytes) {
  let binary = '';
  for (const b of bytes) {
    binary += String.fromCharCode(b);
  }
  return btoa(binary);
}

// ── Download helper ───────────────────────────────────────────
function downloadBlob(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const a    = document.createElement('a');
  const url  = URL.createObjectURL(blob);
  a.href     = url;
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(url); }, 100);
}

// ── HTML escape ───────────────────────────────────────────────
function escapeHtml(str) {
  return String(str || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}

function runSelfChecks() {
  console.assert(sanitizeFileName('Name:Test/Doc') === 'Name_Test_Doc', 'sanitizeFileName failed');
  const encodedSubject = rfc2047EncodeHeaderValue('Привет');
  console.assert(encodedSubject.startsWith('=?UTF-8?B?') && encodedSubject.endsWith('?='), 'rfc2047EncodeHeaderValue failed');
  const wrapped = emlBodyBase64EncodeWithWrap('á'.repeat(120));
  const lines = wrapped.split('\r\n');
  console.assert(lines.every(line => line.length <= 76), 'emlBodyBase64EncodeWithWrap failed');
}

if (new URLSearchParams(window.location.search).has('dev')) {
  runSelfChecks();
}
