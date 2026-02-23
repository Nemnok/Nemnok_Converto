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
const statusEl       = document.getElementById('status');

// ── State ─────────────────────────────────────────────────────
let selectedFile     = null;
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
  selectedFile = pdfInput.files[0] || null;
  generateBtn.disabled = !selectedFile;
});

generateBtn.addEventListener('click', async () => {
  if (!selectedFile) return;
  generateBtn.disabled = true;
  setStatus('Processing PDF…');
  try {
    const arrayBuf = await selectedFile.arrayBuffer();
    const pdf      = await pdfjsLib.getDocument({ data: arrayBuf }).promise;

    // 1. Extract pages
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i++) {
      const page = await pdf.getPage(i);
      pages.push(await extractPage(page));
    }

    // 2. Recipient from first page
    const recipient = extractRecipient(pages[0].textItems);

    // 3. Build HTML body
    let bodyHtml = '';
    for (const pg of pages) {
      bodyHtml += buildPageHtml(pg);
    }

    // 4. Append signature
    const sigId = sigSelect.value;
    if (sigId && signaturesCache[sigId]) {
      bodyHtml += '<br>' + signaturesCache[sigId];
    }

    // 5. Build EML
    const eml = buildEml(bodyHtml);

    // 6. Download
    const filename = sanitizeFilename(recipient) + '.eml';
    downloadBlob(eml, filename, 'message/rfc822');

    setStatus('Done – downloaded ' + filename);
  } catch (err) {
    console.error(err);
    setStatus('Error: ' + err.message, true);
  } finally {
    generateBtn.disabled = false;
  }
});

// ── Helpers ───────────────────────────────────────────────────
function setStatus(msg, isError) {
  statusEl.textContent = msg;
  statusEl.className   = isError ? 'error' : '';
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

function renderTextLine(line) {
  const parts = line.items
    .sort((a, b) => a.x - b.x)
    .map(it => escapeHtml(it.str))
    .join(' ');
  if (!parts.trim()) return '';
  return `<p style="font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:2px 0;">${parts}</p>\n`;
}

function renderTable(tbl) {
  let html = '<table border="1" cellspacing="0" cellpadding="4" style="border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:11pt;margin:8px 0;">\n';
  for (let r = 0; r < tbl.tRows; r++) {
    html += '  <tr>\n';
    for (let c = 0; c < tbl.tCols; c++) {
      const cell = tbl.grid[r][c];
      const text = cell.texts
        .sort((a, b) => a.y - b.y || a.x - b.x)
        .map(t => escapeHtml(t.str))
        .join(' ')
        .trim();

      let align = 'left';
      let noWrap = '';
      if (isEuropeanNumber(text)) {
        align  = 'right';
        noWrap = 'white-space:nowrap;';
      }

      html += `    <td style="border:1px solid #000;padding:4px 6px;text-align:${align};${noWrap}">${text}</td>\n`;
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

function extractRecipient(textItems) {
  // Sort by Y (top first), then X for items on same line
  const sorted = [...textItems].sort((a, b) => a.y - b.y || a.x - b.x);

  // Group into lines
  const lines = groupIntoLines(sorted);

  for (const line of lines) {
    const lineStr = line.items
      .sort((a, b) => a.x - b.x)
      .map(it => it.str)
      .join(' ');
    const m = lineStr.match(NIF_CODE_REGEX);
    if (m) {
      const afterNif = lineStr.substring(m.index + m[1].length).trim();
      if (afterNif) return afterNif;
      return m[1];    // fallback: use NIF itself
    }
  }
  return 'output';    // fallback
}

// ══════════════════════════════════════════════════════════════
//  Filename sanitisation
// ══════════════════════════════════════════════════════════════

function sanitizeFilename(name) {
  return name
    .replace(/[\\/:*?"<>|]/g, '_')
    .trim()
    .substring(0, 120);
}

// ══════════════════════════════════════════════════════════════
//  EML generation (RFC 2822 + HTML body)
// ══════════════════════════════════════════════════════════════

function buildEml(bodyHtml) {
  const boundary = '----=_Part_' + crypto.randomUUID();
  const date     = new Date().toUTCString();

  const fullHtml = `<!DOCTYPE html>
<html><head><meta charset="utf-8"></head>
<body style="font-family:Calibri,Arial,sans-serif;font-size:11pt;">
${bodyHtml}
</body></html>`;

  const lines = [
    'MIME-Version: 1.0',
    `Date: ${date}`,
    'Subject: ',
    'Content-Type: multipart/alternative; boundary="' + boundary + '"',
    '',
    '--' + boundary,
    'Content-Type: text/html; charset=utf-8',
    'Content-Transfer-Encoding: quoted-printable',
    '',
    quotedPrintableEncode(fullHtml),
    '',
    '--' + boundary + '--',
    ''
  ];

  return lines.join('\r\n');
}

/** Very minimal quoted-printable encoder */
function quotedPrintableEncode(str) {
  const encoded = [];
  const rawLines = str.split('\n');
  for (const rawLine of rawLines) {
    let line = '';
    for (let i = 0; i < rawLine.length; i++) {
      const c = rawLine.charCodeAt(i);
      if (c === 9 || (c >= 32 && c <= 126 && c !== 61)) {
        line += rawLine[i];
      } else {
        // encode as =XX (multi-byte for non-ASCII)
        const bytes = new TextEncoder().encode(rawLine[i]);
        for (const b of bytes) {
          line += '=' + b.toString(16).toUpperCase().padStart(2, '0');
        }
      }
      // soft line break at 75 chars
      if (line.length >= 73) {
        encoded.push(line + '=');
        line = '';
      }
    }
    encoded.push(line);
  }
  return encoded.join('\r\n');
}

// ── Download helper ───────────────────────────────────────────
function downloadBlob(text, filename, mime) {
  const blob = new Blob([text], { type: mime });
  const a    = document.createElement('a');
  a.href     = URL.createObjectURL(blob);
  a.download = filename;
  document.body.appendChild(a);
  a.click();
  setTimeout(() => { document.body.removeChild(a); URL.revokeObjectURL(a.href); }, 100);
}

// ── HTML escape ───────────────────────────────────────────────
function escapeHtml(str) {
  return str
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;');
}
