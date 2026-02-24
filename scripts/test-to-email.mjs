/**
 * Test harness for extractToEmail() – verifies that the "A …" recipient line
 * extraction works correctly for the sample PDFs.
 *
 * Usage:  node scripts/test-to-email.mjs
 *
 * Requires:  npm install  (installs pdfjs-dist)
 */
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';

// ── Shared constants / helpers (mirrored from app.js) ─────────
const EPS = 3;
const EMAIL_REGEX = /[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i;
const EMAIL_BLACKLIST = new Set(['conta@nmeconomista.com']);

function normalizePdfText(str) {
  return String(str || '')
    .replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function stripInvisibleChars(str) {
  return String(str || '').replace(/[\u200B\u200C\u200D\uFEFF\u00AD\u00A0]/g, '');
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
    if (!value) continue;
    if (text) {
      const gap = prev ? item.x - (prev.x + prev.w) : LINE_GAP + 1;
      if (gap > LINE_GAP) text += ' ';
    }
    text += value;
    prev = item;
  }
  return text.replace(/\s+/g, ' ').trim();
}

function findFirstEmail(text) {
  const match = text.match(EMAIL_REGEX);
  if (!match) return '';
  const email = match[0].toLowerCase();
  return EMAIL_BLACKLIST.has(email) ? '' : email;
}

// ── extractPage (same as app.js) ──────────────────────────────
async function extractPage(page) {
  const viewport = page.getViewport({ scale: 1 });
  const height   = viewport.height;

  const tc    = await page.getTextContent();
  const textItems = tc.items.filter(it => it.str != null).map(it => {
    const tx = it.transform;
    const x  = tx[4];
    const y  = height - tx[5];
    const w  = it.width  || 0;
    const h  = Math.abs(tx[3]) || it.height || 10;
    return { str: it.str, x, y: y - h, w, h, fontName: it.fontName };
  });

  const annotations = await page.getAnnotations();

  return { textItems, annotations, width: viewport.width, height };
}

// ── extractToEmail (same as app.js) ───────────────────────────
function extractToEmail(page) {
  if (!page) return '';
  const lines = groupIntoLines(page.textItems).sort((a, b) => a.y - b.y || a.x - b.x);

  for (let i = 0; i < lines.length; i++) {
    const sortedItems = [...lines[i].items].sort((a, b) => a.x - b.x);
    const firstIdx = sortedItems.findIndex(it => normalizePdfText(it.str) !== '');
    if (firstIdx === -1) continue;
    const firstToken = normalizePdfText(sortedItems[firstIdx].str);

    let candidate = '';

    if (/^A:?$/i.test(firstToken)) {
      // Case 1: "A" is a separate text item
      const itemsAfterA = sortedItems.slice(firstIdx + 1);
      candidate = stripInvisibleChars(buildLineText(itemsAfterA)).replace(/\s+/g, '');
    } else if (/^A:?\s/i.test(firstToken)) {
      // Case 2: "A" and email are in the same text item
      const afterA = firstToken.replace(/^A:?\s*/i, '');
      const itemsAfterA = sortedItems.slice(firstIdx + 1);
      const rest = buildLineText(itemsAfterA);
      candidate = stripInvisibleChars(afterA + (rest ? ' ' + rest : '')).replace(/\s+/g, '');
    } else {
      continue;
    }

    // If email not found, append next line's text as continuation
    if (!findFirstEmail(candidate) && i + 1 < lines.length) {
      candidate += stripInvisibleChars(buildLineText(lines[i + 1].items)).replace(/\s+/g, '');
    }

    // If still no email in text, check mailto: annotations near the A-line
    if (!findFirstEmail(candidate) && page.annotations) {
      const aLineY = lines[i].y;
      const yTolerance = 20;
      for (const ann of page.annotations) {
        const url = ann.url || ann.unsafeUrl || '';
        if (!url.startsWith('mailto:')) continue;
        const annY = page.height - ann.rect[3];
        if (Math.abs(annY - aLineY) <= yTolerance) {
          const mailtoEmail = url.replace(/^mailto:/i, '').split('?')[0].toLowerCase();
          if (mailtoEmail && !EMAIL_BLACKLIST.has(mailtoEmail)) return mailtoEmail;
        }
      }
    }

    const email = findFirstEmail(candidate);
    return email;
  }
  return '';
}

// ── Test runner ───────────────────────────────────────────────
const TEST_CASES = [
  { path: 'test/2.pdf', expected: 'club@leroymerlin.es' },
  { path: 'test/3.pdf', expected: 'golub.leonid111@gmail.com' },
];

let passed = 0;
let failed = 0;

for (const { path, expected } of TEST_CASES) {
  console.log(`\n── ${path} (expected: ${expected}) ──`);

  const pdf  = await getDocument(path).promise;
  const page = await pdf.getPage(1);
  const extracted = await extractPage(page);

  // Diagnostics
  const atItems = extracted.textItems.filter(it => it.str.includes('@'));
  console.log(`  textItems with "@": ${atItems.length}`);
  for (const it of atItems) {
    console.log(`    str="${it.str}" x=${it.x.toFixed(1)} y=${it.y.toFixed(1)}`);
  }

  const lines = groupIntoLines(extracted.textItems).sort((a, b) => a.y - b.y);
  for (const line of lines) {
    const sorted = [...line.items].sort((a, b) => a.x - b.x);
    const firstIdx = sorted.findIndex(it => normalizePdfText(it.str) !== '');
    if (firstIdx === -1) continue;
    const first = normalizePdfText(sorted[firstIdx].str);
    if (/^A:?\s/i.test(first) || /^A:?$/i.test(first)) {
      console.log(`  A-line found at y=${line.y.toFixed(1)}: "${first}"`);
      // Show surrounding items
      for (const it of sorted) {
        console.log(`    item: str="${it.str}" x=${it.x.toFixed(1)} y=${it.y.toFixed(1)} w=${it.w.toFixed(1)}`);
      }
    }
  }

  if (extracted.annotations && extracted.annotations.length) {
    console.log(`  Annotations: ${extracted.annotations.length}`);
    for (const ann of extracted.annotations) {
      if (ann.url || ann.unsafeUrl) {
        console.log(`    type=${ann.subtype} url="${ann.url || ann.unsafeUrl}"`);
      }
    }
  }

  // Run extraction
  const result = extractToEmail(extracted);
  console.log(`  extractToEmail => "${result}"`);

  if (result === expected) {
    console.log(`  ✅ PASS`);
    passed++;
  } else {
    console.log(`  ❌ FAIL (expected "${expected}", got "${result}")`);
    failed++;
  }
}

console.log(`\n══ Results: ${passed} passed, ${failed} failed ══`);
process.exit(failed > 0 ? 1 : 0);
